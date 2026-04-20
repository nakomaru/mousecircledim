"""MouseCircleDim: cursor highlighter + global screen dimmer in one tray app.

Combines MouseCircle (click-through cursor ring with click ripples) and
VirtualDim (Magnification-API fullscreen color scaling). Single tray icon
exposes dim levels, a mouse-circle toggle, and quit.
"""

# --- venv bootstrap -------------------------------------------------------
import os
import sys
from pathlib import Path

VENV_DIR = Path.home() / "venvs" / ".venv_mousecircledim"
REQUIREMENTS = ["pywin32", "Pillow", "pystray"]


def _bootstrap():
    if os.name != "nt":
        sys.exit("MouseCircleDim only runs on Windows.")
    try:
        in_venv = Path(sys.prefix).resolve() == VENV_DIR.resolve()
    except OSError:
        in_venv = False
    if in_venv:
        return
    py = VENV_DIR / "Scripts" / "python.exe"
    pyw = VENV_DIR / "Scripts" / "pythonw.exe"
    if not py.exists():
        import venv
        import subprocess
        print(f"Creating venv at {VENV_DIR}...", flush=True)
        VENV_DIR.parent.mkdir(parents=True, exist_ok=True)
        venv.create(VENV_DIR, with_pip=True)
        subprocess.check_call([
            str(py), "-m", "pip", "install", "--quiet",
            "--disable-pip-version-check", *REQUIREMENTS,
        ])
    target = pyw if pyw.exists() else py
    os.execv(str(target), [str(target), os.path.abspath(__file__), *sys.argv[1:]])


_bootstrap()

# --- real imports ---------------------------------------------------------
import configparser
import ctypes
import subprocess
import threading
import time
import tkinter as tk
from ctypes import (
    POINTER, Structure, WINFUNCTYPE, byref, c_float, c_int, c_ubyte, c_void_p, wintypes,
)

import win32api
import win32con
import win32ts
from PIL import Image, ImageChops, ImageDraw, ImageTk
import pystray

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
kernel32 = ctypes.windll.kernel32
wtsapi32 = ctypes.windll.wtsapi32

# --- DPI awareness (must run before Magnification init) -------------------
try:
    user32.SetProcessDpiAwarenessContext.argtypes = [c_void_p]
    user32.SetProcessDpiAwarenessContext.restype = c_int
    user32.SetProcessDpiAwarenessContext(c_void_p(-4))  # per-monitor v2
except (AttributeError, OSError):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        user32.SetProcessDPIAware()

magnification = ctypes.windll.Magnification


# --- Win32 constants ------------------------------------------------------
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000
WS_EX_TOPMOST = 0x00000008
WS_POPUP = 0x80000000
ULW_ALPHA = 0x00000002
AC_SRC_OVER = 0
AC_SRC_ALPHA = 1
SW_HIDE = 0
SW_SHOWNOACTIVATE = 4
HWND_TOPMOST = -1
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
WM_DESTROY = 0x0002
WM_WTSSESSION_CHANGE = 0x02B1
WTS_SESSION_LOCK = 0x7
WTS_SESSION_UNLOCK = 0x8
NOTIFY_FOR_THIS_SESSION = 0
BI_RGB = 0
DIB_RGB_COLORS = 0
PM_REMOVE = 0x0001
WM_QUIT = 0x0012
WM_USER = 0x0400
WM_APP_EXIT = WM_USER + 1


# --- ctypes structs -------------------------------------------------------
class BITMAPINFOHEADER(Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


class BLENDFUNCTION(Structure):
    _fields_ = [
        ("BlendOp", c_ubyte), ("BlendFlags", c_ubyte),
        ("SourceConstantAlpha", c_ubyte), ("AlphaFormat", c_ubyte),
    ]


class POINT(Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class SIZE(Structure):
    _fields_ = [("cx", wintypes.LONG), ("cy", wintypes.LONG)]


class MSG(Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND), ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM), ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD), ("pt", POINT),
    ]


class WNDCLASS(Structure):
    _fields_ = [
        ("style", wintypes.UINT), ("lpfnWndProc", c_void_p),
        ("cbClsExtra", c_int), ("cbWndExtra", c_int),
        ("hInstance", wintypes.HINSTANCE), ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HANDLE), ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR), ("lpszClassName", wintypes.LPCWSTR),
    ]


class MAGCOLOREFFECT(Structure):
    _fields_ = [("transform", (c_float * 5) * 5)]


WNDPROC = WINFUNCTYPE(c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

user32.DefWindowProcW.restype = c_void_p
user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.CreateWindowExW.restype = wintypes.HWND
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    c_int, c_int, c_int, c_int,
    wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID,
]
user32.RegisterClassW.argtypes = [POINTER(WNDCLASS)]
user32.UpdateLayeredWindow.argtypes = [
    wintypes.HWND, wintypes.HDC, POINTER(POINT), POINTER(SIZE),
    wintypes.HDC, POINTER(POINT), wintypes.DWORD,
    POINTER(BLENDFUNCTION), wintypes.DWORD,
]
user32.UpdateLayeredWindow.restype = wintypes.BOOL
user32.GetDC.argtypes = [wintypes.HWND]
user32.GetDC.restype = wintypes.HDC
user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, c_int, c_int, c_int, c_int, wintypes.UINT]
gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC
gdi32.CreateDIBSection.argtypes = [
    wintypes.HDC, POINTER(BITMAPINFO), wintypes.UINT,
    POINTER(c_void_p), wintypes.HANDLE, wintypes.DWORD,
]
gdi32.CreateDIBSection.restype = wintypes.HBITMAP
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
wtsapi32.WTSRegisterSessionNotification.argtypes = [wintypes.HWND, wintypes.DWORD]
wtsapi32.WTSUnRegisterSessionNotification.argtypes = [wintypes.HWND]

magnification.MagInitialize.restype = wintypes.BOOL
magnification.MagUninitialize.restype = wintypes.BOOL
magnification.MagSetFullscreenColorEffect.argtypes = [POINTER(MAGCOLOREFFECT)]
magnification.MagSetFullscreenColorEffect.restype = wintypes.BOOL


# === USER SETTINGS ========================================================
# Defaults; overridden by INI at INI_PATH (same folder as this script).
INI_PATH = Path(__file__).resolve().parent / "mousecircledim.ini"

_DEFAULTS = {
    "circle": {
        "IDLE_COLOR":         "0.90, 0.95, 1.00",
        "LEFT_CLICK_COLOR":   "0.50, 0.85, 0.30",
        "RIGHT_CLICK_COLOR":  "0.85, 0.30, 0.20",
        "MIDDLE_CLICK_COLOR": "0.30, 0.50, 0.85",
        "RING_RADIUS":       "41",
        "RING_THICKNESS":    "6",
        "RING_ALPHA_CENTER": "0.50",
        "RING_ALPHA_EDGE":   "0.40",
        "HALO_ENABLED":          "true",
        "HALO_THICKNESS_FACTOR": "2",
        "HALO_COLOR_SCALE":      "0.2",
        "HALO_ALPHA_INNER":      "0.25",
        "HALO_ALPHA_OUTER":      "0.05",
        "RIPPLE_THICKNESS": "3",
        "RIPPLE_GROWTH":    "3.0",
        "RIPPLE_DURATION":  "1.0",
        "TARGET_FPS": "120",
        "PADDING":    "4",
        "ENABLED":    "true",
    },
    "dim": {
        "LEVELS": "90, 85, 80, 70, 60, 50, 40, 30, 20, 10, 0",
        "LAST_LEVEL": "0",
        "LAST_NONZERO": "40",
    },
}

_cfg = configparser.ConfigParser()
for sec, opts in _DEFAULTS.items():
    _cfg[sec] = dict(opts)
if INI_PATH.exists():
    try:
        _cfg.read(INI_PATH, encoding="utf-8")
    except Exception as e:
        sys.stderr.write(f"failed to read {INI_PATH}: {e}\n")


def _tuple(s):
    return tuple(float(x.strip()) for x in s.split(","))


def _int_list(s):
    return [int(x.strip()) for x in s.split(",") if x.strip()]


_c = _cfg["circle"]
IDLE_COLOR         = _tuple(_c["IDLE_COLOR"])
LEFT_CLICK_COLOR   = _tuple(_c["LEFT_CLICK_COLOR"])
RIGHT_CLICK_COLOR  = _tuple(_c["RIGHT_CLICK_COLOR"])
MIDDLE_CLICK_COLOR = _tuple(_c["MIDDLE_CLICK_COLOR"])

RING_RADIUS       = _c.getint("RING_RADIUS")
RING_THICKNESS    = _c.getint("RING_THICKNESS")
RING_ALPHA_CENTER = _c.getfloat("RING_ALPHA_CENTER")
RING_ALPHA_EDGE   = _c.getfloat("RING_ALPHA_EDGE")

HALO_ENABLED          = _c.getboolean("HALO_ENABLED")
HALO_THICKNESS_FACTOR = _c.getint("HALO_THICKNESS_FACTOR")
HALO_COLOR_SCALE      = _c.getfloat("HALO_COLOR_SCALE")
HALO_ALPHA_INNER      = _c.getfloat("HALO_ALPHA_INNER")
HALO_ALPHA_OUTER      = _c.getfloat("HALO_ALPHA_OUTER")

RIPPLE_THICKNESS = _c.getint("RIPPLE_THICKNESS")
RIPPLE_GROWTH    = _c.getfloat("RIPPLE_GROWTH")
RIPPLE_DURATION  = _c.getfloat("RIPPLE_DURATION")

TARGET_FPS = _c.getint("TARGET_FPS")
PADDING    = _c.getint("PADDING")
CIRCLE_ENABLED = _c.getboolean("ENABLED")

LEVELS = _int_list(_cfg["dim"]["LEVELS"])
INITIAL_LEVEL = _cfg["dim"].getint("LAST_LEVEL")
INITIAL_NONZERO = _cfg["dim"].getint("LAST_NONZERO")


def _save_ini():
    try:
        with open(INI_PATH, "w", encoding="utf-8") as f:
            _cfg.write(f)
    except Exception as e:
        sys.stderr.write(f"failed to write {INI_PATH}: {e}\n")


# Ensure INI exists on first run so users can find/edit it.
if not INI_PATH.exists():
    _save_ini()


# === DERIVED ==============================================================
def _u8(x):
    return max(0, min(255, int(round(x * 255))))


def _rgb255(c):
    return tuple(_u8(v) for v in c)


def _odd(n):
    n = max(1, int(round(n)))
    return n if n % 2 == 1 else n + 1


def _window_size(ring_radius_max, stroke_thickness):
    halo_t = stroke_thickness * HALO_THICKNESS_FACTOR if HALO_ENABLED else 0
    extent = ring_radius_max + max(stroke_thickness, halo_t) / 2 + PADDING
    return _odd(extent * 2)


_IDLE_RGB = _LEFT_RGB = _RIGHT_RGB = _MIDDLE_RGB = (0, 0, 0)
_RING_ALPHA_CENTER_U8 = _RING_ALPHA_EDGE_U8 = 0
_HALO_ALPHA_INNER_U8 = _HALO_ALPHA_OUTER_U8 = 0
CURSOR_WINDOW_SIZE = 0
RIPPLE_WINDOW_SIZE = 0


def recompute_derived():
    global _IDLE_RGB, _LEFT_RGB, _RIGHT_RGB, _MIDDLE_RGB
    global _RING_ALPHA_CENTER_U8, _RING_ALPHA_EDGE_U8
    global _HALO_ALPHA_INNER_U8, _HALO_ALPHA_OUTER_U8
    global CURSOR_WINDOW_SIZE, RIPPLE_WINDOW_SIZE
    _IDLE_RGB   = _rgb255(IDLE_COLOR)
    _LEFT_RGB   = _rgb255(LEFT_CLICK_COLOR)
    _RIGHT_RGB  = _rgb255(RIGHT_CLICK_COLOR)
    _MIDDLE_RGB = _rgb255(MIDDLE_CLICK_COLOR)
    _RING_ALPHA_CENTER_U8 = _u8(RING_ALPHA_CENTER)
    _RING_ALPHA_EDGE_U8   = _u8(RING_ALPHA_EDGE)
    _HALO_ALPHA_INNER_U8  = _u8(HALO_ALPHA_INNER)
    _HALO_ALPHA_OUTER_U8  = _u8(HALO_ALPHA_OUTER)
    CURSOR_WINDOW_SIZE = _window_size(RING_RADIUS, RING_THICKNESS)
    RIPPLE_WINDOW_SIZE = _window_size(RING_RADIUS * RIPPLE_GROWTH, RIPPLE_THICKNESS)


recompute_derived()


def apply_circle_settings_from_cfg():
    """Re-read [circle] into module globals + recompute derived. Returns
    True if CURSOR_WINDOW_SIZE changed (caller should recreate cursor window)."""
    global IDLE_COLOR, LEFT_CLICK_COLOR, RIGHT_CLICK_COLOR, MIDDLE_CLICK_COLOR
    global RING_RADIUS, RING_THICKNESS, RING_ALPHA_CENTER, RING_ALPHA_EDGE
    global HALO_ENABLED, HALO_THICKNESS_FACTOR, HALO_COLOR_SCALE
    global HALO_ALPHA_INNER, HALO_ALPHA_OUTER
    global RIPPLE_THICKNESS, RIPPLE_GROWTH, RIPPLE_DURATION
    global TARGET_FPS, PADDING
    old_size = CURSOR_WINDOW_SIZE
    c = _cfg["circle"]
    try:
        IDLE_COLOR         = _tuple(c["IDLE_COLOR"])
        LEFT_CLICK_COLOR   = _tuple(c["LEFT_CLICK_COLOR"])
        RIGHT_CLICK_COLOR  = _tuple(c["RIGHT_CLICK_COLOR"])
        MIDDLE_CLICK_COLOR = _tuple(c["MIDDLE_CLICK_COLOR"])
        RING_RADIUS       = c.getint("RING_RADIUS")
        RING_THICKNESS    = c.getint("RING_THICKNESS")
        RING_ALPHA_CENTER = c.getfloat("RING_ALPHA_CENTER")
        RING_ALPHA_EDGE   = c.getfloat("RING_ALPHA_EDGE")
        HALO_ENABLED          = c.getboolean("HALO_ENABLED")
        HALO_THICKNESS_FACTOR = c.getint("HALO_THICKNESS_FACTOR")
        HALO_COLOR_SCALE      = c.getfloat("HALO_COLOR_SCALE")
        HALO_ALPHA_INNER      = c.getfloat("HALO_ALPHA_INNER")
        HALO_ALPHA_OUTER      = c.getfloat("HALO_ALPHA_OUTER")
        RIPPLE_THICKNESS = c.getint("RIPPLE_THICKNESS")
        RIPPLE_GROWTH    = c.getfloat("RIPPLE_GROWTH")
        RIPPLE_DURATION  = c.getfloat("RIPPLE_DURATION")
        TARGET_FPS = c.getint("TARGET_FPS")
        PADDING    = c.getint("PADDING")
    except Exception as e:
        sys.stderr.write(f"invalid circle setting: {e}\n")
        return False
    recompute_derived()
    return CURSOR_WINDOW_SIZE != old_size


# --- icons ----------------------------------------------------------------
def make_tray_icon(size=64):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    pad = 4
    stroke = max(2, size // 24)
    off = size // 4
    green = (124, 255, 121, 255)
    body_box = (pad, pad, size - pad, size - pad)
    # inner crescent inset from the green ring so they don't overlap
    inset = pad + stroke + 2
    cb = (inset, inset, size - inset, size - inset)
    cut_box = (cb[0] + off, cb[1] - 2, cb[2] + off, cb[3] - 2)
    shape_mask = Image.new("L", (size, size), 0)
    sd = ImageDraw.Draw(shape_mask)
    sd.ellipse(cb, fill=255)
    sd.ellipse(cut_box, fill=0)
    crescent = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    crescent.paste((230, 230, 245, 255), mask=shape_mask)
    img.alpha_composite(crescent)
    # green outer ring
    ImageDraw.Draw(img).ellipse(body_box, outline=green, width=stroke)
    return img


# --- ring drawing ---------------------------------------------------------
def aa_ring(img, cx, cy, radius, thickness, rgb, alpha_center, alpha_edge,
            halo_rgb=None, halo_thickness=0,
            halo_alpha_inner=0, halo_alpha_outer=0):
    if radius <= 0 or thickness <= 0:
        return
    scale = 3
    t_half = thickness / 2
    quart = thickness / 4
    halo = ((halo_thickness - thickness) / 2
            if halo_rgb and halo_thickness > thickness else 0.0)
    h_half = halo / 2
    outer_extent = radius + t_half + halo
    box = int((outer_extent + 2) * 2 * scale)
    tile = Image.new("RGBA", (box, box), (0, 0, 0, 0))
    td = ImageDraw.Draw(tile)
    c = box // 2

    bands = []
    r = radius + t_half + halo
    if halo > 0:
        bands.append((r, r - h_half, halo_rgb, halo_alpha_outer)); r -= h_half
        bands.append((r, r - h_half, halo_rgb, halo_alpha_inner)); r -= h_half
    bands.append((r, r - quart, rgb, alpha_edge)); r -= quart
    bands.append((r, r - 2 * quart, rgb, alpha_center)); r -= 2 * quart
    bands.append((r, r - quart, rgb, alpha_edge)); r -= quart
    if halo > 0:
        bands.append((r, r - h_half, halo_rgb, halo_alpha_inner)); r -= h_half
        bands.append((r, max(0, r - h_half), halo_rgb, halo_alpha_outer))

    for r_out, r_in, color, alpha in bands:
        if r_out <= 0 or alpha <= 0:
            continue
        ro = r_out * scale
        td.ellipse([c - ro, c - ro, c + ro, c + ro], fill=color + (alpha,))
        if r_in > 0:
            ri = r_in * scale
            td.ellipse([c - ri, c - ri, c + ri, c + ri], fill=(0, 0, 0, 0))

    tile = tile.resize((box // scale, box // scale), Image.LANCZOS)
    img.alpha_composite(tile, (int(cx - tile.width / 2), int(cy - tile.height / 2)))


def premultiply_bgra(img):
    r, g, b, a = img.split()
    r = ImageChops.multiply(r, a)
    g = ImageChops.multiply(g, a)
    b = ImageChops.multiply(b, a)
    return Image.merge("RGBA", (b, g, r, a)).tobytes("raw", "RGBA")


# --- layered window -------------------------------------------------------
_windows_by_hwnd = {}


@WNDPROC
def _shared_wnd_proc(hwnd, msg, wparam, lparam):
    w = _windows_by_hwnd.get(int(hwnd))
    if w is not None:
        r = w._handle_msg(msg, wparam, lparam)
        if r is not None:
            return r
    return user32.DefWindowProcW(hwnd, msg, wparam, lparam) or 0


class LayeredWindow:
    CLASS_NAME = "MouseCircleLayered"
    _class_registered = False

    def __init__(self, size, session_notify=False):
        self.size = size
        self.hinst = kernel32.GetModuleHandleW(None)
        self.on_session_change = None
        self._session_notify = session_notify
        if not LayeredWindow._class_registered:
            wc = WNDCLASS()
            wc.style = 0
            wc.lpfnWndProc = ctypes.cast(_shared_wnd_proc, c_void_p)
            wc.hInstance = self.hinst
            wc.lpszClassName = LayeredWindow.CLASS_NAME
            user32.RegisterClassW(byref(wc))
            LayeredWindow._class_registered = True
        ex = (WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW
              | WS_EX_NOACTIVATE | WS_EX_TOPMOST)
        self.hwnd = user32.CreateWindowExW(
            ex, LayeredWindow.CLASS_NAME, "MouseCircle",
            WS_POPUP, 0, 0, size, size,
            None, None, self.hinst, None,
        )
        if not self.hwnd:
            raise ctypes.WinError()
        _windows_by_hwnd[int(self.hwnd)] = self
        self._create_dib()
        if session_notify:
            wtsapi32.WTSRegisterSessionNotification(self.hwnd, NOTIFY_FOR_THIS_SESSION)
        user32.ShowWindow(self.hwnd, SW_SHOWNOACTIVATE)

    def _create_dib(self):
        bi = BITMAPINFO()
        bi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bi.bmiHeader.biWidth = self.size
        bi.bmiHeader.biHeight = -self.size
        bi.bmiHeader.biPlanes = 1
        bi.bmiHeader.biBitCount = 32
        bi.bmiHeader.biCompression = BI_RGB
        self.bits_ptr = c_void_p()
        screen_dc = user32.GetDC(None)
        self.mem_dc = gdi32.CreateCompatibleDC(screen_dc)
        self.hbitmap = gdi32.CreateDIBSection(
            screen_dc, byref(bi), DIB_RGB_COLORS, byref(self.bits_ptr), None, 0,
        )
        user32.ReleaseDC(None, screen_dc)
        self.old_obj = gdi32.SelectObject(self.mem_dc, self.hbitmap)

    def _handle_msg(self, msg, wparam, lparam):
        if msg == WM_WTSSESSION_CHANGE and self.on_session_change:
            self.on_session_change(int(wparam))
            return 0
        if msg == WM_APP_EXIT:
            user32.PostQuitMessage(0)
            return 0
        if msg == WM_DESTROY and self._session_notify:
            user32.PostQuitMessage(0)
            return 0
        return None

    def update_at(self, top_left_x, top_left_y, rgba_bytes):
        ctypes.memmove(self.bits_ptr, rgba_bytes, len(rgba_bytes))
        pt_dst = POINT(top_left_x, top_left_y)
        sz = SIZE(self.size, self.size)
        pt_src = POINT(0, 0)
        blend = BLENDFUNCTION(AC_SRC_OVER, 0, 255, AC_SRC_ALPHA)
        screen_dc = user32.GetDC(None)
        user32.UpdateLayeredWindow(
            self.hwnd, screen_dc, byref(pt_dst), byref(sz),
            self.mem_dc, byref(pt_src), 0, byref(blend), ULW_ALPHA,
        )
        user32.ReleaseDC(None, screen_dc)

    def update_at_cursor(self, cursor_x, cursor_y, rgba_bytes):
        self.update_at(cursor_x - self.size // 2, cursor_y - self.size // 2, rgba_bytes)

    def show(self, visible):
        user32.ShowWindow(self.hwnd, SW_SHOWNOACTIVATE if visible else SW_HIDE)

    def bring_to_top(self):
        user32.SetWindowPos(
            self.hwnd, wintypes.HWND(HWND_TOPMOST), 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        )

    def post_exit(self):
        user32.PostMessageW(self.hwnd, WM_APP_EXIT, 0, 0)

    def destroy(self):
        try:
            _windows_by_hwnd.pop(int(self.hwnd), None)
            if self._session_notify:
                wtsapi32.WTSUnRegisterSessionNotification(self.hwnd)
                self._session_notify = False
            gdi32.SelectObject(self.mem_dc, self.old_obj)
            gdi32.DeleteObject(self.hbitmap)
            gdi32.DeleteDC(self.mem_dc)
            user32.DestroyWindow(self.hwnd)
        except Exception:
            pass


# --- mouse-circle engine (runs in its own thread) -------------------------
class MouseCircle:
    def __init__(self):
        self.running = True
        self.enabled = CIRCLE_ENABLED
        self.visible = True  # session state
        self.cursor_window = None
        self.ripples = []
        self._ready = threading.Event()
        self._reconfig_pending = False

    def request_reconfigure(self):
        self._reconfig_pending = True

    def set_enabled(self, value):
        self.enabled = bool(value)
        if self.cursor_window is not None:
            self.cursor_window.show(self.enabled and self.visible)
            if not self.enabled:
                for r in self.ripples:
                    r["window"].destroy()
                self.ripples = []

    def _on_session_change(self, code):
        if code == WTS_SESSION_LOCK:
            self.visible = False
            self.cursor_window.show(False)
            for r in self.ripples:
                r["window"].show(False)
        elif code == WTS_SESSION_UNLOCK:
            self.visible = True
            if self.enabled:
                self.cursor_window.show(True)
                for r in self.ripples:
                    r["window"].show(True)

    def _spawn_ripple(self, x, y, color):
        try:
            w = LayeredWindow(RIPPLE_WINDOW_SIZE)
        except Exception as e:
            sys.stderr.write(f"ripple window failed: {e}\n")
            return
        self.ripples.append({
            "start": time.monotonic(), "color": color,
            "x": x, "y": y, "window": w,
        })

    def _poll_buttons(self, x, y):
        pressed_color = None
        for vk, color in (
            (win32con.VK_LBUTTON, _LEFT_RGB),
            (win32con.VK_RBUTTON, _RIGHT_RGB),
            (win32con.VK_MBUTTON, _MIDDLE_RGB),
        ):
            try:
                s = win32api.GetKeyState(vk)
            except Exception:
                s = 0
            if s < 0 and self.state[vk] >= 0:
                self._spawn_ripple(x, y, color)
                if self.cursor_window is not None:
                    self.cursor_window.bring_to_top()
                pressed_color = color
            self.state[vk] = s
            if s < 0 and pressed_color is None:
                pressed_color = color
        return pressed_color

    def _render_cursor(self, x, y, color):
        size = self.cursor_window.size
        center = size // 2
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        halo_rgb = (tuple(int(c * HALO_COLOR_SCALE) for c in color)
                    if HALO_ENABLED else None)
        aa_ring(
            img, center, center, RING_RADIUS, RING_THICKNESS, color,
            _RING_ALPHA_CENTER_U8, _RING_ALPHA_EDGE_U8,
            halo_rgb=halo_rgb,
            halo_thickness=RING_THICKNESS * HALO_THICKNESS_FACTOR,
            halo_alpha_inner=_HALO_ALPHA_INNER_U8,
            halo_alpha_outer=_HALO_ALPHA_OUTER_U8,
        )
        self.cursor_window.update_at_cursor(x, y, premultiply_bgra(img))

    def _render_ripples(self):
        now = time.monotonic()
        alive = []
        for r in self.ripples:
            t = (now - r["start"]) / RIPPLE_DURATION
            if t >= 1.0:
                r["window"].destroy()
                continue
            alive.append(r)
            progress = 1 - (1 - (t - 1) ** 2) ** 0.5
            radius = int(RING_RADIUS * RIPPLE_GROWTH * progress)
            fade = max(0.0, 1.0 - t)
            size = r["window"].size
            center = size // 2
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            halo_rgb = (tuple(int(c * HALO_COLOR_SCALE) for c in r["color"])
                        if HALO_ENABLED else None)
            aa_ring(
                img, center, center, radius, RIPPLE_THICKNESS, r["color"],
                int(_RING_ALPHA_CENTER_U8 * fade), int(_RING_ALPHA_EDGE_U8 * fade),
                halo_rgb=halo_rgb,
                halo_thickness=RIPPLE_THICKNESS * HALO_THICKNESS_FACTOR,
                halo_alpha_inner=int(_HALO_ALPHA_INNER_U8 * fade),
                halo_alpha_outer=int(_HALO_ALPHA_OUTER_U8 * fade),
            )
            r["window"].update_at(
                r["x"] - center, r["y"] - center, premultiply_bgra(img),
            )
        self.ripples = alive

    def _render(self):
        try:
            x, y = win32api.GetCursorPos()
        except Exception:
            return
        pressed = self._poll_buttons(x, y)
        color = pressed or _IDLE_RGB
        self._render_cursor(x, y, color)
        self._render_ripples()

    def post_exit(self):
        self.running = False
        if self.cursor_window is not None:
            self.cursor_window.post_exit()

    def run(self):
        self.cursor_window = LayeredWindow(CURSOR_WINDOW_SIZE, session_notify=True)
        self.cursor_window.on_session_change = self._on_session_change
        self.state = {
            win32con.VK_LBUTTON: win32api.GetKeyState(win32con.VK_LBUTTON),
            win32con.VK_RBUTTON: win32api.GetKeyState(win32con.VK_RBUTTON),
            win32con.VK_MBUTTON: win32api.GetKeyState(win32con.VK_MBUTTON),
        }
        self._ready.set()

        msg = MSG()
        next_tick = time.monotonic()
        while self.running:
            while user32.PeekMessageW(byref(msg), None, 0, 0, PM_REMOVE):
                if msg.message == WM_QUIT:
                    self.running = False
                    break
                user32.TranslateMessage(byref(msg))
                user32.DispatchMessageW(byref(msg))
            if not self.running:
                break
            if self._reconfig_pending:
                self._reconfig_pending = False
                try:
                    old = self.cursor_window
                    self.cursor_window = LayeredWindow(
                        CURSOR_WINDOW_SIZE, session_notify=True)
                    self.cursor_window.on_session_change = self._on_session_change
                    if not (self.enabled and self.visible):
                        self.cursor_window.show(False)
                    self.cursor_window.bring_to_top()
                    old.destroy()
                except Exception as e:
                    sys.stderr.write(f"reconfigure failed: {e}\n")
            if self.enabled and self.visible:
                try:
                    self._render()
                except Exception as e:
                    sys.stderr.write(f"render error: {e}\n")
                    time.sleep(0.5)
            elif self.ripples:
                # drain any leftover ripples when disabled mid-animation
                for r in self.ripples:
                    r["window"].destroy()
                self.ripples = []
            next_tick += 1 / max(1, TARGET_FPS)
            sleep = next_tick - time.monotonic()
            if sleep > 0:
                time.sleep(sleep)
            else:
                next_tick = time.monotonic()
        for r in self.ripples:
            r["window"].destroy()
        if self.cursor_window is not None:
            self.cursor_window.destroy()


# --- dimmer ---------------------------------------------------------------
def _scale_matrix(s):
    m = MAGCOLOREFFECT()
    m.transform[0][0] = s
    m.transform[1][1] = s
    m.transform[2][2] = s
    m.transform[3][3] = 1.0
    m.transform[4][4] = 1.0
    return m


def mag_set_scale(s):
    m = _scale_matrix(max(0.0, min(1.0, s)))
    return bool(magnification.MagSetFullscreenColorEffect(byref(m)))


class App:
    def __init__(self):
        self.level = 0
        self.last_nonzero = INITIAL_NONZERO if INITIAL_NONZERO > 0 else 40
        self._anim_job = None
        self._current = 0.0
        self._save_job = None
        self._settings_win = None
        self._reconfig_job = None

        self.root = tk.Tk()
        self.root.withdraw()

        if not magnification.MagInitialize():
            raise RuntimeError("MagInitialize failed")
        mag_set_scale(1.0)

        self.mouse = MouseCircle()
        self.mouse_thread = threading.Thread(target=self.mouse.run, daemon=True)
        self.mouse_thread.start()
        self.mouse._ready.wait(timeout=5)
        self.mouse.set_enabled(CIRCLE_ENABLED)

        def level_item(p):
            def on_click(_icon, _item):
                self.root.after(0, lambda: self._set(p))

            def is_checked(_item):
                return self.level == p

            if p == 0:
                label = "0% (off)"
            elif p == max(LEVELS):
                label = f"{p}% (darkest)"
            else:
                label = f"{p}%"
            return pystray.MenuItem(label, on_click, checked=is_checked, radio=True)

        def on_toggle(_icon, _item):
            self.root.after(0, self._toggle)

        def on_mouse_circle(_icon, _item):
            self.root.after(0, self._toggle_mouse_circle)

        def mouse_checked(_item):
            return self.mouse.enabled

        self.tray = pystray.Icon(
            "mousecircledim",
            make_tray_icon(),
            "MouseCircleDim",
            menu=pystray.Menu(
                pystray.MenuItem("Toggle dim", on_toggle, default=True, visible=False),
                *(level_item(p) for p in LEVELS),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Mouse circle", on_mouse_circle, checked=mouse_checked),
                pystray.MenuItem("Edit settings…", self._open_settings),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._tray_quit),
            ),
        )
        threading.Thread(target=self.tray.run, daemon=True).start()

        if INITIAL_LEVEL > 0:
            self.root.after(0, lambda: self._set(INITIAL_LEVEL))

    def _apply(self, a):
        self._current = a
        try:
            mag_set_scale(1.0 - a)
        except Exception:
            pass

    def _set(self, p):
        self.level = p
        if p > 0:
            self.last_nonzero = p
        self._animate_to(p / 100.0, duration_ms=500)
        _cfg["dim"]["LAST_LEVEL"] = str(p)
        _cfg["dim"]["LAST_NONZERO"] = str(self.last_nonzero)
        self._schedule_save()
        try:
            self.tray.update_menu()
        except Exception:
            pass

    def _schedule_save(self, delay_ms=3000):
        if self._save_job is not None:
            try:
                self.root.after_cancel(self._save_job)
            except Exception:
                pass
        self._save_job = self.root.after(delay_ms, self._do_save)

    def _do_save(self):
        self._save_job = None
        _save_ini()

    def _flush_save(self):
        if self._save_job is not None:
            try:
                self.root.after_cancel(self._save_job)
            except Exception:
                pass
            self._save_job = None
            _save_ini()

    def _animate_to(self, target, duration_ms=500, step_ms=16):
        if self._anim_job is not None:
            try:
                self.root.after_cancel(self._anim_job)
            except Exception:
                pass
            self._anim_job = None

        start = self._current
        delta = target - start
        if abs(delta) < 1e-4 or duration_ms <= 0:
            self._apply(target)
            return

        steps = max(1, duration_ms // step_ms)
        i = {"n": 0}

        def ease(t):
            return 3 * t * t - 2 * t * t * t

        def tick():
            i["n"] += 1
            t = min(1.0, i["n"] / steps)
            self._apply(start + delta * ease(t))
            if t < 1.0:
                self._anim_job = self.root.after(step_ms, tick)
            else:
                self._anim_job = None

        self._anim_job = self.root.after(step_ms, tick)

    def _toggle(self):
        self._set(0 if self.level > 0 else self.last_nonzero)

    def _toggle_mouse_circle(self):
        self.mouse.set_enabled(not self.mouse.enabled)
        _cfg["circle"]["ENABLED"] = "true" if self.mouse.enabled else "false"
        self._schedule_save()
        try:
            self.tray.update_menu()
        except Exception:
            pass

    def _open_settings(self, _icon=None, _item=None):
        self.root.after(0, self._show_settings_window)

    def _show_settings_window(self):
        if self._settings_win is not None and self._settings_win.winfo_exists():
            self._settings_win.lift()
            self._settings_win.focus_force()
            return

        win = tk.Toplevel(self.root)
        self._settings_win = win
        win.title("MouseCircleDim settings")
        win.resizable(False, False)
        try:
            icon_photo = ImageTk.PhotoImage(make_tray_icon(32), master=win)
            win.iconphoto(True, icon_photo)
            win._icon_ref = icon_photo  # prevent GC
        except Exception:
            pass

        # kind: "int" | "f1" | "f2" | "bool" | "color"
        # (section, key, label, kind, min, max, step)
        FIELDS = [
            ("circle", "IDLE_COLOR",         "Idle color (R G B)",     "color"),
            ("circle", "LEFT_CLICK_COLOR",   "Left-click color",       "color"),
            ("circle", "RIGHT_CLICK_COLOR",  "Right-click color",      "color"),
            ("circle", "MIDDLE_CLICK_COLOR", "Middle-click color",     "color"),
            ("circle", "RING_RADIUS",        "Ring radius (px)",       "int",  1,  500, 1),
            ("circle", "RING_THICKNESS",     "Ring thickness (px)",    "int",  1,   60, 1),
            ("circle", "RING_ALPHA_CENTER",  "Ring alpha (center)",    "f2",   0.0, 1.0, 0.05),
            ("circle", "RING_ALPHA_EDGE",    "Ring alpha (edge)",      "f2",   0.0, 1.0, 0.05),
            ("circle", "HALO_ENABLED",          "Halo enabled",        "bool"),
            ("circle", "HALO_THICKNESS_FACTOR", "Halo thickness ×",    "int",  1,  10, 1),
            ("circle", "HALO_COLOR_SCALE",      "Halo color scale",    "f2",   0.0, 1.0, 0.05),
            ("circle", "HALO_ALPHA_INNER",      "Halo alpha (inner)",  "f2",   0.0, 1.0, 0.05),
            ("circle", "HALO_ALPHA_OUTER",      "Halo alpha (outer)",  "f2",   0.0, 1.0, 0.05),
            ("circle", "RIPPLE_THICKNESS", "Ripple thickness (px)",    "int",  1,  40, 1),
            ("circle", "RIPPLE_GROWTH",    "Ripple growth ×",          "f1",   1.0, 10.0, 0.1),
            ("circle", "RIPPLE_DURATION",  "Ripple duration (s)",      "f1",   0.1, 10.0, 0.1),
            ("circle", "TARGET_FPS",       "Target FPS",               "int",  15, 240, 5),
            ("circle", "PADDING",          "Padding (px)",             "int",  0,  50, 1),
        ]

        def bind_wheel(spin):
            def on_wheel(e):
                spin.invoke("buttonup" if e.delta > 0 else "buttondown")
                return "break"
            spin.bind("<MouseWheel>", on_wheel)

        FMT = {"int": "%.0f", "f1": "%.1f", "f2": "%.2f"}

        def make_spin(parent, kind, lo, hi, step, initial, build_edit):
            """build_edit(var) -> on_edit callable. Ensures closure captures
            this iteration's var, not the loop's trailing binding."""
            fmt = FMT[kind]
            var = tk.StringVar(value=fmt % float(initial))
            on_edit = build_edit(var)
            sp = tk.Spinbox(
                parent, from_=lo, to=hi, increment=step, format=fmt,
                textvariable=var, width=7, justify="right",
                command=on_edit,
            )
            var.trace_add("write", lambda *_: on_edit())
            bind_wheel(sp)
            return sp, var

        frame = tk.Frame(win, padx=12, pady=12)
        frame.pack(fill="both", expand=True)

        reset_actions = []

        def commit(section, key, value):
            _cfg[section][key] = value
            self._schedule_save()
            self._apply_live()

        for row, spec in enumerate(FIELDS):
            section, key, label, kind = spec[:4]
            tk.Label(frame, text=label, anchor="w").grid(
                row=row, column=0, sticky="w", padx=(0, 10), pady=2)

            if kind == "bool":
                var = tk.BooleanVar(value=_cfg[section].getboolean(key))
                def make_cb(section=section, key=key, var=var):
                    def cb(*_):
                        commit(section, key, "true" if var.get() else "false")
                    return cb
                tk.Checkbutton(frame, variable=var, command=make_cb()).grid(
                    row=row, column=1, sticky="w", pady=2)
                default_str = _DEFAULTS[section][key]
                reset_actions.append(
                    lambda v=var, d=default_str: v.set(d.lower() == "true"))

            elif kind == "color":
                vals = [float(x.strip()) for x in _cfg[section][key].split(",")]
                vals = (vals + [0.0, 0.0, 0.0])[:3]
                holder = tk.Frame(frame)
                holder.grid(row=row, column=1, sticky="w", pady=2)
                stringvars = []
                def color_on_edit(section=section, key=key, svs=stringvars):
                    try:
                        parts = ["%.2f" % float(sv.get()) for sv in svs]
                    except ValueError:
                        return
                    commit(section, key, ", ".join(parts))
                def build_color_edit(_var):
                    return color_on_edit
                for i, v in enumerate(vals):
                    sp, sv = make_spin(holder, "f2", 0.0, 1.0, 0.05, v, build_color_edit)
                    sp.pack(side="left", padx=(0 if i == 0 else 4, 0))
                    stringvars.append(sv)
                default_parts = [float(x.strip())
                                 for x in _DEFAULTS[section][key].split(",")]
                def reset_color(svs=stringvars, parts=default_parts):
                    for sv, p in zip(svs, parts):
                        sv.set("%.2f" % p)
                reset_actions.append(reset_color)

            else:
                lo, hi, step = spec[4], spec[5], spec[6]
                initial = float(_cfg[section][key])
                fmt = FMT[kind]
                def build_edit(var, section=section, key=key, fmt=fmt):
                    def on_edit():
                        try:
                            commit(section, key, fmt % float(var.get()))
                        except ValueError:
                            return
                    return on_edit
                sp, var = make_spin(frame, kind, lo, hi, step, initial, build_edit)
                sp.grid(row=row, column=1, sticky="w", pady=2)
                default_str = _DEFAULTS[section][key]
                reset_actions.append(
                    lambda v=var, fmt=fmt, d=default_str: v.set(fmt % float(d)))

        btns = tk.Frame(frame)
        btns.grid(row=len(FIELDS), column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Button(btns, text="Defaults",
                  command=lambda: [a() for a in reset_actions]).pack(side="left", padx=4)
        tk.Button(btns, text="Close", command=lambda: on_close()).pack(side="right", padx=4)

        def on_close():
            self._flush_save()
            win.destroy()
            self._settings_win = None

        win.protocol("WM_DELETE_WINDOW", on_close)

    def _apply_live(self):
        resized = apply_circle_settings_from_cfg()
        if resized:
            if self._reconfig_job is not None:
                try:
                    self.root.after_cancel(self._reconfig_job)
                except Exception:
                    pass
            self._reconfig_job = self.root.after(200, self._do_reconfig)

    def _do_reconfig(self):
        self._reconfig_job = None
        self.mouse.request_reconfigure()

    def _tray_quit(self, _icon, _item):
        self.root.after(0, self.quit)

    def quit(self):
        try:
            self._flush_save()
        except Exception:
            pass
        try:
            self.mouse.post_exit()
        except Exception:
            pass
        try:
            mag_set_scale(1.0)
            magnification.MagUninitialize()
        except Exception:
            pass
        try:
            self.tray.stop()
        except Exception:
            pass
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
