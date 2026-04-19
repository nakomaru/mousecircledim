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
import ctypes
import threading
import time
import tkinter as tk
from ctypes import (
    POINTER, Structure, WINFUNCTYPE, byref, c_float, c_int, c_ubyte, c_void_p, wintypes,
)

import win32api
import win32con
import win32ts
from PIL import Image, ImageChops, ImageDraw
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
IDLE_COLOR         = (0.88, 0.94, 1.00)
LEFT_CLICK_COLOR   = (0.48, 0.87, 0.30)
RIGHT_CLICK_COLOR  = (0.87, 0.30, 0.22)
MIDDLE_CLICK_COLOR = (0.30, 0.48, 0.87)

RING_RADIUS       = 41
RING_THICKNESS    = 6
RING_ALPHA_CENTER = 0.50
RING_ALPHA_EDGE   = 0.40

HALO_ENABLED          = True
HALO_THICKNESS_FACTOR = 2
HALO_COLOR_SCALE      = 0.2
HALO_ALPHA_INNER      = 0.25
HALO_ALPHA_OUTER      = 0.05

RIPPLE_THICKNESS = 3
RIPPLE_GROWTH    = 3.0
RIPPLE_DURATION  = 1.0

TARGET_FPS = 120
PADDING    = 4

LEVELS = [90, 85, 80, 70, 60, 50, 40, 30, 20, 10, 0]


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


# --- icons ----------------------------------------------------------------
def make_tray_icon(size=64):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # moon (dimmer)
    pad = 6
    d.ellipse((pad, pad, size - pad, size - pad), fill=(230, 230, 245, 255))
    off = size // 4
    d.ellipse((pad + off, pad - 2, size - pad + off, size - pad - 2),
              fill=(0, 0, 0, 0))
    # small crosshair ring (mouse circle) overlaid bottom-right
    c = (124, 255, 121, 255)
    r = size // 3
    cx = size - r // 2 - 2
    cy = size - r // 2 - 2
    d.ellipse([cx - r // 2, cy - r // 2, cx + r // 2, cy + r // 2],
              outline=c, width=2)
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

    def post_exit(self):
        user32.PostMessageW(self.hwnd, WM_APP_EXIT, 0, 0)

    def destroy(self):
        try:
            _windows_by_hwnd.pop(int(self.hwnd), None)
            if self._session_notify:
                wtsapi32.WTSUnRegisterSessionNotification(self.hwnd)
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
        self.enabled = True
        self.visible = True  # session state
        self.cursor_window = None
        self.ripples = []
        self._ready = threading.Event()

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
                pressed_color = color
            self.state[vk] = s
            if s < 0 and pressed_color is None:
                pressed_color = color
        return pressed_color

    def _render_cursor(self, x, y, color):
        center = CURSOR_WINDOW_SIZE // 2
        img = Image.new("RGBA", (CURSOR_WINDOW_SIZE, CURSOR_WINDOW_SIZE), (0, 0, 0, 0))
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
        center = RIPPLE_WINDOW_SIZE // 2
        for r in self.ripples:
            t = (now - r["start"]) / RIPPLE_DURATION
            if t >= 1.0:
                r["window"].destroy()
                continue
            alive.append(r)
            progress = 1 - (1 - (t - 1) ** 2) ** 0.5
            radius = int(RING_RADIUS * RIPPLE_GROWTH * progress)
            fade = max(0.0, 1.0 - t)
            img = Image.new("RGBA", (RIPPLE_WINDOW_SIZE, RIPPLE_WINDOW_SIZE), (0, 0, 0, 0))
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
        target_dt = 1 / TARGET_FPS
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
            next_tick += target_dt
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
        self.last_nonzero = 40
        self._anim_job = None
        self._current = 0.0

        self.root = tk.Tk()
        self.root.withdraw()

        if not magnification.MagInitialize():
            raise RuntimeError("MagInitialize failed")
        mag_set_scale(1.0)

        self.mouse = MouseCircle()
        self.mouse_thread = threading.Thread(target=self.mouse.run, daemon=True)
        self.mouse_thread.start()
        self.mouse._ready.wait(timeout=5)

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
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._tray_quit),
            ),
        )
        threading.Thread(target=self.tray.run, daemon=True).start()

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
        try:
            self.tray.update_menu()
        except Exception:
            pass

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
        try:
            self.tray.update_menu()
        except Exception:
            pass

    def _tray_quit(self, _icon, _item):
        self.root.after(0, self.quit)

    def quit(self):
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
