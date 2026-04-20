# MouseCircleDim

A small Windows tray app that combines two utilities:

- **MouseCircle** — a click-through topmost ring that follows the cursor, with color-coded click ripples (left / right / middle).
- **VirtualDim** — a global screen dimmer that uses the Magnification API's `MagSetFullscreenColorEffect` to scale RGB output before scan-out, so it dims basically everything (including popups, context menus, and top-most windows) except the hardware cursor.

Both run from a single tray icon. The tray menu lets you pick a dim level, toggle the mouse circle on/off, open a settings window, and quit. Left-clicking the tray icon toggles dim between 0 and the last level.

## Settings

Picking **Edit settings…** from the tray menu opens a live settings window for the cursor ring — colors (idle / left / right / middle), ring radius, thickness, alpha gradient, halo, ripple thickness / growth / duration, target FPS, and padding. Changes apply immediately and are written to `mousecircledim.ini` next to the script, so they persist across runs. The INI is auto-created with defaults on first launch; deleting it restores defaults.

## Running

```
pythonw mousecircledim.pyw
```

On first run it bootstraps a venv at `~/venvs/.venv_mousecircledim` and installs `pywin32`, `Pillow`, and `pystray`.

Windows only.

## License

CC0 1.0