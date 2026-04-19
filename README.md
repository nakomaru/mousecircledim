# MouseCircleDim

A small Windows tray app that combines two utilities:

- **MouseCircle** — a click-through topmost ring that follows the cursor, with color-coded click ripples (left / right / middle).
- **VirtualDim** — a global screen dimmer that uses the Magnification API's `MagSetFullscreenColorEffect` to scale RGB output before scan-out, so it dims basically everything (including popups, context menus, and top-most windows) except the hardware cursor.

Both run from a single tray icon. The tray menu lets you pick a dim level, toggle the mouse circle on/off, and quit. Left-clicking the tray icon toggles dim between 0 and the last level.

## Running

```
pythonw mousecircledim.pyw
```

On first run it bootstraps a venv at `~/venvs/.venv_mousecircledim` and installs `pywin32`, `Pillow`, and `pystray`.

Windows only.

## License

CC0 1.0