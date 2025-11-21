# VB6 HotKey Class

A small event-driven HotKey wrapper for Visual Basic 6.

This project is implemented as a DLL but is intentionally simple:
The code can be copied directly into other VB6 projects and used without external dependencies or runtime modules.
It uses the Windows RegisterHotKey / UnregisterHotKey API and receives WM_HOTKEY messages, exposing a clean VB6-friendly interface.

Key points
- Event-driven: raises a `HotKeyPressed` event when a registered hotkey is pressed.
- Supports multiple hotkeys at once.
- Hooks key combinations made from at least 2 modifier keys chosen from Ctrl / Alt / Shift plus one other key.
- No external dependencies; can be used either by referencing the compiled DLL or by copying the class and supporting modules into your project.

Lightweight, responsive, and optimized for simplicity.

---

## Requirements

- Visual Basic 6

---

## Design Overview

- Class exposes:
  - Event: `HotKeyPressed` — invoked when a registered hotkey is pressed.
  - Properties:
    - `hWnd As Long` — owner window handle (must be set to `Form.hWnd` or similar).
    - `Hotkeys As Long` — read-only count of currently registered hotkeys.
    - `Hotkey(Index) As HOTKEY_Type` — returns the registered hotkey info at given index.
  - Functions:
    - `RegisterHotKey(KeyCode, Modifiers)`
    - `UnRegisterHotKey(KeyCode, Modifiers)`
    - `UnRegisterAllHotKeys()`

- Type:
  - `HOTKEY_Type`:
    - `KeyCode As KeyCodeConstants`
    - `Modifiers As RegisterHotKeyModifiers`

---

## Modifier constants and key codes

## Best practices and tips

- Always set `hWnd` to the form that should own the hotkeys before calling `RegisterHotKey`.
- Unregister hotkeys in `Form_Unload` (or appropriate shutdown path) using `UnRegisterAllHotKeys`.
- The class validates that at least two modifier flags (Ctrl, Alt, Shift) are set.
- The class assigns and manages internal IDs for each registration; you need to call `UnRegisterHotKey` with the same KeyCode and Modifiers.
- If `RegisterHotKey` fails, another application may already hold that global hotkey — handle this gracefully.

---

## License

[MIT License](LICENSE)  
Copyright © Ubehage

---

## Credits

Created by Ubehage  

[GitHub Profile](https://github.com/Ubehage)