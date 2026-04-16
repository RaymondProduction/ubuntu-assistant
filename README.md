# Office Assistant Clone for Ubuntu / Wayland

This version adds a retro Office 95/98 inspired control window and multiple assistants.

## Included assistants
- Clippit
- Dog
- Merlin
- Genius

## Main behavior
- transparent floating assistant window
- Office-style speech bubble above the character
- double-click the assistant to open the Office Assistant window
- watches `~/Desktop`, `~/Downloads`, and `~/Documents`
- reacts to file creation, modification, deletion, moves, and recently opened files

## Office Assistant window
- **Actions** tab first
- **Gallery** tab second
- `Back` and `Next` buttons in Gallery
- `OK` applies the currently previewed assistant
- `Cancel` closes the window without changing the active assistant

## Install
```bash
sudo apt update
sudo apt install -y python3-gi python3-watchdog gir1.2-gtk-3.0 libcanberra-gtk3-module libcanberra-gtk3-0 pulseaudio-utils
```

## Run
```bash
python3 app.py
```

## Notes
- Tray-based menus are intentionally avoided because GNOME/Wayland tray support is inconsistent.
- The control window uses a custom retro UI instead of relying on your system theme.

## Tray support

On Ubuntu, the assistant can show an indicator in the top bar/tray if Ayatana AppIndicator is available.

```bash
sudo apt install -y gir1.2-ayatanaappindicator3-0.1 libayatana-appindicator3-1
```