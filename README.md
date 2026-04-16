# Office Assistant Clone

GTK-based Office Assistant style helper for Ubuntu/Linux.

## Features
- double-click the assistant to open the Office-style control window
- Gallery tab supports Clippit, Dog, Merlin, and Genius
- speech bubble is drawn at the top in an Office-like style
- Behavior tab edits the active JSON behavior profile inside the app
- profile file lives in `profiles/default.json`
- ZIP/RAR/TAR/GZ and similar archive creation defaults to the `Save` animation

## Run

```bash
python3 app.py
```

## Dependencies

```bash
sudo apt update
sudo apt install -y python3-gi python3-watchdog gir1.2-gtk-3.0
```

## License
This project is licensed under the MIT License (see LICENSE file for details).
Important Notice Regarding Microsoft Agents
The JavaScript code and animations framework used in this project are based on clippy.js, which is licensed under the MIT License for the code only.
All Office Assistant characters (Clippit / Clippy, Merlin, Rover / Dog, Genius, and others), including their names, images, animations, and related assets,
remain the property of Microsoft Corporation.
These characters are not covered by the MIT License of this project.
They are used here solely for nostalgic and non-commercial purposes.
No ownership is claimed over the Microsoft characters or trademarks.
This software is a fan-made nostalgic recreation.
For more information, see the original clippy.js repository:
https://github.com/clippyjs/clippy.js

## Notes
- File save reactions ignore noisy temp/autosave patterns from editors by default.
- Global typing / compile / mail / audio events are present in the profile and manual UI, but only the file/recent/startup/exit/archive detectors are implemented by default.
