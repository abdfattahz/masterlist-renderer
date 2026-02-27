# Masterlist Renderer GUI

This project now includes a desktop GUI so non-technical users can run the renderer without command-line arguments.

## Files

- `render_masterlist.py`: core rendering engine + CLI
- `masterlist_gui.py`: desktop GUI app
- `build_windows_exe.bat`: helper to build a Windows `.exe`

## Linux Usage (Python script)

1. Create and activate a virtual environment (recommended)
2. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

3. Launch the GUI:

```bash
python masterlist_gui.py
```

4. In the GUI:
   - Pick the Excel file (`.xlsx` / `.xls`)
   - Choose output folder
   - Optionally choose background image and font
   - Optionally set custom table colors (Row A, Row B, Header, Border)
   - If auto-match and custom table colors are both set, custom colors win
   - Click **Generate PNG Pages**

## Windows EXE Build

> Build the `.exe` on a Windows machine (PyInstaller does not reliably cross-compile Windows executables from Linux).

1. Install Python on Windows
2. Copy this project folder to Windows
3. Double-click `build_windows_exe.bat` or run it in Command Prompt
4. The generated executable will be at:

`dist\MasterlistRendererGUI.exe`

## Notes for Sharing

- Share `MasterlistRendererGUI.exe` with colleagues.
- They do not need to install Python to run the `.exe`.
- Keep the Excel file columns as:
  - `COMPANY NAME`
  - `COMPANY NO.`
