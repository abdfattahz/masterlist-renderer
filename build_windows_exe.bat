@echo off
setlocal

echo Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Building Windows executable...
pyinstaller --noconfirm --clean --onefile --windowed --name MasterlistRendererGUI masterlist_gui.py

echo.
echo Build complete. EXE path:
echo dist\MasterlistRendererGUI.exe

endlocal
