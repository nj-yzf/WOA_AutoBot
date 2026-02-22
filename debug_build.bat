@echo off
cd /d "%~dp0"
echo [DEBUG] PyInstaller build starting...
if not exist "assets\u2.jar" (
    echo [Pre-build] Preparing assets\u2.jar ...
    python -c "import uiautomator2 as u2, os, shutil; d=os.getcwd(); os.makedirs('assets', exist_ok=True); s=os.path.join(os.path.dirname(u2.__file__), 'assets', 'u2.jar'); shutil.copy2(s, os.path.join(d, 'assets', 'u2.jar')) if os.path.isfile(s) else exit(1)" 2>nul
    if errorlevel 1 python -m uiautomator2 copy-assets 2>nul
)
call pyinstaller -D -c --clean --noupx --name "WOA_Debug" --version-file "version_info.txt" --add-data "icon;icon" --add-data "adb_tools;adb_tools" --add-data "assets;assets" --hidden-import ttkbootstrap --hidden-import PIL --hidden-import cv2 --hidden-import numpy --hidden-import emulator_discovery --hidden-import main_adb --hidden-import uiautomator2 gui_launcher.py
echo.
echo Done. Output: dist\WOA_Debug\WOA_Debug.exe
pause
