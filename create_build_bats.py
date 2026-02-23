# -*- coding: utf-8 -*-
"""Generate build.bat and debug_build.bat with ASCII + CRLF (no BOM).
Run: python create_build_bats.py
"""
import os
ROOT = os.path.dirname(os.path.abspath(__file__))

BUILD_BAT = r'''@echo off
cd /d "%~dp0"
echo Current dir: %cd%

echo [1/6] Checking adb_tools...
if not exist "adb_tools" (
    echo ERROR: adb_tools folder not found!
    pause
    exit /b 1
)
if not exist "assets\u2.jar" (
    echo [Pre-build] Preparing assets\u2.jar for uiautomator2...
    python -c "import uiautomator2 as u2, os, shutil; d=os.getcwd(); os.makedirs('assets', exist_ok=True); s=os.path.join(os.path.dirname(u2.__file__), 'assets', 'u2.jar'); shutil.copy2(s, os.path.join(d, 'assets', 'u2.jar')) if os.path.isfile(s) else exit(1)" 2>nul
    if errorlevel 1 python -m uiautomator2 copy-assets 2>nul
)

echo [2/6] Cleaning old build...
if exist dist_nuitka rmdir /s /q dist_nuitka
if exist gui_launcher.build rmdir /s /q gui_launcher.build
if exist gui_launcher.dist rmdir /s /q gui_launcher.dist

echo [3/6] Running Nuitka...
call python -m nuitka --standalone --output-filename=WOA_AutoBot.exe --windows-console-mode=disable --python-flag=no_docstrings --nofollow-import-to=cv2.gapi --nofollow-import-to=cv2.ml --windows-product-name="WOA AutoBot" --windows-product-version=1.2.4 --windows-file-version=1.2.4 --windows-company-name="WOA AutoBot" --windows-file-description="WOA Airport Game Automation Bot" --plugin-enable=tk-inter --include-package=ttkbootstrap --include-package=cv2 --include-package=numpy --include-package=PIL --include-data-dir="%~dp0icon"=icon --include-raw-dir="%~dp0adb_tools"=adb_tools --include-data-dir="%~dp0assets"=assets --windows-icon-from-ico="%~dp0icon\app.ico" --output-dir="%~dp0dist_nuitka" --jobs=4 "%~dp0gui_launcher.py"

if errorlevel 1 (
    echo Build FAILED.
    pause
    exit /b 1
)

echo [5/6] Organizing output...
if exist "dist_nuitka\gui_launcher.dist" ren "dist_nuitka\gui_launcher.dist" "WOA_AutoBot"

(
echo @echo off
echo cd /d "%%~dp0"
echo start "" "WOA_AutoBot.exe"
) > "dist_nuitka\WOA_AutoBot\Launch_WOA_AutoBot.bat"

if exist "dist_nuitka\WOA_AutoBot\icon" attrib +h "dist_nuitka\WOA_AutoBot\icon"

echo [6/6] Done!
echo Output: dist_nuitka\WOA_AutoBot\
pause
'''

DEBUG_BAT = r'''@echo off
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
'''

def main():
    crlf = b"\r\n"
    for name, content in [("build.bat", BUILD_BAT), ("debug_build.bat", DEBUG_BAT)]:
        path = os.path.join(ROOT, name)
        data = content.replace("\n", "\r\n").encode("ascii")
        with open(path, "wb") as f:
            f.write(data)
        print("Created: %s" % path)

if __name__ == "__main__":
    main()
