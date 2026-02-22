@echo off
chcp 65001 >nul
echo ========================================
echo   uiautomator2 傻瓜式配置
echo ========================================
echo.

:: 1. 安装 uiautomator2 与 uiautomator2cache（ALAS 方案，含预置 u2.jar）
echo [1/3] 安装 uiautomator2 与 uiautomator2cache ...
pip install uiautomator2 uiautomator2cache
if errorlevel 1 (
    echo 安装失败！请检查 pip 是否可用。
    pause
    exit /b 1
)
echo 安装完成。
echo.

:: 2. 复制资源到脚本目录（必须在 WOA_Speed_Test 目录下执行）
echo [2/3] 复制 u2.jar 到 assets ...
cd /d "%~dp0"
python -m uiautomator2 copy-assets
if not exist "assets\u2.jar" (
    echo 复制失败，尝试手动方式...
    python -c "import uiautomator2 as u2, os, shutil; d=os.getcwd(); os.makedirs('assets',exist_ok=True); s=os.path.join(os.path.dirname(u2.__file__),'assets','u2.jar'); shutil.copy2(s,os.path.join(d,'assets','u2.jar')) if os.path.isfile(s) else print('未找到 u2.jar')"
)
echo.

:: 3. 初始化设备（可选，需连接模拟器）
echo [3/3] 初始化设备（需已连接模拟器）...
python -m uiautomator2 init 2>nul
if errorlevel 1 (
    echo 若未连接设备可忽略。连接设备后可单独运行: python -m uiautomator2 init
) else (
    echo 设备初始化完成。
)
echo.
echo ========================================
echo   配置完成！可启动脚本并选择 uiautomator2 触控方式。
echo ========================================
pause
