chcp 65001
@echo off
echo 正在安装依赖包...
pip install -r requirements.txt

if %errorlevel% equ 0 (
    echo 依赖包安装成功
) else (
    echo 依赖包安装失败，请检查requirements.txt文件或网络连接
    pause
    exit /b 1
)

echo.
echo 正在创建start_FK.bat的快捷方式到开机启动目录...

:: 获取当前脚本所在目录
set "current_dir=%~dp0"
:: 目标批处理文件路径
set "target_bat=%current_dir%start_FK.bat"

:: 检查目标批处理文件是否存在
if not exist "%target_bat%" (
    echo 错误：未找到 %target_bat%
    pause
    exit /b 1
)

:: 获取当前用户的开机启动目录
set "startup_dir=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

:: 创建VBScript临时文件用于生成快捷方式
set "vbs_file=%temp%\create_shortcut.vbs"

(
echo Set WshShell = WScript.CreateObject("WScript.Shell"^)
echo strDesktop = "%startup_dir%"
echo Set oShellLink = WshShell.CreateShortcut(strDesktop ^& "\start_FK.lnk"^)
echo oShellLink.TargetPath = "%target_bat%"
echo oShellLink.WorkingDirectory = "%current_dir%"
echo oShellLink.Save
) > "%vbs_file%"

:: 执行VBScript创建快捷方式
cscript //nologo "%vbs_file%"

:: 清理临时文件
del /f /q "%vbs_file%" >nul 2>&1

:: 检查快捷方式是否创建成功
if exist "%startup_dir%\start_FK.lnk" (
    echo 快捷方式已成功添加到开机启动目录
) else (
    echo 快捷方式创建失败
    pause
    exit /b 1
)

echo.
echo 所有操作完成
pause