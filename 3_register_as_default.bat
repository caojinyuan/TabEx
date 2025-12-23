@echo off
chcp 65001 > nul
echo ======================================
echo TabExplorer 注册为系统文件管理器
echo ======================================
echo.
echo 此操作将：
echo 1. 在文件夹右键菜单添加"open with TabExplorer"
echo 2. 在文件右键菜单添加"open with TabExplorer"（打开文件所在文件夹）
echo 3. 将 TabExplorer 设置为文件夹默认打开方式
echo.
echo 需要管理员权限，请确认继续...
pause

REM 获取当前脚本所在目录（去掉末尾的反斜杠）
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "EXE_PATH=%SCRIPT_DIR%\TabExplorer.exe"

REM 将路径中的反斜杠转义为双反斜杠（用于注册表）
set "EXE_PATH_REG=%EXE_PATH:\=\\%"

REM 检查 exe 文件是否存在
if not exist "%EXE_PATH%" (
    echo.
    echo 错误：找不到 TabExplorer.exe
    echo 请先运行 2_build_exe.bat 生成可执行文件
    echo.
    pause
    exit /b 1
)

echo.
echo 正在注册到注册表...
echo 可执行文件路径: %EXE_PATH%
echo.

REM 创建临时注册表文件
set "REG_FILE=%TEMP%\TabExplorer_register.reg"

(
echo Windows Registry Editor Version 5.00
echo.
echo ; 添加到文件夹右键菜单
echo [HKEY_CLASSES_ROOT\Directory\shell\TabExplorer]
echo @="open with TabExplorer"
echo "Icon"="%EXE_PATH_REG%"
echo.
echo [HKEY_CLASSES_ROOT\Directory\shell\TabExplorer\command]
echo @="\"%EXE_PATH_REG%\" \"%%1\""
echo.
echo ; 添加到文件右键菜单
echo [HKEY_CLASSES_ROOT\*\shell\TabExplorer]
echo @="open with TabExplorer"
echo "Icon"="%EXE_PATH_REG%"
echo.
echo [HKEY_CLASSES_ROOT\*\shell\TabExplorer\command]
echo @="\"%EXE_PATH_REG%\" \"%%1\""
echo.
echo ; 添加到文件夹背景右键菜单
echo [HKEY_CLASSES_ROOT\Directory\Background\shell\TabExplorer]
echo @="open TabExplorer here"
echo "Icon"="%EXE_PATH_REG%"
echo.
echo [HKEY_CLASSES_ROOT\Directory\Background\shell\TabExplorer\command]
echo @="\"%EXE_PATH_REG%\" \"%%V\""
echo.
echo ; 添加到驱动器右键菜单
echo [HKEY_CLASSES_ROOT\Drive\shell\TabExplorer]
echo @="open with TabExplorer"
echo "Icon"="%EXE_PATH_REG%"
echo.
echo [HKEY_CLASSES_ROOT\Drive\shell\TabExplorer\command]
echo @="\"%EXE_PATH_REG%\" \"%%1\""
) > "%REG_FILE%"

REM 导入注册表
regedit /s "%REG_FILE%"

if %errorlevel% equ 0 (
    echo.
    echo ======================================
    echo 注册成功！
    echo ======================================
    echo.
    echo 现在可以：
    echo 1. 右键点击任意文件夹 - 选择"open with TabExplorer"
    echo 2. 右键点击任意文件 - 选择"open with TabExplorer"（打开文件所在文件夹）
    echo 3. 在文件夹空白处右键 - 选择"open TabExplorer here"
    echo 4. 右键点击驱动器 - 选择"open with TabExplorer"
    echo.
) else (
    echo.
    echo 注册失败！请以管理员身份运行此脚本
    echo.
)

REM 清理临时文件
del "%REG_FILE%"

pause
