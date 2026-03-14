@echo off
setlocal EnableExtensions
chcp 65001 >nul

REM ===============================
REM 项目根目录（build.bat 所在目录）
REM ===============================
set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"

REM ===============================
REM 基本参数
REM ===============================
set "APP_NAME=MyApp"
set "ENTRY_PY=%ROOT%\run.py"
set "VBS_FILE=%ROOT%\run_ai.vbs"
set "JSX_FILE=%ROOT%\AItest_ai.jsx"
if not exist "%JSX_FILE%" set "JSX_FILE=%ROOT%\AItest.jsx"

set "ICON_FILE=%ROOT%\resources\icon.ico"

set "DIST_DIR=%ROOT%\dist"
set "BUILD_DIR=%ROOT%\build"
set "SPEC_DIR=%ROOT%\_spec"
set "CONTENTS_DIR=_internal"

REM ===============================
REM 可选：签名开关
REM 1=启用签名，0=不签名
REM 如果不用签名，保持 0 即可
REM ===============================
set "ENABLE_SIGN=0"

REM 证书方式1：PFX 文件
set "SIGN_PFX="
set "SIGN_PFX_PASSWORD="

REM 时间戳服务器
set "SIGN_TIMESTAMP_URL=http://timestamp.digicert.com"

echo ========================================
echo ROOT      : %ROOT%
echo ENTRY     : %ENTRY_PY%
echo VBS       : %VBS_FILE%
echo JSX       : %JSX_FILE%
echo ICON      : %ICON_FILE%
echo DIST      : %DIST_DIR%
echo BUILD     : %BUILD_DIR%
echo SPEC      : %SPEC_DIR%
echo ========================================
echo.

REM ===============================
REM 检查文件
REM ===============================
if not exist "%ENTRY_PY%" (
    echo [ERROR] 找不到 run.py
    echo         %ENTRY_PY%
    pause
    exit /b 1
)

if not exist "%VBS_FILE%" (
    echo [ERROR] 找不到 run_ai.vbs
    echo         %VBS_FILE%
    pause
    exit /b 1
)

if not exist "%JSX_FILE%" (
    echo [ERROR] 找不到 JSX 文件
    echo         已检查:
    echo         %ROOT%\AItest_ai.jsx
    echo         %ROOT%\AItest.jsx
    pause
    exit /b 1
)

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] 当前环境找不到 python
    pause
    exit /b 1
)

python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [INFO] 未安装 PyInstaller，开始安装...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] PyInstaller 安装失败
        pause
        exit /b 1
    )
)

REM ===============================
REM 清理旧输出
REM ===============================
if exist "%DIST_DIR%\%APP_NAME%" (
    echo [INFO] 删除旧目录: %DIST_DIR%\%APP_NAME%
    rmdir /s /q "%DIST_DIR%\%APP_NAME%"
)

if exist "%BUILD_DIR%" (
    echo [INFO] 删除旧目录: %BUILD_DIR%
    rmdir /s /q "%BUILD_DIR%"
)

if exist "%SPEC_DIR%" (
    echo [INFO] 删除旧目录: %SPEC_DIR%
    rmdir /s /q "%SPEC_DIR%"
)

echo.
echo [INFO] 开始打包 %APP_NAME% ...
echo.

REM ===============================
REM 有图标 / 无图标分开写，避免 cmd 引号变量出错
REM ===============================
if exist "%ICON_FILE%" (
    python -m PyInstaller ^
      "%ENTRY_PY%" ^
      --noconfirm ^
      --clean ^
      --onedir ^
      --name "%APP_NAME%" ^
      --distpath "%DIST_DIR%" ^
      --workpath "%BUILD_DIR%" ^
      --specpath "%SPEC_DIR%" ^
      --contents-directory "%CONTENTS_DIR%" ^
      --add-data "%VBS_FILE%;." ^
      --add-data "%JSX_FILE%;." ^
      --hidden-import get_best ^
      --noupx ^
      --icon "%ICON_FILE%"
) else (
    python -m PyInstaller ^
      "%ENTRY_PY%" ^
      --noconfirm ^
      --clean ^
      --onedir ^
      --name "%APP_NAME%" ^
      --distpath "%DIST_DIR%" ^
      --workpath "%BUILD_DIR%" ^
      --specpath "%SPEC_DIR%" ^
      --contents-directory "%CONTENTS_DIR%" ^
      --add-data "%VBS_FILE%;." ^
      --add-data "%JSX_FILE%;." ^
      --hidden-import get_best ^
      --noupx
)

if errorlevel 1 (
    echo.
    echo [ERROR] 打包失败
    pause
    exit /b 1
)

echo.
echo [OK] 打包完成
echo 输出目录: %DIST_DIR%\%APP_NAME%
echo.

REM ===============================
REM 可选：签名
REM ===============================
if "%ENABLE_SIGN%"=="1" (
    where signtool >nul 2>&1
    if errorlevel 1 (
        echo [WARN] 未找到 signtool，跳过签名
    ) else (
        if not "%SIGN_PFX%"=="" (
            if exist "%DIST_DIR%\%APP_NAME%\%APP_NAME%.exe" (
                echo [INFO] 正在签名 EXE...
                signtool sign /fd SHA256 /f "%SIGN_PFX%" /p "%SIGN_PFX_PASSWORD%" /tr "%SIGN_TIMESTAMP_URL%" /td SHA256 "%DIST_DIR%\%APP_NAME%\%APP_NAME%.exe"
            )
        ) else (
            echo [WARN] ENABLE_SIGN=1，但未设置 SIGN_PFX，跳过签名
        )
    )
)

echo [INFO] 生成文件如下:
dir /b "%DIST_DIR%\%APP_NAME%"
echo.
if exist "%DIST_DIR%\%APP_NAME%\%CONTENTS_DIR%" (
    echo [INFO] 内部资源文件如下:
    dir /b "%DIST_DIR%\%APP_NAME%\%CONTENTS_DIR%"
)

echo.
echo [TIP] run.py 中请用 __file__ 定位资源，不要写死相对路径
echo.
pause
exit /b 0