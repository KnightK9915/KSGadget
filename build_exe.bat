@echo off
chcp 65001 > nul
echo ========================================================
echo  コメントシート集計ツール ビルドスクリプト
echo  Comment Sheet Aggregator Build Script
echo ========================================================
echo.
echo EXEファイルの作成を開始します...
echo Building EXE file...
echo.

:: Check if PyInstaller is installed
pyinstaller --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] PyInstallerが見つかりません。
    echo PyInstaller not found. Installing...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [ERROR] PyInstallerのインストールに失敗しました。Pythonがインストールされているか確認してください。
        echo Failed to install PyInstaller. Please check Python installation.
        pause
        exit /b
    )
)

:: Run PyInstaller
echo.
echo PyInstallerを実行中... (Running PyInstaller...)
pyinstaller --onefile --noconsole --name "コメントシート集計ツール" --clean --paths=src --hidden-import=xlrd --hidden-import=aggregator src/gui_app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] ビルドに失敗しました。(Build Failed)
    pause
    exit /b
)

:: Move EXE to root
echo.
echo EXEファイルを移動中... (Moving EXE...)
move /Y "dist\コメントシート集計ツール.exe" . > nul

:: Cleanup
echo 一時ファイルを削除中... (Cleaning up...)
rmdir /s /q build
rmdir /s /q dist
del /q "コメントシート集計ツール.spec"

echo.
echo ========================================================
echo  ビルド完了！ (Build Complete!)
echo  作成されたファイル: コメントシート集計ツール.exe
echo ========================================================
echo.
pause
