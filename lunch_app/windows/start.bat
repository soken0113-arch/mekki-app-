@echo off
rem 昼食注文システムを起動する
cd /d "%~dp0.."

if not exist ".venv\Scripts\python.exe" (
  echo [エラー] セットアップがまだ済んでいません。
  echo 先に windows フォルダの「setup.bat」を実行してください。
  pause
  exit /b 1
)

".venv\Scripts\python.exe" run_server.py
echo.
echo サーバーが停止しました。
pause
