@echo off
rem 昼食注文システム セットアップ（最初に1回だけ実行）
cd /d "%~dp0.."
echo ============================================================
echo   昼食注文システム セットアップ
echo ============================================================
echo.

where py >nul 2>&1
if errorlevel 1 (
  echo [エラー] Python が見つかりませんでした。
  echo.
  echo   https://www.python.org/downloads/windows/
  echo   から Python をインストールしてください。
  echo   インストール画面の「Add python.exe to PATH」に
  echo   必ずチェックを入れてください。
  echo.
  pause
  exit /b 1
)
echo Python を確認しました。
echo.

if not exist ".venv\Scripts\python.exe" (
  echo 実行環境を作成しています...
  py -3 -m venv .venv
  if errorlevel 1 (
    echo [エラー] 実行環境の作成に失敗しました。
    pause
    exit /b 1
  )
)

echo 必要な部品をインストールしています（数分かかります）...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo [エラー] インストールに失敗しました。
  echo インターネットに接続されているか確認して、もう一度実行してください。
  pause
  exit /b 1
)

echo.
echo ============================================================
echo   セットアップが完了しました。
echo   次は「start.bat」をダブルクリックして起動してください。
echo ============================================================
pause
