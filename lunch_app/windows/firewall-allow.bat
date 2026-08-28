@echo off
rem スマートフォンや他のPCから接続できるように、Windowsファイアウォールを開ける
rem （管理者権限が必要なため、必要に応じて自動で昇格します）

net session >nul 2>&1
if errorlevel 1 (
  echo 管理者権限が必要です。確認画面が出たら「はい」を押してください。
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)

set "RULE=昼食注文システム (TCP 5002)"

echo ============================================================
echo   ファイアウォールの設定
echo ============================================================
echo.
echo 社内ネットワークからポート 5002 への接続を許可します。

netsh advfirewall firewall delete rule name="%RULE%" >nul 2>&1
netsh advfirewall firewall add rule name="%RULE%" dir=in action=allow protocol=TCP localport=5002 profile=private,domain
if errorlevel 1 (
  echo.
  echo [エラー] 設定に失敗しました。
  pause
  exit /b 1
)

echo.
echo ============================================================
echo   設定しました。
echo   スマートフォンから接続できるようになります。
echo ============================================================
echo.
echo   ※ つながらない場合は、この PC のネットワークが
echo      「プライベート ネットワーク」になっているか確認してください。
echo      （設定 ^> ネットワークとインターネット ^> Wi-Fi / イーサネット）
echo.
pause
