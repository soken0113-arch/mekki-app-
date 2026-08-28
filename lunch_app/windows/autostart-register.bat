@echo off
rem Windows の起動時に昼食注文システムが自動で立ち上がるように登録する
set "TARGET=%~dp0start.bat"
set "WORKDIR=%~dp0"
set "SHORTCUT=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\昼食注文システム.lnk"

powershell -NoProfile -Command ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('%SHORTCUT%');" ^
  "$s.TargetPath = '%TARGET%';" ^
  "$s.WorkingDirectory = '%WORKDIR%';" ^
  "$s.Description = '社内昼食注文システム';" ^
  "$s.Save()"

if errorlevel 1 (
  echo [エラー] 登録に失敗しました。
  pause
  exit /b 1
)

echo.
echo Windows の起動時に自動で立ち上がるように登録しました。
echo.
echo 解除したいときは、次のフォルダの中の
echo 「昼食注文システム」のショートカットを削除してください。
echo   %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
echo.
pause
