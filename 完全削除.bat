@echo off
chcp 65001 > nul
echo ============================================================
echo   Wi-Fi 監視システム　完全削除
echo ============================================================
echo.
echo 以下のデータを削除します：
echo   1. C:\WiFiMonitor\（プログラム・ログ・ビューア）
echo   2. タスクスケジューラ「WiFi監視」
echo   3. speedtest-cli（Pythonパッケージ）
echo.
set /p CONFIRM=本当に削除しますか？ [Y/N]: 
if /i not "%CONFIRM%"=="Y" (
    echo.
    echo キャンセルしました。
    pause
    exit /b
)

echo.
echo [1/3] C:\WiFiMonitor\ を削除中...
if exist "C:\WiFiMonitor\" (
    rmdir /s /q "C:\WiFiMonitor\"
    echo       完了
) else (
    echo       フォルダが見つかりませんでした（スキップ）
)

echo [2/3] タスクスケジューラから「WiFi監視」を削除中...
schtasks /delete /tn "WiFi監視" /f > nul 2>&1
if %errorlevel%==0 (
    echo       完了
) else (
    echo       タスクが見つかりませんでした（スキップ）
)

echo [3/3] speedtest-cli を削除中...
"%LocalAppData%\Programs\Python\Python314\python.exe" -m pip uninstall speedtest-cli -y > nul 2>&1
if %errorlevel%==0 (
    echo       完了
) else (
    echo       speedtest-cli が見つかりませんでした（スキップ）
)

echo.
echo ============================================================
echo   アンインストール完了
echo   このバッチファイル自体は手動で削除してください
echo ============================================================
echo.
pause
