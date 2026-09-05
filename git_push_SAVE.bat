@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   Git 一鍵推送 (SAVE)
echo ============================================
echo.

REM 偵測當前分支 (不寫死 main)
for /f "tokens=*" %%b in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set "BRANCH=%%b"
if not defined BRANCH (
    echo [錯誤] 此資料夾不是 git repo, 或 git 未安裝
    pause
    exit /b 1
)
echo 當前分支: !BRANCH!
echo.

echo [1/4] 本地變更狀態:
echo --------------------------------------------
git status --short
echo --------------------------------------------
echo.

REM 用 for /f 直接判斷有無變更, 不依賴 find (避免被 MSYS Unix find 影子化)
set HAS_CHANGES=
for /f "delims=" %%c in ('git status --porcelain') do set HAS_CHANGES=1
if not defined HAS_CHANGES (
    echo [資訊] 沒有本地變更, 只推送已 commit 的紀錄
    goto PUSH_ONLY
)

REM 產生 ISO 格式預設訊息 (避免 Windows 地區「上午/下午」格式)
for /f "tokens=*" %%i in ('powershell -nop -c "(Get-Date).ToString('yyyy-MM-dd HH:mm')"') do set "DT=%%i"
set "DEFAULT_MSG=update !DT!"

echo 預設 commit 訊息: !DEFAULT_MSG!
echo (注意: 請勿在訊息中使用雙引號 ", 會導致 commit 失敗)
set /p COMMIT_MSG="請輸入 commit 訊息 (直接按 Enter 用預設): "
if "!COMMIT_MSG!"=="" set "COMMIT_MSG=!DEFAULT_MSG!"

echo.
echo [2/4] git add -A ...
git add -A
if errorlevel 1 (
    echo [錯誤] git add 失敗
    pause
    exit /b 1
)

echo [3/4] git commit ...
git commit -m "!COMMIT_MSG!"
if errorlevel 1 (
    echo [資訊] commit 失敗或無變更可 commit
)

:PUSH_ONLY
echo.
echo [4/4] git push origin !BRANCH! ...
git push origin !BRANCH!
if errorlevel 1 (
    echo.
    echo [錯誤] 推送失敗, 常見原因:
    echo   1. 遠端有別人新 commit, 你的版本落後
    echo   2. 網路 / 權限問題 / 認證過期
    echo.
    set /p RETRY="要自動 git pull --rebase 後再 push 嗎? (y/N): "
    if /i "!RETRY!"=="y" (
        echo.
        echo 執行 git pull --rebase origin !BRANCH! ...
        git pull --rebase origin !BRANCH!
        if errorlevel 1 (
            echo [錯誤] rebase 失敗, 請手動處理衝突後再執行
            pause
            exit /b 1
        )
        echo.
        echo 重新 push...
        git push origin !BRANCH!
        if errorlevel 1 (
            echo [錯誤] push 仍失敗
            pause
            exit /b 1
        )
    ) else (
        pause
        exit /b 1
    )
)

echo.
echo ============================================
echo   推送完成! 分支 !BRANCH! 已同步到 origin
echo ============================================
echo.
pause
