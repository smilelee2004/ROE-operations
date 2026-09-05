@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   Git 一鍵拉取最新 (LOAD)
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

echo [1/3] 檢查本地是否有未提交變更...
git diff --quiet HEAD 2>nul
if errorlevel 1 (
    echo [警告] 本地有未 commit 的變更, pull 可能會衝突
    echo --------------------------------------------
    git status --short
    echo --------------------------------------------
    set /p CONFIRM="要繼續 pull 嗎? (y/N): "
    if /i not "!CONFIRM!"=="y" (
        echo 已取消
        pause
        exit /b 0
    )
) else (
    echo (沒有未提交變更^)
)

echo.
echo [2/3] 從 origin 拉取最新 (分支: !BRANCH!)...
git pull origin !BRANCH!
if errorlevel 1 (
    echo.
    echo [錯誤] 拉取失敗, 可能原因:
    echo   - 本地有衝突的變更: 先 commit 或 git stash
    echo   - 網路 / 權限問題
    pause
    exit /b 1
)

echo.
echo [3/3] 最新 3 個 commit:
git log --oneline -3

echo.
echo ============================================
echo   拉取完成! 已更新至 !BRANCH! 最新版本
echo ============================================
echo.
pause
