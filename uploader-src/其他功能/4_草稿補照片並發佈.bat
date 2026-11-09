@echo off
chcp 65001 >nul
pushd "%~dp0.."
title 草稿補照片並發佈
echo ===== 草稿補照片並發佈 =====
echo.
node upload.js drafts
echo.
echo ===== 跑完了，可以關掉這個視窗 =====
pause
