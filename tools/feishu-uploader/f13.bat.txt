@echo off
chcp 65001 >nul
pushd "%~dp0"
title 待辦清單：哪些還沒做完
echo ===== 待辦清單：哪些還沒做完 =====
echo.
node upload.js todo
echo.
echo ===== 跑完了，可以關掉這個視窗 =====
pause
