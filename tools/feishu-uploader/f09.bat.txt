@echo off
chcp 65001 >nul
pushd "%~dp0.."
title 照片分組：收件匣的照片依黑畫面分組
echo ===== 照片分組：收件匣的照片依黑畫面分組 =====
echo.
node upload.js photos
echo.
echo ===== 跑完了，可以關掉這個視窗 =====
pause
