@echo off
chcp 65001 >nul
pushd "%~dp0.."
title 拍照順序：照這個順序拍，每件拍完拍一張全黑
echo ===== 拍照順序：照這個順序拍，每件拍完拍一張全黑 =====
echo.
node upload.js photos --order
echo.
echo ===== 跑完了，可以關掉這個視窗 =====
pause
