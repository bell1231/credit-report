@echo off
chcp 65001 >nul
cd /d "D:\CodeBuddy Output\credit-report"

echo ============================================
echo   GitHub 仓库初始化 & 推送
echo ============================================
echo.

REM 配置 Git 用户信息（首次使用需要）
git config --global user.name "bell1231"
git config --global user.email "bell1231@users.noreply.github.com"

echo [1/4] 初始化 Git 仓库...
git init

echo [2/4] 添加所有文件...
git add -A

echo [3/4] 提交...
git commit -m "Initial commit: 企业授信报告生成系统"

echo [4/4] 推送到 GitHub...
git remote add origin https://github.com/bell1231/credit-report.git
git branch -M main
git push -u origin main

echo.
echo ============================================
echo   完成！请检查上方输出确认推送成功
echo ============================================
pause
