@echo off
chcp 65001 >nul
title 一鍵同步 Google 雲端試算表並開啟工具
echo 正在連線 Google 雲端試算表抓取最新資料...
python sync_cloud.py
echo 正在為您開啟家庭照顧與家用花費紀錄工具...
start "" "%~dp0index.html"
exit
