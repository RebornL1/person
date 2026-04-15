@echo off
REM Windows 启动脚本
REM 双击此文件即可启动服务

cd /d "%~dp0"
python start.py
pause