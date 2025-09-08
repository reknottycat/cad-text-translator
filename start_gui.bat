@echo off
REM 启动 CAD 翻译 GUI 程序的批处理脚本
REM 自动设置 Tcl/Tk 环境变量

echo 正在启动 CAD 翻译 GUI 程序...

REM 设置 Tcl/Tk 环境变量
set TCL_LIBRARY=c:\Users\zhenhe\OneDrive\永盛\翻译\cad code\.conda\Library\lib\tcl8.6
set TK_LIBRARY=c:\Users\zhenhe\OneDrive\永盛\翻译\cad code\.conda\Library\lib\tk8.6

REM 启动程序
".conda\python.exe" gui.py

echo 程序已退出。
pause