# 启动 CAD 翻译 GUI 程序的 PowerShell 脚本
# 自动设置 Tcl/Tk 环境变量

Write-Host "正在启动 CAD 翻译 GUI 程序..." -ForegroundColor Green

# 设置 Tcl/Tk 环境变量
$env:TCL_LIBRARY = "c:\Users\zhenhe\OneDrive\永盛\翻译\cad code\.conda\Library\lib\tcl8.6"
$env:TK_LIBRARY = "c:\Users\zhenhe\OneDrive\永盛\翻译\cad code\.conda\Library\lib\tk8.6"

Write-Host "环境变量已设置:" -ForegroundColor Yellow
Write-Host "TCL_LIBRARY = $env:TCL_LIBRARY" -ForegroundColor Cyan
Write-Host "TK_LIBRARY = $env:TK_LIBRARY" -ForegroundColor Cyan

# 启动程序
try {
    & ".conda/python.exe" gui.py
    Write-Host "程序已正常退出。" -ForegroundColor Green
}
catch {
    Write-Host "程序启动失败: $_" -ForegroundColor Red
}

Write-Host "按任意键退出..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")