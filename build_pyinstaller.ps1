Write-Host "============================================"
Write-Host "    Excel工具 - PyInstaller打包脚本"
Write-Host "============================================"
Write-Host

# 检查Python是否安装
try {
    python --version | Out-Null
} catch {
    Write-Host "[错误] 未检测到Python，请先安装Python 3.11或3.12版本"
    Read-Host "按Enter键退出"
    exit 1
}

# 检查pyinstaller是否安装
try {
    pip show pyinstaller | Out-Null
} catch {
    Write-Host "[信息] 正在安装PyInstaller..."
    pip install pyinstaller
    Write-Host
}

# 检查icon.ico是否存在
$iconFlag = ""
if (Test-Path "icon.ico") {
    $iconFlag = "--icon=icon.ico"
} else {
    Write-Host "[警告] 未找到icon.ico文件，将使用默认图标"
}

Write-Host "[信息] 开始打包..."
Write-Host

# PyInstaller打包命令
pyinstaller `
    --onefile `
    --windowed `
    $iconFlag `
    --name=ExcelTools `
    --distpath=dist `
    main_qt6.py

if ($LASTEXITCODE -ne 0) {
    Write-Host
    Write-Host "[错误] 打包失败，请检查上方错误信息"
    Read-Host "按Enter键退出"
    exit 1
}

Write-Host
Write-Host "============================================"
Write-Host "    打包完成！"
Write-Host "============================================"
Write-Host "[信息] 输出文件: dist\ExcelTools.exe"
if (Test-Path "dist\ExcelTools.exe") {
    $fileInfo = Get-Item "dist\ExcelTools.exe"
    Write-Host "[信息] 文件大小: $($fileInfo.Length) 字节"
}
Write-Host
Read-Host "按Enter键退出"