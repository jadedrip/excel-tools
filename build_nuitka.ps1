Write-Host "============================================"
Write-Host "    Excel工具 - Nuitka打包脚本"
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

# 检查nuitka是否安装
try {
    pip show nuitka | Out-Null
} catch {
    Write-Host "[信息] 正在安装Nuitka..."
    pip install nuitka
    Write-Host
}

# 检查icon.ico是否存在
$iconFlag = ""
if (Test-Path "icon.ico") {
    $iconFlag = "--windows-icon-from-ico=icon.ico"
} else {
    Write-Host "[警告] 未找到icon.ico文件，将使用默认图标"
}

Write-Host "[信息] 开始打包..."
Write-Host

# Nuitka打包命令
python -m nuitka `
    --onefile `
    --windows-disable-console `
    --enable-plugins=pyqt6 `
    --include-package=pandas `
    --include-package=numpy `
    --include-package=openpyxl `
    $iconFlag `
    --output-dir=dist `
    --assume-yes-for-downloads `
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
Write-Host "[信息] 输出文件: dist\main_qt6.exe"
if (Test-Path "dist\main_qt6.exe") {
    $fileInfo = Get-Item "dist\main_qt6.exe"
    Write-Host "[信息] 文件大小: $($fileInfo.Length) 字节"
}
Write-Host
Read-Host "按Enter键退出"