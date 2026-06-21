<#
.SYNOPSIS
    一键打包「Zemax 公差分析」GUI 为 onedir EXE（可复用）。

.DESCRIPTION
    封装：环境检查 -> 清理旧产物 -> 校验/安装 PyInstaller -> 调用 gui.spec 构建
    -> 校验产物。务必在 IDE 自带的真实终端（PowerShell）运行，不要在 sandbox。

.PARAMETER Clean
    构建前删除 build\ 与 dist\（默认开启）。传 -Clean:$false 可保留增量缓存。

.PARAMETER Run
    构建成功后立即启动生成的 EXE。

.EXAMPLE
    .\build.ps1
    .\build.ps1 -Run
    .\build.ps1 -Clean:$false
#>
[CmdletBinding()]
param(
    [bool]$Clean = $true,
    [switch]$Run,
    [string]$PythonPath
)

$ErrorActionPreference = "Stop"

# 始终以脚本所在目录（tolerance_analysis）为工作目录
$ProjDir = $PSScriptRoot
Set-Location $ProjDir

# Python 解释器：优先用 -PythonPath 指定，否则用工作区共享 venv
if ($PythonPath) {
    $Python = $PythonPath
} else {
    $Python = Join-Path $ProjDir "..\.venv\Scripts\python.exe"
}
$Spec     = Join-Path $ProjDir "gui.spec"
$AppName  = "公差分析"
$DistDir  = Join-Path $ProjDir ("dist\" + $AppName)
$ExePath  = Join-Path $DistDir ($AppName + ".exe")
$PipIndex = "https://pypi.tuna.tsinghua.edu.cn/simple"

function Write-Step([string]$msg) {
    Write-Host "==> $msg" -ForegroundColor Cyan
}

# 1) 环境检查
Write-Step "检查 Python 解释器"
if (-not (Test-Path $Python)) {
    throw "未找到 Python 解释器：$Python`n请确认工作区 .venv 已创建，或用 -PythonPath 指定解释器，例如：`n  .\build.ps1 -PythonPath C:\path\to\python.exe"
}
& $Python --version

if (-not (Test-Path $Spec)) {
    throw "未找到打包配置：$Spec"
}

# 2) 校验 PyInstaller
Write-Step "校验 PyInstaller"
& $Python -m PyInstaller --version 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "未检测到 PyInstaller，脚本不会自动安装，请确认后手动执行：" -ForegroundColor Red
    Write-Host ("  " + $Python + " -m pip install pyinstaller -i " + $PipIndex) -ForegroundColor Yellow
    throw "缺少 PyInstaller。请手动安装后重新运行本脚本。"
}

# 3) 校验运行时关键依赖（pythonnet 的导入名是 clr）
Write-Step "校验运行时依赖（clr / PySide6 / openpyxl / numpy）"
$deps = @(
    @{ Import = "clr";      Pip = "pythonnet" },
    @{ Import = "PySide6";  Pip = "PySide6"   },
    @{ Import = "openpyxl"; Pip = "openpyxl"  },
    @{ Import = "numpy";    Pip = "numpy"     }
)
$missing = @()
foreach ($d in $deps) {
    & $Python -c ("import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('" + $d.Import + "') else 1)")
    if ($LASTEXITCODE -ne 0) {
        Write-Host ("  缺失：" + $d.Import + "（pip 包名：" + $d.Pip + "）") -ForegroundColor Yellow
        $missing += $d.Pip
    }
}
if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "检测到缺失依赖，脚本不会自动安装，请确认后手动执行：" -ForegroundColor Red
    Write-Host ("  " + $Python + " -m pip install " + ($missing -join " ") + " -i " + $PipIndex) -ForegroundColor Yellow
    throw ("缺少依赖：" + ($missing -join ", ") + "。请手动安装后重新运行本脚本。")
}

# 4) 清理旧产物
if ($Clean) {
    Write-Step "清理旧产物 build\ dist\"
    foreach ($dir in @((Join-Path $ProjDir "build"), (Join-Path $ProjDir "dist"))) {
        if (Test-Path $dir) {
            Remove-Item -Recurse -Force $dir -ErrorAction SilentlyContinue
        }
    }
}

# 5) 构建
Write-Step "开始打包（onedir）"
& $Python -m PyInstaller --noconfirm $Spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 构建失败（退出码 $LASTEXITCODE）。" }

# 6) 校验产物
Write-Step "校验产物"
if (-not (Test-Path $ExePath)) {
    throw "构建结束但未找到 EXE：$ExePath"
}
Write-Host ""
Write-Host "构建成功 ✔" -ForegroundColor Green
Write-Host "EXE 路径：$ExePath" -ForegroundColor Green
Get-ChildItem $DistDir | Where-Object { -not $_.PSIsContainer } |
    Select-Object Name, @{N="Size(KB)";E={[math]::Round($_.Length/1KB,1)}} |
    Format-Table -AutoSize

Write-Host "提示：目标电脑仍需安装 Zemax OpticStudio 并有可用 license。" -ForegroundColor Yellow

# 7) 可选：立即运行
if ($Run) {
    Write-Step "启动 EXE"
    Start-Process -FilePath $ExePath
}
