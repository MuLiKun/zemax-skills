---
name: "zemax-zosapi-connector"
description: "Connects to local Zemax OpticStudio via ZOS-API. Invoke when user has an open Zemax file, asks to connect, analyze, optimize, or operate Zemax."
---

# Zemax ZOS-API 连接器

当用户需要连接本地 Zemax OpticStudio 实例、操作已打开的 `.zmx`/`.zos` 文件、打开点列图或 FFT MTF 等分析、读取镜头数据、优化或导出 Zemax 结果时，使用本技能。

本技能设计为可跨工作区和电脑迁移。除非用户确认适用于当前机器，否则把文档中任何具体的 Zemax 安装路径、文件路径、实例号和分析结果都视为示例或此前缓存的结果。

## 可迁移连接原则

1. 不要假设 Zemax 安装目录是固定的。
2. 不要假设用户打开的文件是固定的。
3. 当用户希望看到可见的 GUI 操作时，优先通过交互扩展（Interactive Extension）连接当前打开的 GUI 文件。
4. 在加载 ZOS-API 程序集之前，于运行时发现或确认路径。
5. 若此前已知路径有效，为提高效率直接复用；若失败，再搜索常见位置或询问用户安装目录。

## 需在运行时确认的输入

连接前，确认以下值：

- `$zosDir`：OpticStudio 安装目录，包含 `ZOSAPI_NetHelper.dll`、`ZOSAPI_Interfaces.dll`、`ZOSAPI.dll`。
- `$instance`：Zemax 交互扩展对话框中显示的实例号，通常为 `1`。
- `$expectedFile`：用户期望的文件路径（可选）。GUI 附着操作时，连接后用 `$sys.SystemFile` 核对，而不要硬编码。
- `$targetFreq`：用户请求的 MTF 频率（可选），例如 `67.0` lp/mm。

## 来自某台已验证机器的示例值

以下仅为示例值。若在其他工作区或电脑上失败，不要把它们当作通用默认值硬编码。

- 示例 OpticStudio 安装目录：
  `C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00`
- 示例主程序：
  `C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00\OpticStudio.exe`
- 示例 ZOS-API 程序集：
  - `C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00\ZOSAPI_NetHelper.dll`
  - `C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00\ZOSAPI_Interfaces.dll`
  - `C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00\ZOSAPI.dll`
- 示例用户文件：
  `F:\个人文件\P3111\0530.zmx`
- 示例交互扩展实例号：
  `1`
- 示例已确认授权：
  `PremiumEdition`

## 性能说明（为什么连接会感觉慢）

每次“连接并运行”都会重复付出固定开销。优化以下几点：

1. 每次运行都重新加载 .NET 程序集。对三个 DLL 执行 `Add-Type` 加上 `Initialize($zosDir)`，每次都要花几秒。缓存 `$zosDir`，已知时跳过发现。
2. 全盘递归搜索。对 `C:\Program Files` 执行 `Get-ChildItem -Recurse` 可能耗时数十秒。仅作为最后手段并限制深度。
3. 每条命令一个新的 PowerShell 进程。程序集加载 + 握手的开销会重复。对密集、连续的操作，使用一个常驻 PowerShell 会话（见“常驻会话”），而不是每次都重新连接。
4. MTF 采样过高。`S_256x256` 较慢；对 67 lp/mm 这类单一频率看趋势，`S_64x64` 已足够。仅在需要高精度时才提高采样。
5. 终端回显开销。优先把结果写入临时文件再读取，而不是大量 `Write-Host` 输出，以避免截断和缓慢回显。

快路径规则：若 `$zosDir` 已缓存且有效，直接进入程序集加载和 `ConnectAsExtension($instance)`，不要再运行发现流程。

## 安装目录发现

仅当 `$zosDir` 未知或缓存路径失效时，才使用此发现流程：

1. 先检查缓存/已知路径。
2. 检查一小份常见 Ansys/Zemax 安装位置清单。
3. 仅当上述失败时，在 `C:\Program Files`、`C:\Program Files (x86)`、`C:\Program Files\ANSYS Inc` 下做限定深度的搜索。
4. 若发现多个版本，优先使用用户请求的版本，或询问使用哪个版本。

优化版 PowerShell 发现模板（限定深度，无全盘递归）：

```powershell
function Resolve-ZosDir {
  param([string]$Cached)
  if($Cached -and (Test-Path (Join-Path $Cached 'ZOSAPI.dll'))){ return $Cached }
  $candidateDirs = @(
    'C:\Program Files\Ansys Zemax OpticStudio 2023 R1.00',
    'C:\Program Files\Ansys Zemax OpticStudio 2024 R1.00',
    'C:\Program Files\Ansys Zemax OpticStudio 2024 R2.00',
    'C:\Program Files\Ansys Zemax OpticStudio 2025 R1.00',
    'C:\Program Files\Zemax OpticStudio'
  )
  foreach($dir in $candidateDirs){
    if((Test-Path (Join-Path $dir 'ZOSAPI_NetHelper.dll')) -and (Test-Path (Join-Path $dir 'ZOSAPI_Interfaces.dll')) -and (Test-Path (Join-Path $dir 'ZOSAPI.dll'))){ return $dir }
  }
  # 限定深度的兜底：仅扫描直接子目录，不做全盘递归。
  $roots = @('C:\Program Files','C:\Program Files (x86)','C:\Program Files\ANSYS Inc') | Where-Object { Test-Path $_ }
  foreach($root in $roots){
    $sub = Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Zemax|OpticStudio' }
    foreach($s in $sub){
      $d=$s.FullName
      if((Test-Path (Join-Path $d 'ZOSAPI_NetHelper.dll')) -and (Test-Path (Join-Path $d 'ZOSAPI.dll'))){ return $d }
    }
  }
  return $null
}
```

解析成功一次后，持久化该值以便复用：写入当前会话的 `$env:ZOS_DIR`，或作为缓存路径记住，供后续调用使用。

## 优先连接策略

当用户说 Zemax 文件已打开且交互扩展已开启时，优先使用 `ConnectAsExtension($instance)`。这会直接连接当前 GUI 会话。

若用户希望看到可见的 GUI 操作，请他们在 OpticStudio 中打开：

```text
Programming / ZOS-API / Interactive Extension
```

或本地化的等效菜单。然后使用交互扩展对话框中显示的实例号。

若直接的扩展连接失败，报告确切错误并请用户保持 Zemax 交互扩展对话框打开。若他们只需文件操作而非 GUI 附着操作，使用 `CreateNewApplication()` 作为兜底，打开一个独立的、由 API 控制的 OpticStudio 会话。

## 当前 GUI 的 PowerShell 连接模板

通过终端使用 PowerShell。除非已确认当前机器上有 PythonNET，否则不要使用 Python。

```powershell
$ErrorActionPreference='Continue'
$instance=1
$expectedFile=$null

# 快路径：复用已缓存的 $zosDir（例如来自 $env:ZOS_DIR）以完全跳过发现。
if((-not $zosDir) -and $env:ZOS_DIR){ $zosDir = $env:ZOS_DIR }
$zosDir = Resolve-ZosDir -Cached $zosDir
if($null -eq $zosDir){ throw 'Set $zosDir to the OpticStudio installation directory containing ZOSAPI.dll files.' }
$env:ZOS_DIR = $zosDir

# 每个进程只加载一次程序集。若类型已存在则跳过重新加载。
if(-not ('ZOSAPI_NetHelper.ZOSAPI_Initializer' -as [type])){
  Add-Type -Path (Join-Path $zosDir 'ZOSAPI_NetHelper.dll') -PassThru | Out-Null
  [ZOSAPI_NetHelper.ZOSAPI_Initializer]::Initialize($zosDir) | Out-Null
  Add-Type -Path (Join-Path $zosDir 'ZOSAPI_Interfaces.dll') -PassThru | Out-Null
  Add-Type -Path (Join-Path $zosDir 'ZOSAPI.dll') -PassThru | Out-Null
}
$conn=New-Object ZOSAPI.ZOSAPI_Connection
$app=$conn.ConnectAsExtension($instance)
if($null -eq $app){ throw 'ConnectAsExtension returned null. Keep Zemax Interactive Extension open and confirm the instance number.' }
$sys=$app.PrimarySystem
if($null -eq $sys){ throw 'PrimarySystem is null.' }
Write-Host "Connected: $($sys.SystemFile)"
Write-Host "System: $($sys.SystemName), Mode: $($sys.Mode), Surfaces: $($sys.LDE.NumberOfSurfaces), Fields: $($sys.SystemData.Fields.NumberOfFields), Wavelengths: $($sys.SystemData.Wavelengths.NumberOfWavelengths)"
if($null -ne $expectedFile -and $expectedFile -ne '' -and $sys.SystemFile -ne $expectedFile){
  Write-Host "Warning: connected file differs from expected file. Expected: $expectedFile; Actual: $($sys.SystemFile)"
}
```

## 独立模式兜底策略

仅当用户不需要操作可见的 GUI 窗口，或交互扩展不可用时使用。

```powershell
$conn=New-Object ZOSAPI.ZOSAPI_Connection
$app=$conn.CreateNewApplication()
if($null -eq $app){ throw 'CreateNewApplication returned null.' }
$sys=$app.PrimarySystem
if($null -ne $targetFile -and $targetFile -ne ''){
  $sys.LoadFile($targetFile, $false)
}
Write-Host "Standalone connected: $($sys.SystemFile)"
```

## 常驻会话（避免每条命令都重新连接）

最大的重复开销是：为每一个操作都新开 PowerShell 进程、重新加载三个 .NET 程序集、并重新运行 `Initialize`。当用户请求多个连续操作（打开布局、点列图、MTF、读参数）时，保持同一个长生命周期的 PowerShell 终端，并复用已连接的 `$app` / `$sys`。

实际做法：

1. 在某个终端中运行一次连接模板。保持该终端存活。
2. 把后续每个分析片段（点列图、FFT MTF、参数）发到同一个终端。`$zosDir`、已加载的程序集、`$app`、`$sys` 仍在内存中，因此每个后续命令都会跳过发现、程序集加载和连接握手。
3. 复用 `$sys` 之前，防范会话已掉线：

```powershell
if($null -eq $app -or $null -eq $app.PrimarySystem){
  $conn=New-Object ZOSAPI.ZOSAPI_Connection
  $app=$conn.ConnectAsExtension($instance)
}
$sys=$app.PrimarySystem
if($null -eq $sys){ throw 'PrimarySystem is null. Keep the Interactive Extension dialog open and retry.' }
```

注意与限制：

- 终端环境可能终止长时间的保活循环（`while($true){ Start-Sleep }`）。不要依赖后台保活循环。而是保持终端会话本身打开，按需把命令发进去；会话在命令之间保持“温热”。
- 若会话丢失（终端关闭，或用户关闭了交互扩展），只有那时才通过 `ConnectAsExtension($instance)` 付出重新连接的开销。
- 对密集的数据提取，把结果写入临时文件再读回，而不是大量 `Write-Host` 输出，以避免终端截断和回显开销。

## 打开点列图

```powershell
$spot=$sys.Analyses.New_Analysis([ZOSAPI.Analysis.AnalysisIDM]::StandardSpot)
$spotSettings=$spot.GetSettings()
try { $spotSettings.Field.SetFieldNumber(0) } catch {}
try { $spotSettings.Wavelength.SetWavelengthNumber(0) } catch {}
try { $spotSettings.ReferTo=[ZOSAPI.Analysis.Settings.RMS.ReferTo]::Centroid } catch {}
$spot.ApplyAndWaitForCompletion() | Out-Null
$spotResults=$spot.GetResults()
$fieldCount=$sys.SystemData.Fields.NumberOfFields
$waveCount=$sys.SystemData.Wavelengths.NumberOfWavelengths
for($f=1; $f -le $fieldCount; $f++){
  for($w=1; $w -le $waveCount; $w++){
    $rms=$spotResults.SpotData.GetRMSSpotSizeFor($f,$w)
    $geo=$spotResults.SpotData.GetGeoSpotSizeFor($f,$w)
    Write-Host ('Field {0}, Wave {1}: RMS={2:N6}, GEO={3:N6}' -f $f,$w,$rms,$geo)
  }
}
```

## 在指定频率打开 FFT MTF

```powershell
$targetFreq=67.0
$mtf=$sys.Analyses.New_FftMtf()
$mtfSettings=$mtf.GetSettings()
try { $mtfSettings.MaximumFrequency = $targetFreq } catch {}
# S_64x64 速度快，对 67 lp/mm 这类单一频率趋势已足够。
# 仅在明确需要高精度时才提高到 S_128x128 或 S_256x256。
try { $mtfSettings.SampleSize = [ZOSAPI.Analysis.SampleSizes]::S_64x64 } catch {}
$mtf.ApplyAndWaitForCompletion() | Out-Null
$mtfResults=$mtf.GetResults()
function Get-InterpolatedValue($xs, $ys, [double]$xTarget) {
  if($xs.Count -eq 0){ return $null }
  for($i=0; $i -lt $xs.Count; $i++){
    if([Math]::Abs([double]$xs[$i]-$xTarget) -lt 1e-9){ return [double]$ys[$i] }
    if($i -gt 0 -and [double]$xs[$i] -ge $xTarget){
      $x0=[double]$xs[$i-1]; $x1=[double]$xs[$i]
      $y0=[double]$ys[$i-1]; $y1=[double]$ys[$i]
      if([Math]::Abs($x1-$x0) -lt 1e-12){ return $y1 }
      return $y0 + ($xTarget-$x0)*($y1-$y0)/($x1-$x0)
    }
  }
  return [double]$ys[$ys.Count-1]
}
for($seriesNum=0; $seriesNum -lt $mtfResults.NumberOfDataSeries; $seriesNum++){
  $series=$mtfResults.GetDataSeries($seriesNum)
  $xs=@($series.XData.Data)
  $yData=$series.YData.Data
  $dim0=$yData.GetLength(0); $dim1=$yData.GetLength(1)
  $tan=@(); $sag=@()
  if($dim0 -eq 2){
    for($i=0; $i -lt $dim1; $i++){
      $tan += [double]$yData.GetValue(0,$i)
      $sag += [double]$yData.GetValue(1,$i)
    }
  } else {
    for($i=0; $i -lt $dim0; $i++){
      $tan += [double]$yData.GetValue($i,0)
      if($dim1 -gt 1){ $sag += [double]$yData.GetValue($i,1) }
    }
  }
  $tanValue=Get-InterpolatedValue $xs $tan $targetFreq
  $sagValue=Get-InterpolatedValue $xs $sag $targetFreq
  Write-Host ('Series {0}: Tangential={1:N6}, Sagittal={2:N6}' -f $seriesNum,$tanValue,$sagValue)
}
```

## 来自某次此前会话的示例验证结果

这仅是 `F:\个人文件\P3111\0530.zmx` 在 `67 lp/mm` 下的示例结果。对新文件或新系统需重新计算。

```text
Connected file: F:\个人文件\P3111\0530.zmx
System: Microscope objective, table 1, 4953962
Mode: Sequential
Surfaces: 17
Fields: 10
Wavelengths: 3
License: PremiumEdition

Series 0: Tangential=0.675306, Sagittal=0.675306
Series 1: Tangential=0.835742, Sagittal=0.733399
Series 2: Tangential=0.873036, Sagittal=0.801491
Series 3: Tangential=0.814010, Sagittal=0.827353
Series 4: Tangential=0.785663, Sagittal=0.812535
Series 5: Tangential=0.735165, Sagittal=0.797863
Series 6: Tangential=0.715715, Sagittal=0.794808
Series 7: Tangential=0.744905, Sagittal=0.799986
Series 8: Tangential=0.750946, Sagittal=0.809058
Series 9: Tangential=0.691461, Sagittal=0.824717
```

## 重要操作注意事项

- 若用户希望直接操作可见的 GUI 文件，要求 Zemax 中保持交互扩展打开。
- 对手动打开的 GUI，`ConnectToApplication()` 可能失败并报：`This application was not launched by Optic Studio`。交互扩展应使用 `ConnectAsExtension($instance)`。
- 长时间保活循环可能被终端环境终止。若发生，后续每个操作用 `ConnectAsExtension($instance)` 重新连接。
- 未经明确要求，不要覆盖用户原始的 Zemax 文件。优先用带描述性名称的副本保存。
- 打开分析窗口时，优先使用 `New_Analysis(...)` 或便捷方法如 `New_FftMtf()`。
- 若 Zemax 安装路径变更，更新 `$zosDir` 或使用上面的发现流程。
- 若用户文件变更，除非需要否则不要改模板；先连接并核对 `$sys.SystemFile`。
- 若交互扩展实例号变更，把 `$instance` 更新为 Zemax 对话框中显示的号码。
