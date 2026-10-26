# reset_scanner.ps1
# 扫描仪软复位工具（本机实测有效）：用于扫描仪"卡死/脱机/被占用"时自愈，无需拔插 USB。
# 手段：①清理残留扫描进程 ②重启图像采集服务 stisvc（实测 ~2 秒完成）
# 注意：本机实测不能做"设备禁用/启用"——会触发未完成的设备拔出，把扫描仪卡死。
# 需要管理员权限；若当前不是管理员，会自动以提权方式重新启动自身（本机 UAC 策略为静默提权，不弹窗）。

param([string]$LogPath = "$env:TEMP\qj_scanner_reset.log")

$NAPS2_CONSOLE = 'D:\Tools\NAPS2\App\NAPS2.Console.exe'

function W($m) { Add-Content -Path $LogPath -Value "$(Get-Date -Format 'HH:mm:ss')  $m" -Encoding utf8 }

# ---- 权限自举 ----
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-LogPath', $LogPath
    ) -ErrorAction SilentlyContinue
    exit 0
}

W "=== 扫描仪软复位开始（已获得管理员权限）==="

# ---- 1) 清理残留扫描进程（它们常常正占用着扫描仪）----
foreach ($n in @('NAPS2.Console.exe', 'NAPS2.Worker.exe')) {
    try {
        $r = taskkill /F /T /IM $n 2>&1 | Out-String
        if ($r -match 'SUCCESS|成功') { W "已清理残留进程: $n" }
    } catch { }
}

# ---- 2) 定位扫描仪设备 ----
$dev = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
       Where-Object { $_.InstanceId -like 'USB\VID_3072*' -or $_.FriendlyName -match 'HUAGOSCAN' } |
       Select-Object -First 1
if ($dev) {
    W "设备: $($dev.FriendlyName)  复位前状态=$($dev.Status)"
} else {
    W "警告: 未找到华高扫描仪设备（可能 USB 未连接）"
}

# ---- 3) 主手段：重启图像采集服务（释放被占用的设备句柄）----
$job = Start-Job -ScriptBlock { Restart-Service -Name stisvc -Force -ErrorAction Stop; (Get-Service stisvc).Status }
if (Wait-Job $job -Timeout 30) {
    $r = Receive-Job $job -ErrorAction SilentlyContinue
    W "重启图像采集服务 stisvc = $r"
} else {
    W "重启图像采集服务超时（已跳过）"
    Stop-Job $job -ErrorAction SilentlyContinue
}
Remove-Job $job -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# ---- 4) 重要：不要对本机扫描仪做"设备禁用/启用"----
# 实测（Kernel-PnP 事件 225）禁用/启用会触发"未完成的设备拔出"，
# 反而把扫描仪卡死（之后所有扫描都无响应）。因此这里只做进程清理 + 服务重启。
Start-Sleep -Seconds 3

# ---- 5) 验证：NAPS2 能否重新枚举到设备 ----
$listOk = $false
if (Test-Path $NAPS2_CONSOLE) {
    try {
        $out = & $NAPS2_CONSOLE --listdevices --driver twain 2>&1 | Out-String
        $flat = ($out -replace '\s+', ' ').Trim()
        W "设备枚举: $flat"
        if ($out -match 'HUAGOSCAN') { $listOk = $true }
    } catch {
        W "设备枚举异常: $($_.Exception.Message)"
    }
} else {
    W "未找到 NAPS2 控制台程序，跳过设备枚举验证"
}

$final = ''
if ($dev) { $final = (Get-PnpDevice -InstanceId $dev.InstanceId -ErrorAction SilentlyContinue).Status }
W "复位后设备状态 = $final"

if ($listOk -or $final -eq 'OK') {
    W "RESULT=OK"
    W "=== 软复位成功，请重新点击扫描 ==="
    exit 0
} else {
    W "RESULT=FAIL"
    W "=== 软复位未生效，请拔掉 USB 线等 10 秒后插回（建议插主机后置接口）==="
    exit 3
}
