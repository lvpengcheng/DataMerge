[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$VMName = "ubuntu",
    [int]$MemoryGB = 24,
    [int]$ProcessorCount = 4,
    [int]$ProcessorMaximumPercent = 50,
    [long]$MaximumIOPS = 5000,
    [switch]$MergeCheckpoint
)

$ErrorActionPreference = "Stop"

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "请右键 PowerShell，选择‘以管理员身份运行’，再执行此脚本。"
}

$vm = Get-VM -Name $VMName -ErrorAction Stop
$hostRamGB = [math]::Floor((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
if ($MemoryGB -gt ($hostRamGB - 24)) {
    throw "拒绝设置：必须至少给 Windows 主机保留 24GB；本机共 ${hostRamGB}GB。"
}

$backupDir = Join-Path $env:ProgramData "DataMerge-HyperV-Guard"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$stateFile = Join-Path $backupDir "$VMName-before-$stamp.clixml"
[pscustomobject]@{
    VM = Get-VM -Name $VMName
    Memory = Get-VMMemory -VMName $VMName
    Processor = Get-VMProcessor -VMName $VMName
    Disks = @(Get-VMHardDiskDrive -VMName $VMName)
    Checkpoints = @(Get-VMSnapshot -VMName $VMName -ErrorAction SilentlyContinue)
} | Export-Clixml -LiteralPath $stateFile
Write-Host "原配置已备份：$stateFile" -ForegroundColor Cyan

# 避免 Windows 关机时保存数十 GB 的 VM 内存；磁盘卡顿时 Save 最容易拖死宿主机。
if ($PSCmdlet.ShouldProcess($VMName, "设置宿主机保护策略")) {
    Set-VM -Name $VMName `
        -AutomaticStopAction TurnOff `
        -AutomaticCheckpointsEnabled $false `
        -CheckpointType ProductionOnly
}

if ($vm.State -ne "Off") {
    Write-Warning "VM 当前为 $($vm.State)。CPU/固定内存设置需要关机后执行。"
    Write-Host "正常时： Stop-VM -Name '$VMName'"
    Write-Host "完全失联时： Stop-VM -Name '$VMName' -TurnOff -Force"
    Write-Host "关机后重新运行本脚本。已先应用无需关机的保护项。"
    exit 2
}

if ($PSCmdlet.ShouldProcess($VMName, "固定内存 ${MemoryGB}GB，CPU ${ProcessorCount} 核/上限 ${ProcessorMaximumPercent}%")) {
    Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $false -StartupBytes ($MemoryGB * 1GB)
    Set-VMProcessor -VMName $VMName `
        -Count $ProcessorCount `
        -Reserve 0 `
        -Maximum $ProcessorMaximumPercent `
        -RelativeWeight 20

    Get-VMHardDiskDrive -VMName $VMName | ForEach-Object {
        Set-VMHardDiskDrive -VMHardDiskDrive $_ -MaximumIOPS $MaximumIOPS
    }
}

$snapshots = @(Get-VMSnapshot -VMName $VMName -ErrorAction SilentlyContinue)
if ($snapshots.Count -gt 0) {
    Write-Warning "发现 $($snapshots.Count) 个检查点。当前 VM 正运行 .avhdx，建议维护窗口内合并。"
    $snapshots | Format-Table Name, SnapshotType, CreationTime -AutoSize
    if ($MergeCheckpoint -and $PSCmdlet.ShouldProcess($VMName, "删除检查点并合并 AVHDX（可能耗时很久）")) {
        $snapshots | Remove-VMSnapshot -Confirm:$false
        Write-Host "已提交检查点合并。请用 Get-VHD / Get-VMHardDiskDrive 确认完成后再启动 VM。" -ForegroundColor Yellow
    } else {
        Write-Host "确认已有备份且磁盘空间充足后运行："
        Write-Host ".\tools\configure_hyperv_guard.ps1 -MergeCheckpoint"
    }
}

Write-Host "保护配置完成。" -ForegroundColor Green
Get-VM -Name $VMName | Format-List Name, State, AutomaticStopAction, AutomaticCheckpointsEnabled, CheckpointType
Get-VMMemory -VMName $VMName | Format-List DynamicMemoryEnabled, Startup, Minimum, Maximum
Get-VMProcessor -VMName $VMName | Format-List Count, Reserve, Maximum, RelativeWeight
