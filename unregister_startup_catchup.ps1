param(
    [string]$TaskName = "CtripOrderNoonCatchupOnLogon"
)

$ErrorActionPreference = "Stop"

$Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $Task) {
    Write-Host "未找到 Windows 登录后补跑任务：$TaskName"
    exit 0
}

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Write-Host "已取消 Windows 登录后补跑任务：$TaskName"
