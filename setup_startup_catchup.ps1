param(
    [string]$TaskName = "CtripOrderNoonCatchupOnLogon"
)

$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $ProjectDir ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $Python)) {
    throw "找不到项目虚拟环境 Python：$Python"
}

$LogsDir = Join-Path $ProjectDir "logs"
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

$Arguments = @(
    "-X",
    "utf8",
    "main.py",
    "--recent-days",
    "7",
    "--update-template",
    "--daily-backup",
    "--yes",
    "--catch-up-missed",
    "--wait-for-network",
    "--log-file",
    "logs\startup_catchup.log"
) -join " "

$Action = New-ScheduledTaskAction -Execute $Python -Argument $Arguments -WorkingDirectory $ProjectDir
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "开机登录后检查携程中午或晚间订单更新是否漏跑，符合时间窗口时自动补跑。" `
    -Force | Out-Null

Write-Host "已创建或更新 Windows 登录后补跑任务：$TaskName"
Write-Host "项目目录：$ProjectDir"
Write-Host "日志文件：$(Join-Path $LogsDir 'startup_catchup.log')"
