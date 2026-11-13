# ============================================================
#  setup_scheduler.ps1  —  Weekly Operations Report
#  Registers a Windows Task Scheduler task that runs the report
#  every Monday at 06:00.
#
#  HOW TO RUN (once, as Administrator):
#    Right-click PowerShell -> "Run as administrator"
#    cd to the weekly_report\ folder, then:
#    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#    .\setup_scheduler.ps1
#
#  To remove the task:
#    Unregister-ScheduledTask -TaskName "WeeklyOpsReport" -Confirm:$false
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TaskName  = "WeeklyOpsReport"
$BatFile   = Join-Path $ScriptDir "run_weekly_ops_report.bat"

if (-not (Test-Path $BatFile)) {
    Write-Error "Cannot find $BatFile"
    return
}

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task '$TaskName'."
}

# Fire every Monday at 06:00
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "06:00"

$Action    = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$BatFile`""
# Interactive: runs as you while you're signed in, so Excel can refresh the
# Power BI workbook with your sign-in and SQL can use Windows authentication.
# (S4U runs without your network credentials, so both of those would fail.)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest
$Settings  = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName  $TaskName `
    -Action    $Action `
    -Trigger   $Trigger `
    -Principal $Principal `
    -Settings  $Settings `
    -Force | Out-Null

Write-Host "Task '$TaskName' registered — runs every Monday at 06:00." -ForegroundColor Green
Write-Host "  Script : $BatFile"
Write-Host "  Test   : Start-ScheduledTask -TaskName '$TaskName'"
Write-Host ""
Write-Host "Before the first run: fill in the goals and SQL servers in config.env"
Write-Host "(see README.md)."
