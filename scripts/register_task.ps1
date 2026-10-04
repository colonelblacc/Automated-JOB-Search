$Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument '/c "c:\JOB SEARCH\AI TOOL\scripts\run_daily_radar.bat"'
$Trigger = New-ScheduledTaskTrigger -Daily -At 08:30AM
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName "DailyEmbeddedJobRadar" -Action $Action -Trigger $Trigger -Settings $Settings -Force
Write-Host "SUCCESS: Windows Scheduled Task 'DailyEmbeddedJobRadar' registered for 08:30 AM daily with StartWhenAvailable."
