﻿# 建立 Windows 工作排程（重新執行會覆蓋舊設定）
#   1.「台股盤後抓取」：每個平日 16:30 與 21:30 執行 run_after_close.bat
#        16:30 抓盤後主要資料（資料不齊會重試到 18:30）；21:30 補上晚上才公布的台股融資
#   2.「台股儀表板伺服器」：登入 Windows 時在背景啟動 server.py（網頁新增／刪除自選股用）
# 使用方式：在這個檔案上按右鍵 →「用 PowerShell 執行」
# 移除：Unregister-ScheduledTask -TaskName "台股盤後抓取" -Confirm:$false
#       Unregister-ScheduledTask -TaskName "台股儀表板伺服器" -Confirm:$false
$root = $PSScriptRoot
$pyw = Join-Path $env:LOCALAPPDATA "Programs\Python\Python313\pythonw.exe"

$bat = Join-Path $root "run_after_close.bat"
$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$bat`"" -WorkingDirectory $root
$days = "Monday","Tuesday","Wednesday","Thursday","Friday"
$triggers = @(
  (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At "16:30"),
  (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At "21:30")
)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 3) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName "台股盤後抓取" -Action $action -Trigger $triggers -Settings $settings -Description "每日盤後抓取台股資料並產生儀表板" -Force | Out-Null

$srvAction = New-ScheduledTaskAction -Execute $pyw -Argument "`"$(Join-Path $root 'server.py')`"" -WorkingDirectory $root
$srvTrigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$srvSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "台股儀表板伺服器" -Action $srvAction -Trigger $srvTrigger -Settings $srvSettings -Description "本機儀表板伺服器（http://127.0.0.1:8765）" -Force | Out-Null
Start-ScheduledTask -TaskName "台股儀表板伺服器"

Write-Host "已建立排程：平日 16:30、21:30 抓資料；登入時自動啟動儀表板伺服器（已立即啟動）"
Read-Host "按 Enter 關閉"
