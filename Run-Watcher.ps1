param(
    [Parameter(Mandatory=$false)]
    [string]$TaskName = "ExcelToPdfWatcher-RealTime"
)

Write-Host "[RUN] タスクを再起動します: $TaskName"
try {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
} catch {}

try {
    Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    Write-Host "[RUN] タスクを開始しました: $TaskName"
} catch {
    Write-Host "[RUN][ERROR] タスク開始に失敗: $($_.Exception.Message)"
    exit 1
}
