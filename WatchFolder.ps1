<#
.SYNOPSIS
    unprocessedフォルダを監視し、Excelファイルの追加を検知してConvert-ExcelToPdf.ps1を自動実行するサービス

.DESCRIPTION
    このスクリプトは、FileSystemWatcherを使用してunprocessedフォルダを監視し、
    Excelファイルが追加された際にConvert-ExcelToPdf.ps1を自動実行します。
    Windowsサービスとして動作することを前提としています。

.PARAMETER ConfigPath
    設定ファイル（config.json）のパスを指定します。
    指定されない場合は、スクリプトと同じディレクトリのconfig.jsonを使用します。

.EXAMPLE
    .\WatchFolder.ps1 -ConfigPath "C:\Path\To\config.json"
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [switch]$BatchMode
)

# デフォルト値の設定
if (-not $ConfigPath) {
    $ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
}

# --- パス解決関数 ---
function Resolve-ConfigPath {
    param (
        [string]$RelativePath,
        [string]$BasePath = $PSScriptRoot
    )

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return $RelativePath
    }
    else {
        return Join-Path -Path $BasePath -ChildPath $RelativePath
    }
}

# --- 設定ファイルの読み込み ---
function Read-Config {
    param ([string]$ConfigFilePath)

    try {
        if (-not (Test-Path -Path $ConfigFilePath -PathType Leaf)) {
            throw "設定ファイルが見つかりません: $ConfigFilePath"
        }

        $configContent = Get-Content -Path $ConfigFilePath -Raw -Encoding UTF8
        $config = $configContent | ConvertFrom-Json

        # 必須項目の確認
        $requiredFields = @("ServiceName", "WatchPath", "ScriptPath", "LogPath", "FileFilters")
        foreach ($field in $requiredFields) {
            if (-not $config.PSObject.Properties.Name -contains $field) {
                throw "設定ファイルに必須項目が不足しています: $field"
            }
        }

        # 相対パスを絶対パスに変換
        $config.WatchPath = Resolve-ConfigPath -RelativePath $config.WatchPath
        $config.ScriptPath = Resolve-ConfigPath -RelativePath $config.ScriptPath
        $config.LogPath = Resolve-ConfigPath -RelativePath $config.LogPath

        return $config
    }
    catch {
        Write-Error "設定ファイルの読み込みに失敗しました: $($_.Exception.Message)"
        exit 1
    }
}

# --- ログ出力関数 ---
function Write-ServiceLog {
    param (
        [string]$Message,
        [string]$LogLevel = "INFO"
    )

    $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$logTimestamp] [$LogLevel] [SERVICE] $Message"

    # コンソール出力
    Write-Host $logMessage

    # ログファイル出力
    try {
        $logFilePath = Join-Path -Path $script:config.LogPath -ChildPath "watcher-$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFilePath -Value $logMessage -Encoding UTF8
    }
    catch {
        # ログファイル出力に失敗しても処理は継続
        Write-Host "ログファイル出力に失敗: $($_.Exception.Message)"
    }
}

# --- ファイル処理関数 ---
function Invoke-NewFileProcessing {
    param (
        [string]$FilePath,
        [string]$FileName,
        [string]$OverrideScriptPath,
        [string]$OverrideWatchPath
    )

    try {
        Write-ServiceLog "新しいファイルを検知: $FileName"

        # ファイルが完全に書き込まれるまで待機
        $maxWaitTime = 30 # 30秒
        $waitTime = 0
        while ($waitTime -lt $maxWaitTime) {
            try {
                $fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                $fileStream.Close()
                break
            }
            catch {
                Start-Sleep -Seconds 1
                $waitTime++
            }
        }

        if ($waitTime -ge $maxWaitTime) {
            Write-ServiceLog "ファイルの書き込み完了を待機中にタイムアウトしました: $FileName" -LogLevel "WARN"
        }

        # Convert-ExcelToPdf.ps1を実行
        $scriptPathToRun = if ($OverrideScriptPath) { $OverrideScriptPath } else { $script:config.ScriptPath }
        $watchPathForBase = if ($OverrideWatchPath) { $OverrideWatchPath } else { $script:config.WatchPath }
        Write-ServiceLog "変換スクリプトを実行開始: $scriptPathToRun"

        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        # サービス環境でも安定動作するようにPS実行を明示・プロファイル無効化・ポリシー回避
        $psExe = Join-Path -Path $env:SystemRoot -ChildPath "System32/WindowsPowerShell/v1.0/powershell.exe"
        $processInfo.FileName = $psExe
        # Convert スクリプトの BasePath は WatchPath の親ディレクトリに固定
        $basePath = Split-Path -Path $watchPathForBase -Parent
        $processInfo.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPathToRun`" -BasePath `"$basePath`""
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($processInfo)
        $process.WaitForExit()

        if ($process.ExitCode -eq 0) {
            Write-ServiceLog "変換スクリプトの実行が完了しました: $FileName"
        }
        else {
            Write-ServiceLog "変換スクリプトの実行でエラーが発生しました (終了コード: $($process.ExitCode)): $FileName" -LogLevel "ERROR"
        }
    }
    catch {
        Write-ServiceLog "ファイル処理中にエラーが発生しました: $($_.Exception.Message)" -LogLevel "ERROR"
    }
}

# --- メイン処理 ---
try {
    # 設定ファイルの読み込み
    $script:config = Read-Config -ConfigFilePath $ConfigPath

    Write-ServiceLog "===== ファイル監視サービス開始 ====="
    Write-ServiceLog "監視パス: $($script:config.WatchPath)"
    Write-ServiceLog "実行スクリプト: $($script:config.ScriptPath)"
    Write-ServiceLog "ログパス: $($script:config.LogPath)"

    # 監視パスの存在確認
    if (-not (Test-Path -Path $script:config.WatchPath -PathType Container)) {
        throw "監視パスが存在しません: $($script:config.WatchPath)"
    }

    # 実行スクリプトの存在確認
    if (-not (Test-Path -Path $script:config.ScriptPath -PathType Leaf)) {
        throw "実行スクリプトが存在しません: $($script:config.ScriptPath)"
    }

    # ログディレクトリの作成
    if (-not (Test-Path -Path $script:config.LogPath -PathType Container)) {
        New-Item -Path $script:config.LogPath -ItemType Directory -Force | Out-Null
        Write-ServiceLog "ログディレクトリを作成しました: $($script:config.LogPath)"
    }

    # FileSystemWatcherの設定
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $script:config.WatchPath
    $watcher.Filter = "*.*"
    $watcher.IncludeSubdirectories = $false
    $watcher.InternalBufferSize = 64KB
    $watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::Size
    # 先にイベントを登録してから有効化（起動直後の取りこぼし防止）

    # イベントアクション（Created/Changed/Renamed を処理）
    $action = {
        $path = $Event.SourceEventArgs.FullPath
        $name = $Event.SourceEventArgs.Name
        $eventName = $Event.SourceEventArgs.ChangeType
        $filters = $Event.MessageData.Filters
        $scriptPath = $Event.MessageData.ScriptPath
        $watchPath = $Event.MessageData.WatchPath

        # デバッグ: 受信イベントを必ず記録
        Write-ServiceLog "イベント受信: $eventName - $name ($path)"

        # Excelの一時ロックファイル(~$)は無視
        if ($name -like "~$*") {
            Write-ServiceLog "ロックファイルをスキップ: $name" -LogLevel "DEBUG"
            return
        }

        # ワイルドカードでのパターン一致（*.xlsx, *.xls など）
        $matched = $false
        foreach ($pat in $filters) {
            if ($name -like $pat) { $matched = $true; break }
        }

        if ($matched) {
            Invoke-NewFileProcessing -FilePath $path -FileName $name -OverrideScriptPath $scriptPath -OverrideWatchPath $watchPath
        }
        else {
            Write-ServiceLog "対象外パターン: $name (filters: $($filters -join ', '))" -LogLevel "DEBUG"
        }
    }

    $msg = @{ Filters = $script:config.FileFilters; ScriptPath = $script:config.ScriptPath; WatchPath = $script:config.WatchPath }
    Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action -MessageData $msg | Out-Null
    Register-ObjectEvent -InputObject $watcher -EventName "Changed" -Action $action -MessageData $msg | Out-Null
    Register-ObjectEvent -InputObject $watcher -EventName "Renamed" -Action $action -MessageData $msg | Out-Null

    # バッファオーバーフロー等のエラー監視
    Register-ObjectEvent -InputObject $watcher -EventName "Error" -Action {
        try {
            $ex = $Event.SourceEventArgs.GetException()
            Write-ServiceLog "監視エラー: $($ex.Message)" -LogLevel "ERROR"
        }
        catch {
            Write-ServiceLog "監視エラー: 詳細不明" -LogLevel "ERROR"
        }
    } | Out-Null

    $watcher.EnableRaisingEvents = $true

    Write-ServiceLog "ファイル監視を開始しました。監視対象拡張子: $($script:config.FileFilters -join ', ')"

    # 起動時の既存ファイルを初期スキャンして取りこぼしを防止
    try {
        $patterns = $script:config.FileFilters
        $existing = @()
        foreach ($pat in $patterns) {
            $existing += Get-ChildItem -Path $script:config.WatchPath -Filter $pat -File -ErrorAction SilentlyContinue
        }
        foreach ($f in $existing | Sort-Object FullName -Unique) {
            Invoke-NewFileProcessing -FilePath $f.FullName -FileName $f.Name
        }
    }
    catch {
        Write-ServiceLog "初期スキャンでエラー: $($_.Exception.Message)" -LogLevel "WARN"
    }

    # バッチモードの場合は一度だけ実行して終了
    if ($BatchMode) {
        Write-ServiceLog "バッチモードで実行完了。5分後に再実行されます。"
        return
    }

    # サービスとして動作させるため、無限ループで待機
    while ($true) {
        # イベントキューをポンプしつつ待機
        Wait-Event -Timeout 10 | Out-Null
    }
}
catch {
    Write-ServiceLog "サービス開始時にエラーが発生しました: $($_.Exception.Message)" -LogLevel "ERROR"
    exit 1
}
finally {
    # クリーンアップ
    if ($watcher) {
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
    }

    # イベントの登録解除
    Get-EventSubscriber | Where-Object { $_.SourceObject -eq $watcher } | Unregister-Event

    Write-ServiceLog "===== ファイル監視サービス終了 ====="
}
