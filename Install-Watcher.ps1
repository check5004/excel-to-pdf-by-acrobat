<#
.SYNOPSIS
    Excel to PDF変換のファイル監視サービスをインストールするスクリプト

.DESCRIPTION
    このスクリプトは、WatchFolder.ps1をWindowsサービスとして登録し、
    unprocessedフォルダの監視を開始します。
    管理者権限で実行する必要があります。

.PARAMETER ConfigPath
    設定ファイル（config.json）のパスを指定します。
    指定されない場合は、スクリプトと同じディレクトリのconfig.jsonを使用します。

.PARAMETER Force
    既存のサービスが存在する場合、強制的に再インストールします。

.EXAMPLE
    .\Install-Watcher.ps1
    .\Install-Watcher.ps1 -ConfigPath "C:\Path\To\config.json"
    .\Install-Watcher.ps1 -Force
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# デフォルト値の設定
if (-not $ConfigPath) {
    $ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
}

# --- 管理者権限の確認 ---
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# --- ログ出力関数 ---
function Write-InstallLog {
    param (
        [string]$Message,
        [string]$LogLevel = "INFO"
    )

    $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$logTimestamp] [$LogLevel] [INSTALL] $Message"

    Write-Host $logMessage

    # ログファイル出力
    try {
        $logDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
        if (-not (Test-Path -Path $logDir -PathType Container)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $logFilePath = Join-Path -Path $logDir -ChildPath "install-$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFilePath -Value $logMessage -Encoding UTF8
    }
    catch {
        Write-Host "ログファイル出力に失敗: $($_.Exception.Message)"
    }
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
        Write-InstallLog "設定ファイルの読み込みに失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
        exit 1
    }
}

# --- サービスの存在確認 ---
function Test-ServiceExists {
    param ([string]$ServiceName)

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        return $null -ne $service
    }
    catch {
        return $false
    }
}

# --- サービスの停止 ---
function Stop-ServiceIfExists {
    param ([string]$ServiceName)

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq "Running") {
            Write-InstallLog "既存のサービスを停止中: $ServiceName"
            Stop-Service -Name $ServiceName -Force
            Start-Sleep -Seconds 3
            Write-InstallLog "サービスを停止しました: $ServiceName"
        }
    }
    catch {
        Write-InstallLog "サービスの停止に失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
    }
}

# --- サービスの削除 ---
function Remove-ServiceIfExists {
    param ([string]$ServiceName)

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($service) {
            Write-InstallLog "既存のサービスを削除中: $ServiceName"
            sc.exe delete $ServiceName | Out-Null
            Start-Sleep -Seconds 2
            Write-InstallLog "サービスを削除しました: $ServiceName"
        }
    }
    catch {
        Write-InstallLog "サービスの削除に失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
    }
}

# --- サービスの登録 ---
function Register-WatcherService {
    param (
        [object]$Config,
        [string]$ScriptPath
    )

    try {
        Write-InstallLog "サービスを登録中: $($Config.ServiceName)"

        # PowerShellスクリプトのパスを絶対パスに変換
        $fullScriptPath = Resolve-Path -Path $ScriptPath -ErrorAction Stop

        # サービス登録用のコマンドライン構築
        $configFilePath = Resolve-Path -Path $ConfigPath -ErrorAction Stop
        # サービス実行の信頼性向上: フルパス/プロファイル無効化/ポリシー回避/ウィンドウ非表示
        $psExe = Join-Path -Path $env:SystemRoot -ChildPath "System32\WindowsPowerShell\v1.0\powershell.exe"
        $arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$fullScriptPath`" -ConfigPath `"$configFilePath`""
        $binPath = "`"$psExe`" $arguments"

        # サービス登録（New-Serviceコマンドレットを使用）
        try {
            $null = New-Service -Name $Config.ServiceName -BinaryPathName $binPath -StartupType Automatic -Description "Excel to PDF変換のファイル監視サービス"
            Write-InstallLog "サービスを登録しました: $($Config.ServiceName)"
        }
        catch {
            # New-Serviceが失敗した場合はsc.exeを使用
            Write-InstallLog "New-Serviceで失敗、sc.exeを使用します: $($_.Exception.Message)" -LogLevel "WARN"
            sc.exe create $Config.ServiceName binPath= `"$binPath`" start= auto obj= LocalSystem | Out-Null
            if ($LASTEXITCODE -ne 0) {
                throw "サービス登録に失敗しました。終了コード: $LASTEXITCODE"
            }
            Write-InstallLog "サービスを登録しました: $($Config.ServiceName)"
        }

        # サービスアカウントを現在のユーザーに設定（Adobe AcrobatのCOMエラーを回避）
        try {
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            sc.exe config $Config.ServiceName obj= $currentUser | Out-Null
            Write-InstallLog "サービスアカウントを現在のユーザーに設定しました: $($Config.ServiceName) ($currentUser)"
        }
        catch {
            Write-InstallLog "サービスアカウントの設定に失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
        }

        # サービスの開始
        Start-Sleep -Seconds 2
        try {
            Start-Service -Name $Config.ServiceName -ErrorAction Stop
            $svc = Get-Service -Name $Config.ServiceName -ErrorAction Stop
            try { $svc.WaitForStatus('Running', '00:00:15') } catch {}
            if ($svc.Status -ne 'Running') {
                throw "サービスの状態が Running になりません (現在: $($svc.Status))"
            }
            Write-InstallLog "サービスを開始しました: $($Config.ServiceName)"
        }
        catch {
            Write-InstallLog "サービスの開始に失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
            return $false
        }

        return $true
    }
    catch {
        Write-InstallLog "サービスの登録に失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
        return $false
    }
}

# --- タスクスケジューラ登録（5分ごとのバッチフォールバック） ---
function Register-WatcherScheduledTask {
    param (
        [object]$Config,
        [string]$ScriptPath
    )

    try {
        Write-InstallLog "タスクスケジューラに登録中: $($Config.ServiceName)"

        $fullScriptPath = Resolve-Path -Path $ScriptPath -ErrorAction Stop
        $configFilePath = Resolve-Path -Path $ConfigPath -ErrorAction Stop

        $psExe = Join-Path -Path $env:SystemRoot -ChildPath "System32\\WindowsPowerShell\\v1.0\\powershell.exe"
        $arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$fullScriptPath`" -ConfigPath `"$configFilePath`" -BatchMode"

        $action = New-ScheduledTaskAction -Execute $psExe -Argument $arguments
        # 修正: 5分ごとに実行するトリガーに変更（フォールバック手段）
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 365)
        # 修正: 実行時間制限を短時間に設定（5分ごとのチェック用）
        $settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -AllowStartIfOnBatteries -StartWhenAvailable
        # 修正: 現在のユーザーで実行（Adobe AcrobatのCOMエラーを回避）
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

        $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal

        $taskName = $Config.ServiceName
        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName

        Write-InstallLog "タスクを登録・開始しました: $taskName"
        Write-InstallLog "実行アカウント: $env:USERNAME"
        Write-InstallLog "実行タイミング: 5分ごと（フォールバック手段）"
        return $true
    }
    catch {
        Write-InstallLog "タスク登録に失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
        return $false
    }
}

# --- タスクスケジューラ登録（ログオン時に常駐・リアルタイム監視） ---
function Register-WatcherRealTimeScheduledTask {
    param (
        [object]$Config,
        [string]$ScriptPath
    )

    try {
        Write-InstallLog "タスクスケジューラ(リアルタイム)に登録中: $($Config.ServiceName)-RealTime"

        $fullScriptPath = Resolve-Path -Path $ScriptPath -ErrorAction Stop
        $configFilePath = Resolve-Path -Path $ConfigPath -ErrorAction Stop

        $psExe = Join-Path -Path $env:SystemRoot -ChildPath "System32\\WindowsPowerShell\\v1.0\\powershell.exe"
        # リアルタイムはBatchModeを付けずに常駐
        $arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$fullScriptPath`" -ConfigPath `"$configFilePath`""

        $action = New-ScheduledTaskAction -Execute $psExe -Argument $arguments
        # ログオン時トリガーで対話セッション内に起動
        $triggerLogon = New-ScheduledTaskTrigger -AtLogOn
        # 即時起動させるため、単発のOnceトリガー（現在時刻+5秒）を追加
        $triggerOnce = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddSeconds(5))
        # 長時間常駐・失敗時リスタート
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 3650) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable
        # 権限に応じてRunLevelを切替（非管理者はLimitedで登録）
        $isAdmin = Test-Administrator
        $runLevel = if ($isAdmin) { 'Highest' } else { 'Limited' }
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel $runLevel

        $task = New-ScheduledTask -Action $action -Trigger $triggerLogon,$triggerOnce -Settings $settings -Principal $principal

        $taskName = "$($Config.ServiceName)-RealTime"
        # 既存があれば置換
        try { Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue } catch {}
        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force -ErrorAction Stop | Out-Null
        # Start-ScheduledTaskはOnceトリガーが付与されているため省略可能だが、明示的にも起動を試行
        try { Start-ScheduledTask -TaskName $taskName -ErrorAction Stop } catch {}

        Write-InstallLog "リアルタイム監視タスクを登録・開始しました: $taskName"
        Write-InstallLog "実行アカウント: $env:USERNAME"
        return $true
    }
    catch {
        Write-InstallLog "リアルタイムタスク登録に失敗しました（Register/Start失敗）: $($_.Exception.Message)" -LogLevel "WARN"
        # フォールバック: schtasksでLIMITED権限のAt logonタスクを登録（非管理者環境向け）
        try {
            $taskName = "$($Config.ServiceName)-RealTime"
            $psExe = Join-Path -Path $env:SystemRoot -ChildPath "System32\\WindowsPowerShell\\v1.0\\powershell.exe"
            $fullScriptPath = Resolve-Path -Path $ScriptPath -ErrorAction Stop
            $configFilePath = Resolve-Path -Path $ConfigPath -ErrorAction Stop
            $arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$fullScriptPath`" -ConfigPath `"$configFilePath`""
            # ONLOGONに加えてONCEトリガー（現在+5秒）を追加し確実に起動
            $at = (Get-Date).AddSeconds(5).ToString('HH:mm')
            $cmd = "schtasks /Create /TN `"$taskName`" /TR `"`"$psExe`" $arguments`" /SC ONLOGON /RL LIMITED /F /RU $env:USERNAME & schtasks /Create /TN `"$taskName`" /TR `"`"$psExe`" $arguments`" /SC ONCE /ST $at /F /RU $env:USERNAME"
            Write-InstallLog "schtasksフォールバックで登録を試行: $taskName"
            cmd.exe /c $cmd | Out-Null
            if ($LASTEXITCODE -ne 0) {
                throw "schtasks登録に失敗（コード: $LASTEXITCODE）"
            }
            # 起動
            cmd.exe /c "schtasks /Run /TN `"$taskName`"" | Out-Null
            Write-InstallLog "リアルタイム監視タスク(schtasks)を登録・開始しました: $taskName"
            return $true
        }
        catch {
            Write-InstallLog "フォールバック(schtasks)でも失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
            return $false
        }
    }
}

# --- パスの検証 ---
function Test-Paths {
    param ([object]$Config)

    $errors = @()

    # 監視パスの確認
    if (-not (Test-Path -Path $Config.WatchPath -PathType Container)) {
        $errors += "監視パスが存在しません: $($Config.WatchPath)"
    }

    # 実行スクリプトの確認
    if (-not (Test-Path -Path $Config.ScriptPath -PathType Leaf)) {
        $errors += "実行スクリプトが存在しません: $($Config.ScriptPath)"
    }

    # ログパスの確認（存在しなければ作成）
    if (-not (Test-Path -Path $Config.LogPath -PathType Container)) {
        try {
            New-Item -Path $Config.LogPath -ItemType Directory -Force | Out-Null
            Write-InstallLog "ログディレクトリを作成しました: $($Config.LogPath)"
        }
        catch {
            $errors += "ログディレクトリの作成に失敗しました: $($Config.LogPath)"
        }
    }

    return $errors
}

# --- メイン処理 ---
try {
    Write-InstallLog "===== ファイル監視サービス インストール開始 ====="

    # 管理者権限の確認（リアルタイムタスクは非管理者でも登録可能なため警告に緩和）
    if (-not (Test-Administrator)) {
        Write-InstallLog "管理者権限ではありません。一部の操作（サービス登録）は失敗する可能性があります。" -LogLevel "WARN"
        Write-InstallLog "リアルタイム監視タスクの登録を試行します。必要に応じて管理者で再実行してください。" -LogLevel "WARN"
    }

    # 設定ファイルの読み込み
    $config = Read-Config -ConfigFilePath $ConfigPath
    Write-InstallLog "設定ファイルを読み込みました: $ConfigPath"

    # パスの検証
    $pathErrors = Test-Paths -Config $config
    if ($pathErrors.Count -gt 0) {
        foreach ($pathError in $pathErrors) {
            Write-InstallLog $pathError -LogLevel "ERROR"
        }
        exit 1
    }

    # 既存サービスの確認
    if (Test-ServiceExists -ServiceName $config.ServiceName) {
        if ($Force) {
            Write-InstallLog "既存のサービスが見つかりました。強制再インストールを実行します。"
            Stop-ServiceIfExists -ServiceName $config.ServiceName
            Remove-ServiceIfExists -ServiceName $config.ServiceName
        }
        else {
            Write-InstallLog "既存のサービスが見つかりました: $($config.ServiceName)" -LogLevel "ERROR"
            Write-InstallLog "再インストールする場合は -Force パラメータを使用してください。" -LogLevel "ERROR"
            exit 1
        }
    }

    # リアルタイム常駐タスクを優先して登録
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "WatchFolder.ps1"
    $success = Register-WatcherRealTimeScheduledTask -Config $config -ScriptPath $scriptPath

    if (-not $success) {
        Write-InstallLog "リアルタイムタスク登録に失敗。5分ごとのバッチ監視へフォールバックします。" -LogLevel "WARN"
        $success = Register-WatcherScheduledTask -Config $config -ScriptPath $scriptPath
    }

    if ($success) {
        Write-InstallLog "===== インストール完了 ====="
        Write-InstallLog "監視パス: $($config.WatchPath)"
        Write-InstallLog "実行スクリプト: $($config.ScriptPath)"
        Write-InstallLog ""
        Write-InstallLog "タスク管理コマンド:"
        Write-InstallLog "  状態確認: Get-ScheduledTask -TaskName `"$($config.ServiceName)-RealTime`""
        Write-InstallLog "  手動起動: Start-ScheduledTask -TaskName `"$($config.ServiceName)-RealTime`""
        Write-InstallLog "  停止: Stop-ScheduledTask -TaskName `"$($config.ServiceName)-RealTime`""
        Write-InstallLog "  アンインストール: .\Uninstall-Watcher.ps1"
    }
    else {
        Write-InstallLog "インストールに失敗しました。" -LogLevel "ERROR"
        exit 1
    }
}
catch {
    Write-InstallLog "インストール中にエラーが発生しました: $($_.Exception.Message)" -LogLevel "ERROR"
    exit 1
}
