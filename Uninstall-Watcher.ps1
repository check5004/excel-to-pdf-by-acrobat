<#
.SYNOPSIS
    Excel to PDF変換のファイル監視サービスをアンインストールするスクリプト

.DESCRIPTION
    このスクリプトは、登録されたファイル監視サービスを停止・削除し、
    関連するファイルをクリーンアップします。
    管理者権限で実行する必要があります。

.PARAMETER ConfigPath
    設定ファイル（config.json）のパスを指定します。
    指定されない場合は、スクリプトと同じディレクトリのconfig.jsonを使用します。

.PARAMETER KeepLogs
    ログファイルを保持する場合は指定します。

.EXAMPLE
    .\Uninstall-Watcher.ps1
    .\Uninstall-Watcher.ps1 -ConfigPath "C:\Path\To\config.json"
    .\Uninstall-Watcher.ps1 -KeepLogs
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [switch]$KeepLogs
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
function Write-UninstallLog {
    param (
        [string]$Message,
        [string]$LogLevel = "INFO"
    )

    $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$logTimestamp] [$LogLevel] [UNINSTALL] $Message"

    Write-Host $logMessage

    # ログファイル出力
    try {
        $logDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
        if (-not (Test-Path -Path $logDir -PathType Container)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $logFilePath = Join-Path -Path $logDir -ChildPath "uninstall-$(Get-Date -Format 'yyyy-MM-dd').log"
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
            Write-UninstallLog "設定ファイルが見つかりません: $ConfigFilePath" -LogLevel "WARN"
            return $null
        }

        $configContent = Get-Content -Path $ConfigFilePath -Raw -Encoding UTF8
        $config = $configContent | ConvertFrom-Json

        # 必須項目の確認
        $requiredFields = @("ServiceName", "WatchPath", "ScriptPath", "LogPath", "FileFilters")
        foreach ($field in $requiredFields) {
            if (-not $config.PSObject.Properties.Name -contains $field) {
                Write-UninstallLog "設定ファイルに必須項目が不足しています: $field" -LogLevel "WARN"
                return $null
            }
        }

        # 相対パスを絶対パスに変換
        $config.WatchPath = Resolve-ConfigPath -RelativePath $config.WatchPath
        $config.ScriptPath = Resolve-ConfigPath -RelativePath $config.ScriptPath
        $config.LogPath = Resolve-ConfigPath -RelativePath $config.LogPath

        return $config
    }
    catch {
        Write-UninstallLog "設定ファイルの読み込みに失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
        return $null
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
        if ($service) {
            if ($service.Status -eq "Running") {
                Write-UninstallLog "サービスを停止中: $ServiceName"
                Stop-Service -Name $ServiceName -Force

                # 停止完了まで待機
                $timeout = 30
                $elapsed = 0
                while ($service.Status -ne "Stopped" -and $elapsed -lt $timeout) {
                    Start-Sleep -Seconds 1
                    $elapsed++
                    $service.Refresh()
                }

                if ($service.Status -eq "Stopped") {
                    Write-UninstallLog "サービスを停止しました: $ServiceName"
                }
                else {
                    Write-UninstallLog "サービスの停止にタイムアウトしました: $ServiceName" -LogLevel "WARN"
                }
            }
            else {
                Write-UninstallLog "サービスは既に停止しています: $ServiceName"
            }
        }
        else {
            Write-UninstallLog "サービスが見つかりません: $ServiceName"
        }
    }
    catch {
        Write-UninstallLog "サービスの停止に失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
    }
}

# --- サービスの削除 ---
function Remove-ServiceIfExists {
    param ([string]$ServiceName)

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($service) {
            Write-UninstallLog "サービスを削除中: $ServiceName"

            # sc.exeを使用してサービスを削除
            sc.exe delete $ServiceName | Out-Null
            Start-Sleep -Seconds 2

            # 削除確認
            $serviceAfter = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            if (-not $serviceAfter) {
                Write-UninstallLog "サービスを削除しました: $ServiceName"
            }
            else {
                Write-UninstallLog "サービスの削除に失敗しました: $ServiceName" -LogLevel "ERROR"
            }
        }
        else {
            Write-UninstallLog "削除対象のサービスが見つかりません: $ServiceName"
        }
    }
    catch {
        Write-UninstallLog "サービスの削除に失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
    }
}

# --- ログファイルのクリーンアップ ---
function Remove-LogFiles {
    param (
        [string]$LogPath,
        [bool]$KeepLogs
    )

    if ($KeepLogs) {
        Write-UninstallLog "ログファイルを保持します: $LogPath"
        return
    }

    try {
        if (Test-Path -Path $LogPath -PathType Container) {
            $logFiles = Get-ChildItem -Path $LogPath -Filter "watcher-*.log" -ErrorAction SilentlyContinue
            if ($logFiles) {
                foreach ($logFile in $logFiles) {
                    Remove-Item -Path $logFile.FullName -Force
                    Write-UninstallLog "ログファイルを削除しました: $($logFile.Name)"
                }
            }
            else {
                Write-UninstallLog "削除対象のログファイルが見つかりませんでした"
            }
        }
        else {
            Write-UninstallLog "ログディレクトリが存在しません: $LogPath"
        }
    }
    catch {
        Write-UninstallLog "ログファイルの削除に失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
    }
}

# --- 設定ファイルのバックアップ ---
function Backup-ConfigFile {
    param ([string]$ConfigPath)

    try {
        if (Test-Path -Path $ConfigPath -PathType Leaf) {
            $backupPath = $ConfigPath + ".backup." + (Get-Date -Format "yyyyMMdd-HHmmss")
            Copy-Item -Path $ConfigPath -Destination $backupPath
            Write-UninstallLog "設定ファイルをバックアップしました: $backupPath"
        }
    }
    catch {
        Write-UninstallLog "設定ファイルのバックアップに失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
    }
}

# --- メイン処理 ---
try {
    Write-UninstallLog "===== ファイル監視サービス アンインストール開始 ====="

    # 管理者権限の確認
    if (-not (Test-Administrator)) {
        Write-UninstallLog "このスクリプトは管理者権限で実行する必要があります。" -LogLevel "ERROR"
        Write-UninstallLog "PowerShellを管理者として実行し直してください。" -LogLevel "ERROR"
        exit 1
    }

    # 設定ファイルの読み込み
    $config = Read-Config -ConfigFilePath $ConfigPath

    # サービス名の決定
    $serviceName = if ($config) { $config.ServiceName } else { "ExcelToPdfWatcher" }

    Write-UninstallLog "対象サービス: $serviceName"

    # サービスの存在確認
    if (-not (Test-ServiceExists -ServiceName $serviceName)) {
        Write-UninstallLog "対象のサービスが見つかりません: $serviceName"
        Write-UninstallLog "アンインストール対象がありません。"
        exit 0
    }

    # 設定ファイルのバックアップ
    if ($config) {
        Backup-ConfigFile -ConfigPath $ConfigPath
    }

    # サービスの停止
    Stop-ServiceIfExists -ServiceName $serviceName

    # サービスの削除
    Remove-ServiceIfExists -ServiceName $serviceName

    # ログファイルのクリーンアップ
    if ($config) {
        Remove-LogFiles -LogPath $config.LogPath -KeepLogs $KeepLogs
    }

    Write-UninstallLog "===== アンインストール完了 ====="
    Write-UninstallLog "ファイル監視サービスが正常にアンインストールされました。"

    if (-not $KeepLogs) {
        Write-UninstallLog "ログファイルも削除されました。"
    }
    else {
        Write-UninstallLog "ログファイルは保持されています。"
    }
}
catch {
    Write-UninstallLog "アンインストール中にエラーが発生しました: $($_.Exception.Message)" -LogLevel "ERROR"
    exit 1
}
