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
        $arguments = "-ExecutionPolicy RemoteSigned -File `"$fullScriptPath`" -ConfigPath `"$ConfigPath`""

        # サービス登録
        sc.exe create $Config.ServiceName binPath= "powershell.exe $arguments" start= auto | Out-Null

        if ($LASTEXITCODE -eq 0) {
            Write-InstallLog "サービスを登録しました: $($Config.ServiceName)"

            # サービスの開始
            Start-Sleep -Seconds 2
            Start-Service -Name $Config.ServiceName
            Write-InstallLog "サービスを開始しました: $($Config.ServiceName)"

            return $true
        }
        else {
            throw "サービス登録に失敗しました。終了コード: $LASTEXITCODE"
        }
    }
    catch {
        Write-InstallLog "サービスの登録に失敗しました: $($_.Exception.Message)" -LogLevel "ERROR"
        return $false
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

    # 管理者権限の確認
    if (-not (Test-Administrator)) {
        Write-InstallLog "このスクリプトは管理者権限で実行する必要があります。" -LogLevel "ERROR"
        Write-InstallLog "PowerShellを管理者として実行し直してください。" -LogLevel "ERROR"
        exit 1
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

    # サービスの登録
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "WatchFolder.ps1"
    $success = Register-WatcherService -Config $config -ScriptPath $scriptPath

    if ($success) {
        Write-InstallLog "===== インストール完了 ====="
        Write-InstallLog "サービス名: $($config.ServiceName)"
        Write-InstallLog "監視パス: $($config.WatchPath)"
        Write-InstallLog "実行スクリプト: $($config.ScriptPath)"
        Write-InstallLog ""
        Write-InstallLog "サービス管理コマンド:"
        Write-InstallLog "  開始: Start-Service -Name `"$($config.ServiceName)`""
        Write-InstallLog "  停止: Stop-Service -Name `"$($config.ServiceName)`""
        Write-InstallLog "  状態確認: Get-Service -Name `"$($config.ServiceName)`""
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
