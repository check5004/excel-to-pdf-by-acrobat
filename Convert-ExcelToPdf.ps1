<#
.SYNOPSIS
    指定されたディレクトリ内のExcelファイルをPDFに変換するPowerShellスクリプト。

.DESCRIPTION
    このスクリプトは、Adobe AcrobatのCOMコンポーネントを利用して、特定のフォルダ構造内でExcelからPDFへの変換を自動化します。
    - 'unprocessed'フォルダ内のExcelファイルを検索します。
    - ファイルを'processing'フォルダに移動してから変換処理を開始します。
    - 成功した場合、PDFは'completed/pdf'に、元のExcelは'completed/excel'に移動します。
    - 失敗した場合、元のExcelは'failed'フォルダに移動します。
    - 全ての処理は日付ベースのログファイルに記録されます。

.PARAMETER BasePath
    処理の起点となるルートディレクトリのパスを指定します。
    このディレクトリ内に'unprocessed', 'processing'などのサブディレクトリが作成されます。
    指定されない場合は、スクリプト実行時のカレントディレクトリが使用されます。

.EXAMPLE
    .\Convert-ExcelToPdf.ps1 -BasePath "C:\ExcelConversion"

    'C:\ExcelConversion'をベースとして処理を実行します。
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$BasePath = $PWD
)

# --- 設定ファイルの読み込み（UseAcrobatフラグ対応） ---
function Read-ProjectConfig {
    try {
        $configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
        if (-not (Test-Path -Path $configPath -PathType Leaf)) { return @{ UseAcrobat = $true } }
        $json = Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $useAcrobat = $true
        if ($json.PSObject.Properties.Name -contains 'UseAcrobat') { $useAcrobat = [bool]$json.UseAcrobat }
        return @{ UseAcrobat = $useAcrobat; ConfigPath = $configPath; Raw = $json }
    }
    catch {
        return @{ UseAcrobat = $true }
    }
}

function Update-ProjectConfigUseAcrobatFalse {
    param(
        [string]$ConfigPath,
        [object]$Raw
    )
    try {
        if (-not $ConfigPath) { return }
        if (-not $Raw) {
            $Raw = @{}
        }
        $Raw.UseAcrobat = $false
        $Raw.LastUpdated = (Get-Date -Format 'yyyy-MM-dd')
        ($Raw | ConvertTo-Json -Depth 5) | Set-Content -Path $ConfigPath -Encoding UTF8
        Write-Log -Message "config.jsonのUseAcrobatをfalseに更新しました（自動フォールバック）"
    }
    catch {
        Write-Log -Message "config.jsonの更新に失敗しました: $($_.Exception.Message)" -LogLevel "WARN"
    }
}

# 設定のロード
$projectConfig = Read-ProjectConfig
$useAcrobatMode = $projectConfig.UseAcrobat

# --- 関数定義: ログ出力 ---
function Write-Log {
    param (
        [string]$Message,
        [string]$LogLevel = "INFO"
    )
    $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$logTimestamp] [$LogLevel] $Message"

    # ログファイルパスをグローバルスコープから参照
    Add-Content -Path $script:logFilePath -Value $logMessage
}

# --- 関数定義: ディレクトリ説明ファイル作成 ---
function New-DirectoryDescriptionFiles {
    param (
        [hashtable]$DirectoryDescriptions
    )

    foreach ($dirPath in $DirectoryDescriptions.Keys) {
        $description = $DirectoryDescriptions[$dirPath]
        $descriptionFilePath = Join-Path -Path $dirPath -ChildPath $description

        try {
            # ファイルが既に存在する場合はスキップ
            if (-not (Test-Path -Path $descriptionFilePath -PathType Leaf)) {
                # 空のファイルを作成（ファイル名が説明文）
                New-Item -Path $descriptionFilePath -ItemType File -Force | Out-Null
                Write-Log -Message "説明ファイルを作成しました: $descriptionFilePath"
            }
        }
        catch {
            Write-Log -Message "説明ファイルの作成に失敗しました: $descriptionFilePath - エラー: $($_.Exception.Message)" -LogLevel "WARN"
        }
    }
}

# --- 1. 初期設定とディレクトリ構造の確認 ---
try {
    # 処理ディレクトリのパスを定義
    $unprocessedDirPath = Join-Path -Path $BasePath -ChildPath "unprocessed"
    $processingDirPath  = Join-Path -Path $BasePath -ChildPath "processing"
    $completedDirPath   = Join-Path -Path $BasePath -ChildPath "completed"
    $completedPdfPath   = Join-Path -Path $completedDirPath -ChildPath "pdf"
    $completedExcelPath = Join-Path -Path $completedDirPath -ChildPath "excel"
    $failedDirPath      = Join-Path -Path $BasePath -ChildPath "failed"
    $logsDirPath        = Join-Path -Path $BasePath -ChildPath "logs"

    # ログファイルパスをスクリプト全体で使えるように定義
    $script:logFilePath = Join-Path -Path $logsDirPath -ChildPath "$(Get-Date -Format 'yyyy-MM-dd').log"

    # 各ディレクトリが存在しなければ作成
    $directories = @(
        $BasePath, $unprocessedDirPath, $processingDirPath, $completedDirPath,
        $completedPdfPath, $completedExcelPath, $failedDirPath, $logsDirPath
    )

    foreach ($dir in $directories) {
        if (-not (Test-Path -Path $dir -PathType Container)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
    }

    # --- ディレクトリ説明ファイルの作成 ---
    $directoryDescriptions = @{
        $unprocessedDirPath = "未処理のExcelファイルを格納"
        $processingDirPath = "処理中のファイルを格納"
        $completedDirPath = "処理完了したファイルを格納"
        $completedPdfPath = "変換されたPDFファイルを格納"
        $completedExcelPath = "元のExcelファイルを格納"
        $failedDirPath = "処理に失敗したファイルを格納"
        $logsDirPath = "ログファイルを格納"
    }

    New-DirectoryDescriptionFiles -DirectoryDescriptions $directoryDescriptions
}
catch {
    # ディレクトリ作成に失敗した場合、処理を中断
    $errorMessage = "初期化に失敗しました。ディレクトリの作成権限を確認してください。エラー: $($_.Exception.Message)"
    # ログファイルが確定できないため、コンソールにエラー出力
    Write-Error $errorMessage
    exit 1
}

# --- スクリプト開始ログ ---
Write-Log -Message "===== スクリプト実行開始 ====="

# --- 2. 未処理ファイルの検索 ---
$excelFiles = Get-ChildItem -Path $unprocessedDirPath -Include "*.xlsx", "*.xls" -Recurse -Depth 0
if ($null -eq $excelFiles) {
    Write-Log -Message "処理対象のExcelファイルが見つかりませんでした。"
    Write-Log -Message "===== スクリプト実行終了 ====="
    exit 0
}

Write-Log -Message "$($excelFiles.Count)件のファイルが見つかりました。処理を開始します。"

# --- 3. ファイルごとの変換処理ループ ---
function Get-SafeFileName {
    param([string]$Name)
    # 禁止文字を置換、先頭末尾のピリオド/スペースをトリム、長すぎる場合はカット
    $sanitized = ($Name -replace '[\\/:*?"<>|]', "_")
    $sanitized = $sanitized.Trim().Trim('.')
    if ([string]::IsNullOrWhiteSpace($sanitized)) { $sanitized = "Sheet" }
    if ($sanitized.Length -gt 120) { $sanitized = $sanitized.Substring(0,120) }
    return $sanitized
}

foreach ($file in $excelFiles) {
    $processingPath = Join-Path -Path $processingDirPath -ChildPath $file.Name

    # 出力ディレクトリ: completed/pdf/<ExcelBaseName>/
    $excelBaseName = $file.BaseName
    $excelOutputDir = Join-Path -Path $completedPdfPath -ChildPath (Get-SafeFileName -Name $excelBaseName)
    if (-not (Test-Path -Path $excelOutputDir -PathType Container)) {
        New-Item -Path $excelOutputDir -ItemType Directory -Force | Out-Null
    }

    # 一時ディレクトリ: processing/<ExcelBaseName>/_temp_pdf
    $tempPdfDir = Join-Path -Path $processingDirPath -ChildPath ((Get-SafeFileName -Name $excelBaseName) + "_temp_pdf")
    if (-not (Test-Path -Path $tempPdfDir -PathType Container)) {
        New-Item -Path $tempPdfDir -ItemType Directory -Force | Out-Null
    }

    # COMオブジェクト変数を初期化
    $excelApp = $null
    $workbook = $null
    $acrobatApp = $null
    $avDoc = $null
    $pdDoc = $null

    try {
        Write-Log -Message "処理開始: $($file.Name)"

        # ファイルを作業中ディレクトリへ移動
        Move-Item -Path $file.FullName -Destination $processingPath -Force

        # Excel COMの起動
        try {
            $excelApp = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excelApp.Visible = $false
            $excelApp.DisplayAlerts = $false
        }
        catch {
            throw "Excel COMの初期化に失敗しました: $($_.Exception.Message)"
        }

        # ブックを開く
        try {
            $workbook = $excelApp.Workbooks.Open($processingPath)
        }
        catch {
            throw "Excelでブックを開けません: $($_.Exception.Message)"
        }

        # Acrobat使用モード時のみ、既存プロセス停止とCOM生成を行う
        if ($useAcrobatMode) {
            try {
                $existingAcrobat = Get-Process -Name "Acrobat" -ErrorAction SilentlyContinue
                if ($existingAcrobat) {
                    Write-Log -Message "既存のAcrobatプロセスを終了します: $($existingAcrobat.Count)個"
                    $existingAcrobat | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                }
            }
            catch {
                Write-Log -Message "既存プロセスの確認中にエラー: $($_.Exception.Message)" -LogLevel "WARN"
            }

            # Acrobat COMオブジェクト（1回生成して使い回し）
            $maxRetries = 3
            $retryCount = 0
            $comObjectsCreated = $false
            while (-not $comObjectsCreated -and $retryCount -lt $maxRetries) {
                try {
                    $acrobatApp = New-Object -ComObject AcroExch.App -ErrorAction Stop
                    $avDoc = New-Object -ComObject AcroExch.AVDoc -ErrorAction Stop
                    $comObjectsCreated = $true
                    Write-Log -Message "Acrobat COMオブジェクトを生成しました（試行回数: $($retryCount + 1)）"
                }
                catch {
                    $retryCount++
                    Write-Log -Message "COMオブジェクト生成失敗（試行回数: $retryCount/$maxRetries）: $($_.Exception.Message)" -LogLevel "WARN"
                    if ($retryCount -lt $maxRetries) { Start-Sleep -Seconds 3 } else {
                        Write-Log -Message "Acrobat初期化に失敗したため非Acrobatモードへ自動切替" -LogLevel "ERROR"
                        $useAcrobatMode = $false
                        Update-ProjectConfigUseAcrobatFalse -ConfigPath $projectConfig.ConfigPath -Raw $projectConfig.Raw
                        break
                    }
                }
            }
        }

        # 全ワークシートを個別PDFにエクスポート
        $sheetNameCount = @{}
        for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            $sheet = $workbook.Worksheets.Item($i)
            $rawName = [string]$sheet.Name
            $safeSheet = Get-SafeFileName -Name $rawName
            if ($sheetNameCount.ContainsKey($safeSheet)) { $sheetNameCount[$safeSheet]++ } else { $sheetNameCount[$safeSheet] = 1 }
            $suffix = if ($sheetNameCount[$safeSheet] -gt 1) { "_" + $sheetNameCount[$safeSheet] } else { "" }

            $tempPdfPath = Join-Path -Path $tempPdfDir -ChildPath ("$($safeSheet)$suffix.pdf")
            $finalPdfPath = Join-Path -Path $excelOutputDir -ChildPath ("$((Get-SafeFileName -Name $excelBaseName))__${safeSheet}${suffix}.pdf")

            try {
                # シートを可視化して安定して出力（VeryHidden対策）
                $originalVisible = $sheet.Visible
                if ($originalVisible -ne -1) { $sheet.Visible = -1 }
                $xlFixedFormatType = 0 # xlTypePDF
                $sheet.ExportAsFixedFormat($xlFixedFormatType, $tempPdfPath)
                Write-Log -Message "一時PDF出力: $tempPdfPath"
            }
            catch {
                Write-Log -Message "ExcelでのPDF出力に失敗: シート=$rawName, エラー=$($_.Exception.Message)" -LogLevel "WARN"
                continue
            }

            if ($useAcrobatMode -and $avDoc) {
                # 一時PDFをAcrobatで開いて最終保存
                try {
                    if ($avDoc.Open($tempPdfPath, "")) {
                        $pdDoc = $avDoc.GetPDDoc()
                        $saveSuccess = $pdDoc.Save(1, $finalPdfPath)
                        if (-not $saveSuccess) { throw "PDFの最終保存に失敗" }
                        Write-Log -Message "最終PDF保存: $finalPdfPath"
                    }
                    else {
                        throw "Acrobatで一時PDFを開けませんでした: $tempPdfPath"
                    }
                }
                catch {
                    Write-Log -Message "Acrobatでの最終保存に失敗: $($_.Exception.Message)" -LogLevel "ERROR"
                    # ここで自動フォールバック
                    $useAcrobatMode = $false
                    Update-ProjectConfigUseAcrobatFalse -ConfigPath $projectConfig.ConfigPath -Raw $projectConfig.Raw
                    # 非Acrobatモードの保存に切替
                    try {
                        Copy-Item -Path $tempPdfPath -Destination $finalPdfPath -Force
                        Write-Log -Message "非Acrobatモードで最終PDF保存（コピー）: $finalPdfPath"
                    }
                    catch {
                        Write-Log -Message "非Acrobatモード保存にも失敗: $($_.Exception.Message)" -LogLevel "ERROR"
                    }
                }
            }
            else {
                # 非Acrobatモード: Excelが出力したPDFをそのまま最終保存パスへ
                try {
                    Copy-Item -Path $tempPdfPath -Destination $finalPdfPath -Force
                    Write-Log -Message "最終PDF保存（非Acrobatモード）: $finalPdfPath"
                }
                catch {
                    Write-Log -Message "非AcrobatモードでのPDF保存に失敗: $($_.Exception.Message)" -LogLevel "ERROR"
                }
            }
            finally {
                # 可視状態を元に戻す
                try { if ($null -ne $originalVisible) { $sheet.Visible = $originalVisible } } catch {}
                if ($null -ne $pdDoc) { $pdDoc.Close(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null; $pdDoc = $null }
                if ($useAcrobatMode -and $null -ne $avDoc) { $avDoc.Close($true) | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null; $avDoc = New-Object -ComObject AcroExch.AVDoc }
                # 一時PDFを削除
                try { if (Test-Path -Path $tempPdfPath) { Remove-Item -Path $tempPdfPath -Force } } catch {}
            }
        }
    }
    catch {
        $errorMessage = "エラー発生: $($file.Name) の処理に失敗しました。詳細: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -LogLevel "ERROR"

        # --- 異常終了時のファイル移動 ---
        if (Test-Path -Path $processingPath) {
            $failedPath = Join-Path -Path $failedDirPath -ChildPath $file.Name
            Move-Item -Path $processingPath -Destination $failedPath -Force
            Write-Log -Message "ファイルをfailedディレクトリに移動しました: $failedPath" -LogLevel "ERROR"
        }
    }
    finally {
        # Excel COMクリーンアップ
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        if ($null -ne $excelApp) { try { $excelApp.Quit() } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null }

        # Acrobat COMクリーンアップ
        if ($useAcrobatMode) {
            if ($null -ne $pdDoc) { try { $pdDoc.Close() } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null }
            if ($null -ne $avDoc) { try { $avDoc.Close($true) } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null }
            if ($null -ne $acrobatApp) { try { $acrobatApp.Exit() } catch {}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null }
        }

        # --- Excel解放後のファイル移動（ロック回避のためリトライ） ---
        try {
            $finalExcelPath = Join-Path -Path $completedExcelPath -ChildPath $file.Name
            $attempt = 0
            $moved = $false
            while (-not $moved -and $attempt -lt 3) {
                try {
                    if (Test-Path -Path $processingPath) {
                        Move-Item -Path $processingPath -Destination $finalExcelPath -Force
                        Write-Log -Message "ファイル移動完了: $finalExcelPath"
                    }
                    $moved = $true
                }
                catch {
                    $attempt++
                    Start-Sleep -Milliseconds 800
                    if ($attempt -ge 3) {
                        Write-Log -Message "Excelファイルの移動に失敗しました（processingに残置）: $processingPath - エラー: $($_.Exception.Message)" -LogLevel "WARN"
                    }
                }
            }
        }
        catch {
            Write-Log -Message "Excelファイルの移動処理で予期せぬエラー: $($_.Exception.Message)" -LogLevel "WARN"
        }

        # 一時フォルダクリーンアップ
        try { if (Test-Path -Path $tempPdfDir) { Remove-Item -Path $tempPdfDir -Recurse -Force } } catch {}

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Log -Message "===== スクリプト実行終了 ====="