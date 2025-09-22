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
foreach ($file in $excelFiles) {
    $processingPath = Join-Path -Path $processingDirPath -ChildPath $file.Name
    $pdfPath = Join-Path -Path $completedPdfPath -ChildPath "$($file.BaseName).pdf"

    # COMオブジェクト変数を初期化
    $acrobatApp = $null
    $avDoc = $null
    $pdDoc = $null

    try {
        Write-Log -Message "処理開始: $($file.Name)"

        # ファイルを作業中ディレクトリへ移動
        Move-Item -Path $file.FullName -Destination $processingPath -Force

        # --- Acrobat COMオブジェクトの生成 ---
        # COMエラーの回避: 既存のAcrobatプロセスを確認・終了
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

        # COMオブジェクトの生成（リトライ機能付き）
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
                if ($retryCount -lt $maxRetries) {
                    Start-Sleep -Seconds 3
                }
                else {
                    throw "COMオブジェクトの生成に失敗しました: $($_.Exception.Message)"
                }
            }
        }

        # --- PDF変換実行 ---
        if ($avDoc.Open($processingPath, "")) {
            $pdDoc = $avDoc.GetPDDoc()
            # Saveメソッドの第一引数 '1' は PDSaveFull (フル保存) を意味する
            $saveSuccess = $pdDoc.Save(1, $pdfPath)
            if (-not $saveSuccess) {
                # SaveメソッドがFalseを返した場合、例外を投げてcatchブロックへ
                throw "PDFファイルの保存に失敗しました。"
            }
            Write-Log -Message "変換成功: $($file.Name) -> $($pdfPath)"

            # --- 正常終了時のファイル移動 ---
            $finalExcelPath = Join-Path -Path $completedExcelPath -ChildPath $file.Name
            Move-Item -Path $processingPath -Destination $finalExcelPath -Force
            Write-Log -Message "ファイル移動完了: $finalExcelPath"
        }
        else {
            throw "AcrobatでExcelファイルを開けませんでした。ファイルが破損しているか、Acrobatが対応していない可能性があります。"
        }
    }
    catch {
        $errorMessage = "エラー発生: $($file.Name) の処理に失敗しました。詳細: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -LogLevel "ERROR"

        # --- 異常終了時のファイル移動 ---
        # 処理中ファイルをfailedディレクトリへ移動
        if (Test-Path -Path $processingPath) {
            $failedPath = Join-Path -Path $failedDirPath -ChildPath $file.Name
            Move-Item -Path $processingPath -Destination $failedPath -Force
            Write-Log -Message "ファイルをfailedディレクトリに移動しました: $failedPath" -LogLevel "ERROR"
        }
    }
    finally {
        # --- COMオブジェクトのクリーンアップ ---
        # このブロックは成功・失敗に関わらず必ず実行される
        if ($null -ne $pdDoc) {
            $pdDoc.Close()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        }
        if ($null -ne $avDoc) {
            $avDoc.Close($true) # $trueはウィンドウを強制的に閉じる
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        }
        if ($null -ne $acrobatApp) {
            $acrobatApp.Exit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null
        }
        # メモリ解放を促す
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Log -Message "===== スクリプト実行終了 ====="