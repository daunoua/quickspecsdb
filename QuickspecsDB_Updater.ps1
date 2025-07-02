# ========== QuickspecsDB_Updater.ps1 ==========
# Version: v1.2 (Modified to use AppData for logs and DB)

Clear-Host
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "    🔄 QuickspecsDB Updater Script (v1.2) 🔄" -ForegroundColor Green
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "Checking for updated Quickspecs PDFs..." -ForegroundColor Yellow

# Base folders
$localFolder = "C:\Quickspecs"
$baseAppDataFolder = Join-Path $env:APPDATA "QuickspecsDB"
$logFolder = Join-Path $baseAppDataFolder "Logs"

# Ensure folders exist
foreach ($folder in @($baseAppDataFolder, $logFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# Excel DB and log path
$localExcelPath = Join-Path $baseAppDataFolder "QuickspecsDB.xlsx"
$xlsxUrl = "https://raw.githubusercontent.com/daunoua/quickspecsdb/main/QuickspecsDB.xlsx"
$logPath = Join-Path $logFolder "Quickspecs_Update_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# ===== File-lock safe Excel download =====
Write-Host "`n📥 Checking and downloading latest Excel file..."
$canWrite = $true

if (Test-Path $localExcelPath) {
    try {
        $stream = [System.IO.File]::Open($localExcelPath, 'Open', 'ReadWrite', 'None')
        $stream.Close()
    } catch {
        Write-Host "❌ Excel file is locked or open. Cannot overwrite: $localExcelPath" -ForegroundColor Red
        Add-Content $logPath "`n[$(Get-Date -Format 'u')] ❌ Excel file locked. Skipping download."
        exit 1
    }
}

try {
    Invoke-WebRequest -Uri $xlsxUrl -OutFile $localExcelPath -ErrorAction Stop
    Write-Host "✅ Excel downloaded to: $localExcelPath"
    Add-Content $logPath "`n[$(Get-Date -Format 'u')] Excel downloaded successfully."
} catch {
    Write-Host "❌ Failed to download Excel file: $($_.Exception.Message)" -ForegroundColor Red
    Add-Content $logPath "`n[$(Get-Date -Format 'u')] ❌ Excel download failed: $($_.Exception.Message)"
    exit 1
}

# ===== Load Excel =====
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($localExcelPath)
    $sheet = $workbook.Sheets.Item("QuickSpecsList")
} catch {
    Write-Host "❌ Unable to open Excel QuickSpecsList." -ForegroundColor Red
    Add-Content $logPath "`n[$(Get-Date -Format 'u')] ❌ Excel open failed."
    $excel.Quit()
    exit 1
}

$usedRange = $sheet.UsedRange
$rowCount = $usedRange.Rows.Count
$colCount = $usedRange.Columns.Count

# ===== Identify Headers =====
$headers = @{}
for ($col = 1; $col -le $colCount; $col++) {
    $header = $usedRange.Cells.Item(1, $col).Text.Trim()
    $headers[$header] = $col
}

$required = @("docID", "Title", "URL", "maincategory", "category", "generation", "Skip")
foreach ($h in $required) {
    if (-not $headers.ContainsKey($h)) {
        Write-Host "❌ Missing required column: $h" -ForegroundColor Red
        Add-Content $logPath "`n❌ Missing column: $h"
        $excel.Quit()
        exit 1
    }
}

# ===== Process Entries =====
$success = 0
$skipped = 0
$failures = 0
$total = $rowCount - 1
$index = 1

Write-Host "`n🔍 Starting scan of $total entries..."

for ($row = 2; $row -le $rowCount; $row++) {
    $skip = $usedRange.Cells.Item($row, $headers["Skip"]).Text.Trim()
    if ($skip -eq "Yes") { continue }

    $docID = $usedRange.Cells.Item($row, $headers["docID"]).Text.Trim()
    $title = $usedRange.Cells.Item($row, $headers["Title"]).Text.Trim()
    $url = $usedRange.Cells.Item($row, $headers["URL"]).Text.Trim()
    $main = $usedRange.Cells.Item($row, $headers["maincategory"]).Text.Trim()
    $cat = $usedRange.Cells.Item($row, $headers["category"]).Text.Trim()
    $gen = $usedRange.Cells.Item($row, $headers["generation"]).Text.Trim()

    $safeMain = $main -replace '[\\\/:*?"<>|]', '_'
    $safeCat = $cat -replace '[\\\/:*?"<>|]', '_'
    $safeGen = $gen -replace '[\\\/:*?"<>|]', '_'
    $safeTitle = $title -replace '[\\\/:*?"<>|]', '_'
    $safeDocID = $docID -replace '[\\\/:*?"<>|]', '_'

    $path = Join-Path $localFolder "$safeMain\$safeCat"
    if ($safeGen -ne "") {
        $path = Join-Path $path $safeGen
    }

    New-Item -ItemType Directory -Path $path -Force | Out-Null
    $filename = "$safeTitle - $safeDocID.pdf"
    $pdfPath = Join-Path $path $filename

    Write-Host "`n[$index/$total] Checking: ${filename}"

    $download = $true
    if (Test-Path $pdfPath) {
        try {
            $head = Invoke-WebRequest -Uri $url -Method Head -ErrorAction Stop
            $remoteSize = [int64]$head.Headers["Content-Length"]
            $localSize = (Get-Item $pdfPath).Length

            if ($remoteSize -eq $localSize) {
                Write-Host "⏭️  Skipped (up-to-date)"
                Add-Content $logPath "⏭️  Skipped: ${filename}"
                $download = $false
                $skipped++
            } else {
                Write-Host "🔁 Updating (remote: $remoteSize, local: $localSize)"
            }
        } catch {
            Write-Host "⚠️  Could not retrieve size info. Will re-download."
        }
    }

    if ($download) {
        try {
            Invoke-WebRequest -Uri $url -OutFile $pdfPath -UseDefaultCredentials -ErrorAction Stop
            Write-Host "✅ Downloaded: ${filename}" -ForegroundColor Green
            Add-Content $logPath "✅ Downloaded: ${filename}"
            $success++
        } catch {
            Write-Host "❌ Failed to download ${filename}: $($_.Exception.Message)" -ForegroundColor Red
            Add-Content $logPath "❌ Failed: ${filename}"
            $failures++
        }
    }

    $index++
}

# ===== Cleanup =====
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel

# ===== Summary =====
Write-Host "`n📊 Update Complete"
Write-Host "✅ Downloaded: $success"
Write-Host "⏭️  Skipped: $skipped"
Write-Host "❌ Failed: $failures"
Add-Content $logPath "`n========== SUMMARY =========="
Add-Content $logPath "✅ Downloaded: $success"
Add-Content $logPath "⏭️  Skipped: $skipped"
Add-Content $logPath "❌ Failed: $failures"
Write-Host "📝 Log saved to: $logPath"