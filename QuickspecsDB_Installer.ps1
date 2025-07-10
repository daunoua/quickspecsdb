# ========== Quickspecs Download Script ==========
# Version: v1.2
# Purpose: Download and organize Personal Systems Quickspecs locally

Clear-Host
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "       🔽 Quickspecs Download Script (v1.2) 🔽" -ForegroundColor Green
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "`nThis script will download active Personal Systems Quickspecs"
Write-Host "and organize them into local folders by category."
Write-Host "Files will be saved to: C:\Quickspecs"
Write-Host "`nCleaning any existing files and preparing environment..." -ForegroundColor Yellow
Start-Sleep -Seconds 2

# ========== SETUP ==========

# Define folders and URLs
$localFolder = "C:\Quickspecs"
$baseAppDataFolder = Join-Path $env:APPDATA "QuickspecsDB"
$logFolder = Join-Path $baseAppDataFolder "Logs"
$localExcelPath = Join-Path $baseAppDataFolder "QuickspecsDB.xlsx"
$localUpdaterPath = Join-Path $baseAppDataFolder "QuickspecsDB_Updater.ps1"
$shortcutPath = Join-Path $localFolder "Launch QuickspecsDB Updater.lnk"

$xlsxUrl = "https://raw.githubusercontent.com/daunoua/quickspecsdb/main/QuickspecsDB.xlsx"
$updaterUrl = "https://raw.githubusercontent.com/daunoua/quickspecsdb/main/QuickspecsDB_Updater.ps1"

# Ensure required folders exist
foreach ($folder in @($localFolder, $baseAppDataFolder, $logFolder)) {
    if (-not (Test-Path $folder)) {
        Write-Host "📁 Creating folder: $folder"
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# Clean local Quickspecs PDF directory
Write-Host "🧹 Cleaning all contents under: $localFolder..." -ForegroundColor Yellow
Get-ChildItem -Path $localFolder -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "✅ Cleaned: $localFolder"

# Download Excel file
Write-Host "`n📥 Downloading Quickspecs Excel database..."
try {
    Invoke-WebRequest -Uri $xlsxUrl -OutFile $localExcelPath -ErrorAction Stop
    Write-Host "✅ Downloaded: $localExcelPath"
} catch {
    Write-Error "❌ Failed to download Excel file: $($_.Exception.Message)"
    exit 1
}

# Download Updater script
Write-Host "📥 Downloading QuickspecsDB_Updater.ps1..."
try {
    Invoke-WebRequest -Uri $updaterUrl -OutFile $localUpdaterPath -ErrorAction Stop
    Write-Host "✅ Downloaded: $localUpdaterPath"
} catch {
    Write-Error "❌ Failed to download updater script: $($_.Exception.Message)"
    exit 1
}

# Create shortcut to updater script in C:\Quickspecs
Write-Host "📎 Creating shortcut: Launch QuickspecsDB Updater"
try {
    $wShell = New-Object -ComObject WScript.Shell
    $shortcut = $wShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$localUpdaterPath`""
    $shortcut.WorkingDirectory = $baseAppDataFolder
    $shortcut.WindowStyle = 1
    $shortcut.IconLocation = "powershell.exe,0"
    $shortcut.Save()
    Write-Host "✅ Shortcut created at: $shortcutPath"


# Ask user if they want to create a scheduled task for daily updates
$createTask = Read-Host "`n🛠️  Do you want to configure a daily task to automatically run the updater? (Y/N)"
if ($createTask -match '^[Yy]') {
    $taskName = "QuickspecsDB_Updater"
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        $overwrite = Read-Host "⚠️  A scheduled task named '$taskName' already exists. Do you want to overwrite it? (Y/N)"
        if ($overwrite -notmatch '^[Yy]') {
            Write-Host "⏭️  Existing task retained. No changes made." -ForegroundColor Yellow
            return
        } else {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host "🗑️  Existing task removed. Proceeding to create a new one..." -ForegroundColor Yellow
        }
    }

    try {
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$localUpdaterPath`""
        $trigger = New-ScheduledTaskTrigger -Daily -At 11:00AM
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Description "Runs QuickspecsDB_Updater.ps1 daily"

        Write-Host "✅ Scheduled task '$taskName' has been created successfully." -ForegroundColor Green
    } catch {
        Write-Warning "⚠️  Failed to create scheduled task: $($_.Exception.Message)"
    }
} else {
    Write-Host "⏭️  Skipping task scheduler setup." -ForegroundColor Yellow
}






} catch {
    Write-Warning "⚠️ Failed to create shortcut: $($_.Exception.Message)"
}


# ========== STEP 2: Load Excel ==========
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($localExcelPath)
    $sheet = $workbook.Sheets.Item("QuickSpecsList")
} catch {
    Write-Error "❌ Unable to open QuickSpecsList in Excel workbook"
    $excel.Quit()
    exit 1
}

$usedRange = $sheet.UsedRange
$rowCount = $usedRange.Rows.Count
$colCount = $usedRange.Columns.Count

# ========== STEP 3: Identify Headers ==========
$headers = @{ }
for ($col = 1; $col -le $colCount; $col++) {
    $header = $usedRange.Cells.Item(1, $col).Text.Trim()
    $headers[$header] = $col
}

# Required columns
$required = @("docID", "Title", "URL", "maincategory", "category", "generation", "Skip")
foreach ($h in $required) {
    if (-not $headers.ContainsKey($h)) {
        Write-Error "❌ Missing column: $h"
        $excel.Quit()
        exit 1
    }
}

# ========== STEP 4: Process Rows ==========
$success = 0
$failures = @()
$total = $rowCount - 1
$index = 1

for ($row = 2; $row -le $rowCount; $row++) {
    $skip = $usedRange.Cells.Item($row, $headers["Skip"]).Text.Trim()
    if ($skip -eq "Yes") {
        continue
    }

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

    $path = Join-Path -Path $localFolder -ChildPath "$safeMain\$safeCat"
    if ($safeGen -ne "") {
        $path = Join-Path $path $safeGen
    }

    New-Item -ItemType Directory -Path $path -Force | Out-Null

    $filename = "$safeTitle - $safeDocID.pdf"
    $pdfPath = Join-Path $path $filename

    Write-Host "📄 Downloading [$index/$total]: $filename"

    try {
        Invoke-WebRequest -Uri $url -OutFile $pdfPath -UseDefaultCredentials -ErrorAction Stop
        Write-Host "✅ Success: $filename"
        $success++
    } catch {
        Write-Warning "❌ Failed to download '$title' (DocID: $docID) from $url"
        $failures += [PSCustomObject]@{
            docID = $docID
            Title = $title
            URL = $url
            Error = $_.Exception.Message
        }
    }

    $index++
}

# ========== STEP 5: Cleanup ==========
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel


# ========== STEP 6: Summary ==========
Write-Host "`n========== SUMMARY =========="
Write-Host "Total Quickspecs processed: $($rowCount - 1)"
Write-Host "✅ Successfully downloaded: $success"
Write-Host "❌ Failed downloads: $($failures.Count)"

if ($failures.Count -gt 0) {
    $failLog = Join-Path $logFolder "Quickspecs_Failures_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $failures | Export-Csv -Path $failLog -NoTypeInformation
    Write-Host "Failure log saved to: $failLog"
}

Write-Host "`n📢 REMINDER:" -ForegroundColor Yellow
Write-Host "Quickspecs documents are updated regularly."
Write-Host "To keep your local repository up to date, use the 'Launch QuickspecsDB Updater' shortcut located in the C:\Quickspecs folder."
Write-Host "You may also configure it in Windows Task Scheduler for automatic daily updates."

Write-Host "`n💬 If you encounter issues or want to report missing Quickspecs, please contact:" -ForegroundColor Cyan
Write-Host "ww.presales.expert.team@hp.com"


# ========== STEP 7: Open Folder in Explorer ==========
Start-Process "explorer.exe" $localFolder