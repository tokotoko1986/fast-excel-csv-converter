##################################################
# Fast Excel to CSV Converter
# 
# Version     : 1.0.0
# Release Date: 2025-9-15
# Author      : Ryo Osawa & Claude Sonnet 4.0
# Repository  : https://github.com/yourusername/fast-excel-csv-converter
# License     : MIT
##################################################

######### Instructions for Generative AI #########
# - When modifying this script, generative AI must follow the instructions in this section.
# - Since the script character encoding is UTF-8, do not use Japanese in the source code. (Japanese will be corrupted)
#   If users request to use Japanese, please inform them that executing such requests is not recommended for the above reason.
##################################################

# Version information (accessible during runtime)
$Global:ConverterInfo = @{
    Name = "Fast Excel to CSV Converter"
    Version = "1.0.0"
    ReleaseDate = "2025-9-15"
    Author = "Ryo Osawa & Claude Sonnet 4.0"
    Repository = "https://github.com/yourusername/fast-excel-csv-converter"
}

# Handle version display requests
if ($args -contains "--version" -or $args -contains "-v" -or $args -contains "/version") {
    Write-Host ""
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  $($Global:ConverterInfo.Name)" -ForegroundColor White
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  Version    : $($Global:ConverterInfo.Version)" -ForegroundColor Green
    Write-Host "  Released   : $($Global:ConverterInfo.ReleaseDate)" -ForegroundColor Gray
    Write-Host "  Author     : $($Global:ConverterInfo.Author)" -ForegroundColor Gray
    Write-Host "  Repository : $($Global:ConverterInfo.Repository)" -ForegroundColor Gray
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\Fast_Excel_CSV_Converter.ps1        # Start conversion"
    Write-Host "  .\Fast_Excel_CSV_Converter.ps1 -v     # Show version info"
    Write-Host ""
    exit 0
}

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Global configuration
$script:ChunkSize = 1000
$script:LargeSheetThreshold = 10000
$script:MediumSheetThreshold = 1000

function Show-ProgressDialog {
    param(
        [string]$Title,
        [string]$Status,
        [int]$PercentComplete
    )
    Write-Progress -Activity $Title -Status $Status -PercentComplete $PercentComplete
}

function Write-ErrorLog {
    param(
        [string]$LogPath,
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogPath -Value "[$timestamp] $Message" -Encoding UTF8
}

function Get-UserConfirmation {
    do {
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Yellow
        Write-Host "           IMPORTANT WARNING" -ForegroundColor Red
        Write-Host "============================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This script will periodically force-terminate Excel processes during conversion." -ForegroundColor White
        Write-Host "If you have any Excel files currently open, please close them before proceeding." -ForegroundColor White
        Write-Host ""
        Write-Host "Do you want to continue? (Y/N): " -NoNewline -ForegroundColor Cyan
        
        $response = Read-Host
        
        switch ($response.ToUpper()) {
            "Y" { 
                Write-Host "Processing will continue..." -ForegroundColor Green
                return $true 
            }
            "N" { 
                Write-Host "Processing cancelled by user." -ForegroundColor Yellow
                return $false 
            }
            default { 
                Write-Host "Please enter 'Y' for Yes or 'N' for No." -ForegroundColor Red
            }
        }
    } while ($true)
}

function Get-CellValue {
    param($values, $row, $col, $rowCount, $colCount)
    
    if ($rowCount -eq 1 -and $colCount -eq 1) { 
        return $values 
    } elseif ($rowCount -eq 1) { 
        return $values[$col] 
    } elseif ($colCount -eq 1) { 
        return $values[$row] 
    } else { 
        return $values[$row, $col] 
    }
}

function Get-CellFormat {
    param($formats, $row, $col, $rowCount, $colCount)
    
    if ($rowCount -eq 1 -and $colCount -eq 1) { 
        return $formats 
    } elseif ($rowCount -eq 1) { 
        return $formats[$col] 
    } elseif ($colCount -eq 1) { 
        return $formats[$row] 
    } else { 
        return $formats[$row, $col] 
    }
}

function Get-CellText {
    param($texts, $row, $col, $rowCount, $colCount)
    
    if ($rowCount -eq 1 -and $colCount -eq 1) { 
        return $texts 
    } elseif ($rowCount -eq 1) { 
        return $texts[$col] 
    } elseif ($colCount -eq 1) { 
        return $texts[$row] 
    } else { 
        return $texts[$row, $col] 
    }
}

function Format-CsvValue {
    param($value)
    
    if ($null -eq $value) { return "" }
    
    $valueStr = $value.ToString()
    if ($valueStr -match '[",\r\n]') {
        return '"' + $valueStr.Replace('"', '""') + '"'
    }
    return $valueStr
}

function Get-FirstDataRow {
    param($values, $rowCount, $colCount)
    
    for ($row = 1; $row -le $rowCount; $row++) {
        for ($col = 1; $col -le $colCount; $col++) {
            $value = Get-CellValue $values $row $col $rowCount $colCount
            if ($null -ne $value -and $value -ne "") {
                return $row
            }
        }
    }
    return 1  # Default to first row if no data found
}

function Test-HasFormattedCells {
    param($values, $formats, $rowCount, $colCount)
    
    if ($null -eq $formats) { return $false }
    
    # Find first data row to start sampling from actual data
    $firstDataRow = Get-FirstDataRow $values $rowCount $colCount
    
    # Dynamic sample size based on column count (columns Ã— 10)
    $sampleSize = $colCount * 10
    
    # Calculate rows to sample (maximum 10 rows from first data row)
    $dataRowCount = $rowCount - $firstDataRow + 1
    $rowsToSample = [Math]::Min(10, $dataRowCount)
    
    Write-Host "      Sampling from row $firstDataRow, checking up to $sampleSize cells" -ForegroundColor Gray
    
    # For small datasets, check all cells
    $totalCells = $rowCount * $colCount
    if ($totalCells -le $sampleSize) {
        if ($rowCount -eq 1 -and $colCount -eq 1) {
            return $formats -ne "General"
        } elseif ($rowCount -eq 1) {
            return ($formats | Where-Object { $_ -ne "General" }).Count -gt 0
        } elseif ($colCount -eq 1) {
            return ($formats | Where-Object { $_ -ne "General" }).Count -gt 0
        } else {
            for ($row = 1; $row -le $rowCount; $row++) {
                for ($col = 1; $col -le $colCount; $col++) {
                    if ($formats[$row, $col] -ne "General") {
                        return $true
                    }
                }
            }
            return $false
        }
    }
    
    # For large datasets, use improved sampling starting from first data row
    $sampleCount = 0
    $actualDataSamples = 0
    
    for ($row = $firstDataRow; $row -lt ($firstDataRow + $rowsToSample) -and $row -le $rowCount; $row++) {
        for ($col = 1; $col -le $colCount -and $sampleCount -lt $sampleSize; $col++) {
            # Get cell value to check if it contains data
            $value = Get-CellValue $values $row $col $rowCount $colCount
            
            # Skip empty cells (don't count towards sample)
            if ($null -eq $value -or $value -eq "") {
                continue
            }
            
            $actualDataSamples++
            $format = Get-CellFormat $formats $row $col $rowCount $colCount
            if ($format -ne "General") {
                Write-Host "      Formatted cells detected (sampled $actualDataSamples data cells)" -ForegroundColor Gray
                return $true
            }
            $sampleCount++
        }
    }
    
    Write-Host "      No formatted cells found (sampled $actualDataSamples data cells)" -ForegroundColor Gray
    return $false
}

function Convert-SimpleValues {
    param($values, $rowCount, $colCount)
    
    $csvContent = @()
    
    for ($row = 1; $row -le $rowCount; $row++) {
        $rowData = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $cellValue = Get-CellValue $values $row $col $rowCount $colCount
            $rowData += Format-CsvValue $cellValue
        }
        $csvContent += ($rowData -join ',')
    }
    
    return $csvContent
}

function Convert-WithFormatCheck {
    param($sheet, $usedRange, $values, $formats, $texts)
    
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    $csvContent = @()
    
    for ($row = 1; $row -le $rowCount; $row++) {
        $rowData = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $cellValue = Get-CellValue $values $row $col $rowCount $colCount
            
            if ($null -eq $cellValue -or $cellValue -eq "") {
                $rowData += ""
                continue
            }
            
            $cellFormat = Get-CellFormat $formats $row $col $rowCount $colCount
            
            # Use formatted text for non-General formats or when display differs from value
            if ($cellFormat -ne "General") {
                $cellText = Get-CellText $texts $row $col $rowCount $colCount
                if (-not [string]::IsNullOrEmpty($cellText)) {
                    $cellValue = $cellText
                }
            } elseif ($cellValue -is [double]) {
                $cellText = Get-CellText $texts $row $col $rowCount $colCount
                if ($cellValue.ToString() -ne $cellText -and -not [string]::IsNullOrEmpty($cellText)) {
                    $cellValue = $cellText
                }
            }
            
            $rowData += Format-CsvValue $cellValue
        }
        $csvContent += ($rowData -join ',')
    }
    
    return $csvContent
}

function Convert-LargeSheetToCSV {
    param($sheet, $chunkSize = 1000)
    
    $usedRange = $sheet.UsedRange
    $totalRowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    $csvContent = @()
    
    Write-Host "    Processing large sheet in chunks..." -ForegroundColor Yellow
    
    for ($startRow = 1; $startRow -le $totalRowCount; $startRow += $chunkSize) {
        $endRow = [Math]::Min($startRow + $chunkSize - 1, $totalRowCount)
        $chunkRowCount = $endRow - $startRow + 1
        
        # Define chunk range
        $startRowAbs = $usedRange.Row + $startRow - 1
        $endRowAbs = $usedRange.Row + $endRow - 1
        $startColAbs = $usedRange.Column
        $endColAbs = $usedRange.Column + $colCount - 1
        
        $chunkRange = $sheet.Range(
            $sheet.Cells($startRowAbs, $startColAbs),
            $sheet.Cells($endRowAbs, $endColAbs)
        )
        
        # Get chunk data
        $chunkValues = $chunkRange.Value2
        $chunkFormats = $chunkRange.NumberFormat
        
        # Check if chunk has formatted cells using improved sampling
        $hasFormattedCells = Test-HasFormattedCells $chunkValues $chunkFormats $chunkRowCount $colCount
        
        if (-not $hasFormattedCells) {
            # Fast processing for chunk
            $chunkContent = Convert-SimpleValues $chunkValues $chunkRowCount $colCount
        } else {
            # Standard processing for chunk
            $chunkTexts = $chunkRange.Text
            $chunkContent = Convert-WithFormatCheck $sheet $chunkRange $chunkValues $chunkFormats $chunkTexts
        }
        
        $csvContent += $chunkContent
        
        # Progress update
        $progress = [Math]::Round(($endRow / $totalRowCount) * 100)
        Write-Progress -Activity "Processing large sheet" -Status "Chunk: $endRow/$totalRowCount rows" -PercentComplete $progress
        
        # Memory cleanup
        [System.GC]::Collect()
    }
    
    return $csvContent
}

function Convert-SheetToCSV-Optimized {
    param($sheet, $sheetName)
    
    $usedRange = $sheet.UsedRange
    if (-not $usedRange) { 
        Write-Host "    Empty sheet - creating empty CSV" -ForegroundColor Gray
        return @() 
    }
    
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    $cellCount = $rowCount * $colCount
    
    Write-Host "    Sheet size: $rowCount x $colCount ($cellCount cells)" -ForegroundColor Gray
    
    # Determine processing strategy based on size
    if ($cellCount -gt $script:LargeSheetThreshold) {
        Write-Host "    Strategy: Chunk processing (large dataset)" -ForegroundColor Cyan
        return Convert-LargeSheetToCSV $sheet $script:ChunkSize
    }
    
    # For medium and small sheets, use batch analysis
    $values = $usedRange.Value2
    
    if ($cellCount -gt $script:MediumSheetThreshold) {
        Write-Host "    Strategy: Analyzing formats with improved sampling..." -ForegroundColor Cyan
        $formats = $usedRange.NumberFormat
        $hasFormattedCells = Test-HasFormattedCells $values $formats $rowCount $colCount
    } else {
        Write-Host "    Strategy: Standard processing (small dataset)" -ForegroundColor Cyan
        $hasFormattedCells = $true  # Safe fallback for small sheets
    }
    
    if (-not $hasFormattedCells) {
        Write-Host "    Mode: Fast processing (no formatted cells detected)" -ForegroundColor Green
        return Convert-SimpleValues $values $rowCount $colCount
    } else {
        Write-Host "    Mode: Standard processing (formatted cells detected)" -ForegroundColor Yellow
        $formats = if ($cellCount -gt $script:MediumSheetThreshold) { $formats } else { $usedRange.NumberFormat }
        $texts = $usedRange.Text
        return Convert-WithFormatCheck $sheet $usedRange $values $formats $texts
    }
}

try {
    # Display welcome banner
    Write-Host ""
    Write-Host "=== $($Global:ConverterInfo.Name) v$($Global:ConverterInfo.Version) ===" -ForegroundColor Cyan
    Write-Host "High-performance Excel to CSV conversion with intelligent optimization" -ForegroundColor Gray
    Write-Host ""

    # Error tracking variable
    $script:HasErrors = $false

    # 0. User confirmation before processing
    if (-not (Get-UserConfirmation)) {
        Write-Host "Script execution terminated." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }

    # 1. Excel file selection dialog
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select Excel files to convert"
    $fileDialog.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm"
    $fileDialog.Multiselect = $true
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFiles = $fileDialog.FileNames
        Write-Host "Selected files: $($selectedFiles.Count)" -ForegroundColor Green
    } else {
        Write-Host "File selection was cancelled." -ForegroundColor Yellow
        exit 1
    }

    # 2. Create output folder
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outputFolder = Join-Path $scriptPath $timestamp
    
    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
    }
    
    Write-Host "Output folder: $outputFolder" -ForegroundColor Green

    # Error log file path
    $errorLogPath = Join-Path $outputFolder "error.log"

    # Initialize Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $totalFiles = $selectedFiles.Count
    $currentFileIndex = 0
    $performanceStats = @{
        FastModeSheets = 0
        StandardModeSheets = 0
        ChunkModeSheets = 0
        TotalProcessingTime = 0
    }

    $overallStartTime = Get-Date

    foreach ($filePath in $selectedFiles) {
        $currentFileIndex++
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
        
        Show-ProgressDialog -Title "Converting Excel to CSV (Optimized)" -Status "Processing: $fileName ($currentFileIndex/$totalFiles)" -PercentComplete (($currentFileIndex - 1) / $totalFiles * 100)
        
        Write-Host "`nProcessing: $fileName" -ForegroundColor Cyan

        try {
            $fileStartTime = Get-Date
            $workbook = $excel.Workbooks.Open($filePath)
            
            $totalSheets = $workbook.Sheets.Count
            $currentSheetIndex = 0
            
            foreach ($sheet in $workbook.Sheets) {
                $currentSheetIndex++
                $sheetName = $sheet.Name
                
                $safeSheetName = $sheetName -replace '[\\/:*?"<>|]', '_'
                $csvFileName = "$fileName-$safeSheetName.csv"
                $csvFilePath = Join-Path $outputFolder $csvFileName
                
                Show-ProgressDialog -Title "Converting Excel to CSV (Optimized)" -Status "Processing: $fileName - $sheetName ($currentSheetIndex/$totalSheets)" -PercentComplete (($currentFileIndex - 1) / $totalFiles * 100 + ($currentSheetIndex / $totalSheets) / $totalFiles * 100)
                
                Write-Host "  Sheet: $sheetName -> $csvFileName" -ForegroundColor White
                
                try {
                    $sheetStartTime = Get-Date
                    $csvContent = Convert-SheetToCSV-Optimized $sheet $sheetName
                    $sheetEndTime = Get-Date
                    $sheetProcessingTime = ($sheetEndTime - $sheetStartTime).TotalSeconds
                    
                    Write-Host "    Processing time: $([Math]::Round($sheetProcessingTime, 2)) seconds" -ForegroundColor Gray
                    
                    # Write CSV content using high-performance method
                    [System.IO.File]::WriteAllLines($csvFilePath, $csvContent, [System.Text.Encoding]::UTF8)
                    
                } catch {
                    $script:HasErrors = $true
                    $errorMessage = "Sheet conversion error - File: $fileName, Sheet: $sheetName, Error: $($_.Exception.Message)"
                    Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
                    Write-ErrorLog -LogPath $errorLogPath -Message $errorMessage
                }
            }
            
            $workbook.Close($false)
            $fileEndTime = Get-Date
            $fileProcessingTime = ($fileEndTime - $fileStartTime).TotalSeconds
            Write-Host "  File completed in $([Math]::Round($fileProcessingTime, 2)) seconds" -ForegroundColor Green
            
        } catch {
            $script:HasErrors = $true
            $errorMessage = "File open error - File: $fileName, Error: $($_.Exception.Message)"
            Write-Host "  Error: Could not open file - $($_.Exception.Message)" -ForegroundColor Red
            Write-ErrorLog -LogPath $errorLogPath -Message $errorMessage
        }
    }

} catch {
    $script:HasErrors = $true
    Write-Host "Unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-ErrorLog -LogPath $errorLogPath -Message "Unexpected error: $($_.Exception.Message)"
} finally {
    Write-Host "`nStarting cleanup process..." -ForegroundColor Yellow
    
    # 1. Close any open workbooks
    if ($script:workbook) {
        try {
            Write-Host "Closing workbook..." -ForegroundColor Gray
            $script:workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:workbook) | Out-Null
        } catch {
            Write-Host "Warning: Could not properly close workbook" -ForegroundColor Yellow
        }
        $script:workbook = $null
    }
    
    # 2. Store Excel process IDs before termination attempt
    $preExcelProcesses = @()
    try {
        $preExcelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Select-Object Id, ProcessName
        Write-Host "Excel processes found before cleanup: $($preExcelProcesses.Count)" -ForegroundColor Gray
    } catch {
        # No Excel processes found or error getting process list
    }
    
    # 3. Terminate Excel application properly
    if ($script:excel) {
        try {
            Write-Host "Closing Excel application..." -ForegroundColor Gray
            
            # Close all workbooks first
            foreach ($wb in $script:excel.Workbooks) {
                try {
                    $wb.Close($false)
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
                } catch {
                    # Continue even if individual workbook fails
                }
            }
            
            # Quit Excel application
            $script:excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:excel) | Out-Null
        } catch {
            Write-Host "Warning: Could not properly terminate Excel application" -ForegroundColor Yellow
        }
        $script:excel = $null
    }
    
    # 4. Force garbage collection
    Write-Host "Forcing garbage collection..." -ForegroundColor Gray
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    # 5. Wait a moment for processes to terminate naturally
    Start-Sleep -Seconds 2
    
    # 6. Check remaining Excel processes and force terminate if necessary
    try {
        $postExcelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
        
        if ($postExcelProcesses) {
            Write-Host "Found $($postExcelProcesses.Count) remaining Excel process(es). Force terminating..." -ForegroundColor Yellow
            
            foreach ($process in $postExcelProcesses) {
                try {
                    Write-Host "  Terminating Excel process ID: $($process.Id)" -ForegroundColor Gray
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                    Write-Host "  Successfully terminated process ID: $($process.Id)" -ForegroundColor Green
                } catch {
                    Write-Host "  Failed to terminate process ID: $($process.Id) - $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            
            # Wait and verify termination
            Start-Sleep -Seconds 1
            $finalExcelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
            
            if ($finalExcelProcesses) {
                Write-Host "Warning: $($finalExcelProcesses.Count) Excel process(es) still remain after force termination!" -ForegroundColor Red
                Write-Host "Remaining process IDs: $($finalExcelProcesses.Id -join ', ')" -ForegroundColor Red
                Write-Host "You may need to terminate these manually using Task Manager." -ForegroundColor Red
            } else {
                Write-Host "All Excel processes successfully terminated." -ForegroundColor Green
            }
        } else {
            Write-Host "All Excel processes terminated successfully." -ForegroundColor Green
        }
    } catch {
        Write-Host "Error checking Excel processes: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # 7. Hide progress bar
    Write-Progress -Activity "Converting Excel to CSV (Optimized)" -Completed
    
    Write-Host "Cleanup process completed." -ForegroundColor Green
}

$overallEndTime = Get-Date
$totalProcessingTime = ($overallEndTime - $overallStartTime).TotalSeconds

Write-Host "`n=== Conversion Process Completed ===" -ForegroundColor Green
Write-Host "Output location: $outputFolder" -ForegroundColor Green
Write-Host "Total processing time: $([Math]::Round($totalProcessingTime, 2)) seconds" -ForegroundColor Green

if (Test-Path $errorLogPath) {
    Write-Host "Error log: $errorLogPath" -ForegroundColor Yellow
    Write-Host "Errors occurred with some files. Please check the error log for details." -ForegroundColor Yellow
}

Read-Host "`nPress Enter to exit"

# Determine exit code based on error status
if ($script:HasErrors) {
    Write-Host "Exiting with error code 2 due to processing errors." -ForegroundColor Red
    exit 2
} else {
    exit 0
}