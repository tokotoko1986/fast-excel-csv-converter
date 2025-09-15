##################################################
# Fast Excel to CSV Converter
# 
# Version     : 1.0.0
# Release Date: 2025-9-15
# Author      : Ryo Osawa & Claude Sonnet 4.0
# Repository  : https://github.com/yourusername/fast-excel-csv-converter
# License     : MIT
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
        Write-Host "This script may force-terminate Excel processes during conversion." -ForegroundColor White
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

function Get-ConversionMode {
    do {
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Yellow
        Write-Host "           SELECT CONVERSION MODE" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "1. Normal Mode (Formats preserved)" -ForegroundColor White
        Write-Host "   - Uses the .Text property to get formatted values (e.g., dates, currencies)." -ForegroundColor Gray
        Write-Host "   - Slower, but preserves all cell formatting." -ForegroundColor Gray
        Write-Host ""
        Write-Host "2. High-Speed Mode (Formats NOT preserved)" -ForegroundColor White
        Write-Host "   - Uses the .Value2 property to read raw cell values at once." -ForegroundColor Gray
        Write-Host "   - Much faster, but formats like dates may appear as serial numbers." -ForegroundColor Gray
        Write-Host ""
        Write-Host "Select mode (1 or 2): " -NoNewline -ForegroundColor Cyan
        
        $response = Read-Host
        
        switch ($response) {
            "1" { 
                Write-Host ""
                Write-Host "Selected mode: Normal (Formats preserved)" -ForegroundColor Green
                return @{
                    Mode = "Normal"
                    Description = "Normal (Formats preserved)"
                }
            }
            "2" { 
                Write-Host ""
                Write-Host "Selected mode: High-Speed (Formats NOT preserved)" -ForegroundColor Green
                return @{
                    Mode = "HighSpeed"
                    Description = "High-Speed (Formats NOT preserved)"
                }
            }
            default { 
                Write-Host "Please enter '1' for Normal Mode or '2' for High-Speed Mode." -ForegroundColor Red
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

function Format-CsvValue {
    param($value)
    
    if ($null -eq $value) { return "" }
    
    $valueStr = $value.ToString()
    if ($valueStr -match '[",\r\n]') {
        return '"' + $valueStr.Replace('"', '""') + '"'
    }
    return $valueStr
}

function Convert-SheetToCSV {
    param($sheet, $sheetName, $conversionMode)
    
    # Get the last cell with data using SpecialCells
    try {
        $lastCell = $sheet.Cells.SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeLastCell)
        $maxRow = $lastCell.Row
        $maxCol = $lastCell.Column
    } catch {
        # SpecialCells throws exception if no data exists in the sheet
        Write-Host "    Empty sheet - creating empty CSV" -ForegroundColor Gray
        return @()
    }
    
    Write-Host "    Processing range: A1 to $($sheet.Cells($maxRow, $maxCol).Address(0, 0))" -ForegroundColor Gray
    Write-Host "    Processing cell by cell using $($conversionMode.Description) mode..." -ForegroundColor Cyan
    
    $csvContent = @()
    
    for ($row = 1; $row -le $maxRow; $row++) {
        $rowData = @()
        
        for ($col = 1; $col -le $maxCol; $col++) {
            # Get individual cell text (using .Text property for format preservation)
            $cellText = $sheet.Cells($row, $col).Text
            $rowData += Format-CsvValue $cellText
        }
        
        $csvContent += ($rowData -join ',')
    }
    
    Write-Host "    Processed $maxRow rows, $maxCol columns" -ForegroundColor Gray
    return $csvContent
}

function Convert-SheetToCSV-Fast {
    param($sheet, $sheetName)
    
    Write-Host "    Processing using High-Speed mode (UsedRange + Value2)..." -ForegroundColor Cyan
    
    # Get UsedRange
    $usedRange = $sheet.UsedRange
    
    # Handle completely empty sheets
    if (-not $usedRange) {
        Write-Host "    Completely empty sheet - creating single empty cell CSV" -ForegroundColor Gray
        return @("")
    }
    
    # Get UsedRange boundaries
    $startRow = $usedRange.Row
    $startCol = $usedRange.Column
    $endRow = $startRow + $usedRange.Rows.Count - 1
    $endCol = $startCol + $usedRange.Columns.Count - 1
    
    # Complete range always starts from A1
    $fullStartRow = 1
    $fullStartCol = 1
    $fullEndRow = $endRow
    $fullEndCol = $endCol
    
    Write-Host "    UsedRange: $($sheet.Cells($startRow, $startCol).Address(0, 0)) to $($sheet.Cells($endRow, $endCol).Address(0, 0))" -ForegroundColor Gray
    Write-Host "    Full range: A1 to $($sheet.Cells($fullEndRow, $fullEndCol).Address(0, 0))" -ForegroundColor Gray
    
    # Get all values from UsedRange using Value2 for high performance
    $values = $usedRange.Value2
    
    # Determine if we have array or single value
    $isArray = $values -is [System.Array]
    $usedRowCount = $usedRange.Rows.Count
    $usedColCount = $usedRange.Columns.Count
    
    # Debug: Check actual array dimensions
    $arrayDimension = "N/A"
    $actualArrayLength = "N/A"
    if ($isArray) {
        $arrayDimension = $values.Rank
        if ($values.Rank -eq 1) {
            $actualArrayLength = $values.Length
        } else {
            $actualArrayLength = "$($values.GetLength(0))x$($values.GetLength(1))"
        }
    }
    
    Write-Host "    Processing $fullEndRow rows, $fullEndCol columns (UsedRange: $usedRowCount x $usedColCount)" -ForegroundColor Gray
    Write-Host "    Array info: IsArray=$isArray, Dimension=$arrayDimension, Length=$actualArrayLength" -ForegroundColor Gray
    
    $csvContent = @()
    
    # Process each row from A1 to full range
    for ($row = $fullStartRow; $row -le $fullEndRow; $row++) {
        $rowData = @()
        
        for ($col = $fullStartCol; $col -le $fullEndCol; $col++) {
            $cellValue = $null
            
            # Check if current cell is within UsedRange
            if ($row -ge $startRow -and $row -le $endRow -and $col -ge $startCol -and $col -le $endCol) {
                # Within UsedRange - get actual data
                if ($isArray) {
                    # Calculate array indices (1-based)
                    $arrayRow = $row - $startRow + 1
                    $arrayCol = $col - $startCol + 1
                    
                    try {
                        if ($usedRowCount -eq 1 -and $usedColCount -eq 1) {
                            # Single cell case - but returned as array somehow
                            $cellValue = $values[1]
                        } elseif ($usedRowCount -eq 1) {
                            # Single row, multiple columns - 1D array indexed by column
                            $cellValue = $values[$arrayCol]
                        } elseif ($usedColCount -eq 1) {
                            # Multiple rows, single column - 1D array indexed by row
                            $cellValue = $values[$arrayRow]
                        } else {
                            # Multiple rows and columns - 2D array
                            $cellValue = $values[$arrayRow, $arrayCol]
                        }
                    } catch {
                        # Fallback: try different access patterns
                        try {
                            if ($values.Rank -eq 1) {
                                # 1D array - use linear index
                                $linearIndex = ($arrayRow - 1) * $usedColCount + $arrayCol
                                $cellValue = $values[$linearIndex]
                            } else {
                                # 2D array - use standard indexing
                                $cellValue = $values[$arrayRow, $arrayCol]
                            }
                        } catch {
                            Write-Host "      Warning: Could not access array at [$arrayRow, $arrayCol], using null" -ForegroundColor Yellow
                            $cellValue = $null
                        }
                    }
                } else {
                    # Single cell in UsedRange
                    $cellValue = $values
                }
            } else {
                # Outside UsedRange (leading empty rows/columns) - use empty value
                $cellValue = $null
            }
            
            $rowData += Format-CsvValue $cellValue
        }
        
        $csvContent += ($rowData -join ',')
    }
    
    Write-Host "    Processed $fullEndRow rows, $fullEndCol columns" -ForegroundColor Gray
    return $csvContent
}

try {
    # Display welcome banner
    Write-Host ""
    Write-Host "=== $($Global:ConverterInfo.Name) v$($Global:ConverterInfo.Version) ===" -ForegroundColor Cyan
    Write-Host "Excel to CSV conversion with format preservation" -ForegroundColor Gray
    Write-Host ""

    # Error tracking variable
    $script:HasErrors = $false

    # User confirmation before processing
    if (-not (Get-UserConfirmation)) {
        Write-Host "Script execution terminated." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }

    # Get conversion mode selection
    $conversionMode = Get-ConversionMode

    # Excel file selection dialog
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

    # Create output folder
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

    $overallStartTime = Get-Date

    foreach ($filePath in $selectedFiles) {
        $currentFileIndex++
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
        
        Write-Host "`nProcessing: $fileName ($currentFileIndex/$totalFiles)" -ForegroundColor Cyan

        try {
            $fileStartTime = Get-Date
            $workbook = $excel.Workbooks.Open($filePath)
            
            $totalSheets = $workbook.Sheets.Count
            $currentSheetIndex = 0
            
            foreach ($sheet in $workbook.Sheets) {
                $currentSheetIndex++
                $sheetName = $sheet.Name
                
                $safeSheetName = $sheetName -replace '[\\/:*?"<>|]', '_'
                
                # Add mode suffix to filename
                $modeSuffix = if ($conversionMode.Mode -eq "Normal") { "-normal" } else { "-highspeed" }
                $csvFileName = "$fileName-$safeSheetName$modeSuffix.csv"
                $csvFilePath = Join-Path $outputFolder $csvFileName
                
                Write-Host "  Sheet: $sheetName -> $csvFileName ($currentSheetIndex/$totalSheets)" -ForegroundColor White
                
                try {
                    $sheetStartTime = Get-Date
                    
                    # Choose conversion method based on mode
                    if ($conversionMode.Mode -eq "Normal") {
                        $csvContent = Convert-SheetToCSV $sheet $sheetName $conversionMode
                    } else {
                        $csvContent = Convert-SheetToCSV-Fast $sheet $sheetName
                    }
                    
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
    
    # Close any open workbooks
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
    
    # Store Excel process IDs before termination attempt
    $preExcelProcesses = @()
    try {
        $preExcelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Select-Object Id, ProcessName
        Write-Host "Excel processes found before cleanup: $($preExcelProcesses.Count)" -ForegroundColor Gray
    } catch {
        # No Excel processes found or error getting process list
    }
    
    # Terminate Excel application properly
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
    
    # Force garbage collection
    Write-Host "Forcing garbage collection..." -ForegroundColor Gray
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    # Wait a moment for processes to terminate naturally
    Start-Sleep -Seconds 2
    
    # Check remaining Excel processes and force terminate if necessary
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
            Start-Sleep -Seconds 3
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