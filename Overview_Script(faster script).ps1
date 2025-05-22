#ignores only the file with extension .lnk -overview
param (
    [string]$ServerName = "AUADL017MF6001P",  # Default server name
    [string[]]$PathsToScan = @("G:\AdaptivaCache"),       # Supports both full drive and subfolder paths
    [string]$OutputDir = "C:\powershell_script\anish_overview" # Default output directory
)

# Generate dynamic output file name with date and time
$DateTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Convert PathsToScan array into a valid filename segment
$SafePaths = ($PathsToScan -join "_") -replace '[\\/:*?"<>|]', '_'

# Construct output file name including ServerName and PathsToScan
$OutputCSV = "$OutputDir\Overview_File_Report_${ServerName}_${SafePaths}_$DateTimeStamp.csv"

$LogFile = "$OutputDir\File_Report_Log_$DateTimeStamp.txt"

# Capture Start Time
$StartTime = Get-Date

# Ensure output folder exists
if (!(Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# Initialize output array
$FinalReportData = @()

# Function to log messages
function Write-Log {
    param ([string]$Message)
    $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Message"
    Add-Content -Path $LogFile -Value $LogEntry
    Write-Host $Message -ForegroundColor Cyan
}

# Function to get folder details
function Get-FolderDetails {
    param ([string]$ScanPath)

    $ReportData = @()
    Write-Log "Scanning path: $ScanPath"

    # Get immediate subfolders inside the scan path
    $SubFolders = Get-ChildItem -Path $ScanPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne ".lnk" }

    # Count files directly in the scanned path (excluding subfolders)
    $FilesInScanPath = Get-ChildItem -Path $ScanPath -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne ".lnk" }
    $FileCount = $FilesInScanPath.Count
    $TotalSize = ($FilesInScanPath | Measure-Object -Property Length -Sum).Sum
    $TotalSizeGB = if ($TotalSize) { [math]::Round($TotalSize / 1GB, 2) } else { 0 }

    # **Always add "Not Applicable" row for files in the scanned path**
    Write-Log "Adding 'Not Applicable' row for: $ScanPath"

    $ReportData += [PSCustomObject]@{
        Server_Name        = $ServerName
        Drive              = $ScanPath
        "Top Level Folder" = "Not Applicable"
        "Data(GB)"         = $TotalSizeGB
        "Number of SubFolders" = 0   # Hardcoded value
        "Number of Files"  = $FileCount
    }

    # Process each subfolder inside the scan path
    foreach ($Folder in $SubFolders) {
        $FolderPath = $Folder.FullName
        Write-Log "Processing subfolder: $FolderPath"

        # Get total size in bytes inside the subfolder, excluding .lnk files
        try {
            $SubfolderFiles = Get-ChildItem -Path $FolderPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne ".lnk" }
            $SubfolderSize = ($SubfolderFiles | Measure-Object -Property Length -Sum).Sum
            $SubfolderSizeGB = if ($SubfolderSize) { [math]::Round($SubfolderSize / 1GB, 2) } else { 0 }
        } catch {
            Write-Log "Error calculating size for folder: $FolderPath"
            $SubfolderSizeGB = "N/A"
        }

        # Get total number of files and subfolders inside this subfolder
        try {
            $SubfolderFileCount = $SubfolderFiles.Count
            $SubfolderSubfolderCount = (Get-ChildItem -Path $FolderPath -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { 
                $_.Extension -ne ".lnk" 
            } | Measure-Object).Count
        } catch {
            Write-Log "Error counting files or subfolders in: $FolderPath"
            $SubfolderFileCount = "N/A"
            $SubfolderSubfolderCount = "N/A"
        }

        # Add subfolder data to report array
        $ReportData += [PSCustomObject]@{
            Server_Name        = $ServerName
            Drive              = $ScanPath
            "Top Level Folder" = $Folder.Name
            "Data(GB)"         = $SubfolderSizeGB
            "Number of SubFolders" = $SubfolderSubfolderCount
            "Number of Files"  = $SubfolderFileCount
        }
    }

    return $ReportData
}

# Loop through each scan path (supports both full drives and subfolders)
foreach ($ScanPath in $PathsToScan) {
    if (Test-Path -Path $ScanPath) {
        Write-Log "Scanning path: $ScanPath"
        $FinalReportData += Get-FolderDetails -ScanPath $ScanPath
    } else {
        Write-Log "Path not found: $ScanPath"
    }
}

# Export data to CSV
if ($FinalReportData.Count -gt 0) {
    $FinalReportData | Export-Csv -NoTypeInformation -Delimiter '|' -Encoding UTF8 -Path $OutputCSV
    Write-Log "Report generated: $OutputCSV"
} else {
    Write-Log "No data collected. Report not generated."
}

# Capture End Time and Calculate Runtime
$EndTime = Get-Date
$Duration = $EndTime - $StartTime

# Final Summary
Write-Host "----------------------------------------------------" -ForegroundColor Green
Write-Host "File discovery completed. Output saved to: $OutputCSV" -ForegroundColor Green
Write-Host "Log file: $LogFile" -ForegroundColor Yellow
Write-Host "Total Runtime: $($Duration.Hours)h $($Duration.Minutes)m $($Duration.Seconds)s" -ForegroundColor Yellow
Write-Host "----------------------------------------------------" -ForegroundColor Green

Write-Log "File discovery completed in $($Duration.Hours)h $($Duration.Minutes)m $($Duration.Seconds)s"
