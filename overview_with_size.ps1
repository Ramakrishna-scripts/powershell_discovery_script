param (
    [string]$netPath,
    [string]$csvPath,  # Argument for input CSV file
    [string]$outputDir # Directory for output files
)

# Check if CSV path is provided and valid
if (-not $csvPath -or -not (Test-Path $csvPath)) {
    Write-Host "Error: Please provide a valid CSV file path." -ForegroundColor Red
    exit
}

# Check if output directory is provided and valid
if (-not $outputDir -or -not (Test-Path $outputDir)) {
    Write-Host "Error: Please provide a valid output directory path." -ForegroundColor Red
    exit
}

# Start timing the script execution
$startTime = Get-Date

# Generate timestamp in 'yyyyMMdd_HHmmss' format
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Generate output file paths in the specified output directory
$outputCsvPath = Join-Path -Path $outputDir -ChildPath ("$(Split-Path -Leaf $csvPath -Resolve)".Replace('.csv', "_${netPath}_${timestamp}_output.csv"))
$errorLogPath = Join-Path -Path $outputDir -ChildPath ("$(Split-Path -Leaf $csvPath -Resolve)".Replace('.csv', "_${netPath}_${timestamp}_error.log"))
$fullLogPath = Join-Path -Path $outputDir -ChildPath ("$(Split-Path -Leaf $csvPath -Resolve)".Replace('.csv', "_${netPath}_${timestamp}_full.log"))

# Start transcript for full console output
Start-Transcript -Path $fullLogPath -Append

# Import CSV and initialize an empty array to collect results
$data = Import-Csv -Path $csvPath
$results = @()

# Loop through each entry
$data | ForEach-Object {
    # $vserver = $_.'vserver'
    $path = $_.'path'

    # Append '_share' to the first word in the path and replace '/' with '\'
    $modifiedPath = ($path -replace '^\/([^\/]+)', '$1_share') -replace '/', '\'

    # Construct the UNC path
    $uncPath = "\\$netPath\$modifiedPath"

    # Run the Get-ChildItem command and count files
    Write-Host "Processing path: $uncPath"
    try {
        $fileSize ="{0:N2} GB" -f ((Get-ChildItem -LiteralPath $uncPath -Recurse | Measure-Object -Property Length -Sum).Sum / 1GB)
    } catch {
        $errorDetails = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error accessing: $uncPath - $($_.Exception.Message)"
        Write-Host $errorDetails -ForegroundColor Red
        Add-Content -Path $errorLogPath -Value $errorDetails
        $fileSize = "N/A"  # Default to zero if path is inaccessible
    }

    # Add results to the array (correct column order)
    $entry = $_ | Select-Object *, @{Name='UNC_Path'; Expression={$uncPath}}, @{Name='Size'; Expression={$fileSize}}
    $results += $entry
}

# Export the updated data to a new CSV
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation

# Stop timing the script execution
$endTime = Get-Date
$timeTaken = $endTime - $startTime

Write-Host "File created successfully at: $outputCsvPath"
Write-Host "Errors (if any) logged at: $errorLogPath"
Write-Host "Full log saved at: $fullLogPath"
Write-Host "Total execution time: $($timeTaken.Hours)h $($timeTaken.Minutes)m $($timeTaken.Seconds)s"
  
# Stop transcript to complete the full log file
Stop-Transcript
