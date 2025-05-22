#ignores only the file with extension .lnkq
param (
    [string]$ServerName,
    [string[]]$ServerFolderPathsToScan,
    [string]$OutputFolder,
    [int]$MaxFileSizeMB = 500  # Maximum file size before splitting (default: 500MB)
)

if (-not $ServerName) {
    Write-Host "Please provide a ServerName." -ForegroundColor Red
    exit
}

if (-not $ServerFolderPathsToScan) {
    Write-Host "Please provide at least one folder path to scan." -ForegroundColor Red
    exit
}

if (-not $OutputFolder) {
    Write-Host "Please provide an output folder path." -ForegroundColor Red
    exit
}

$StartTime = Get-Date
$DateTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SanitizedFolderPath = ($ServerFolderPathsToScan -join "_") -replace '[\\/:*?"<>|]', '_'
$OutputCSVBase = "File_Discovery_${ServerName}_${SanitizedFolderPath}_$DateTimeStamp"
$OutputCSV = "$OutputFolder\$OutputCSVBase.csv"

if (!(Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -LiteralPath $OutputFolder -Force | Out-Null
}

# Function to add BOM to large CSV files without loading entire content into memory
function Add-BOM {
    param (
        [string]$FilePath
    )

    if (Test-Path -Path $FilePath) {
        $BOM = [System.Text.UTF8Encoding]::new($true).GetPreamble()
        $TempFile = "$FilePath.tmp"

        $SourceStream = [System.IO.File]::OpenRead($FilePath)
        $TargetStream = [System.IO.File]::OpenWrite($TempFile)

        try {
            $TargetStream.Write($BOM, 0, $BOM.Length)

            $BufferSize = 4MB
            $Buffer = New-Object byte[] $BufferSize
            while (($BytesRead = $SourceStream.Read($Buffer, 0, $BufferSize)) -gt 0) {
                $TargetStream.Write($Buffer, 0, $BytesRead)
            }
        }
        finally {
            $SourceStream.Close()
            $TargetStream.Close()
        }

        Remove-Item $FilePath -Force
        Rename-Item $TempFile $FilePath
    }
}

# Function to retrieve file metadata
function Get-FileMetadata {
    param (
        [string]$FilePath
    )

    # Ensure the file exists
    if (!(Test-Path -LiteralPath $FilePath)) {
        Write-Host "Skipping: File not found - $FilePath" -ForegroundColor Red
        return [PSCustomObject]@{
            Owner = ""
            Authors = ""
            LastSavedBy = ""
        }
    }

    try {
        $Item = Get-Item -LiteralPath $FilePath
        $Shell = New-Object -ComObject Shell.Application
        $Folder = $Shell.Namespace($Item.DirectoryName)
        $File = $Folder.ParseName($Item.Name)

        if ($Folder -and $File) {
            $Authors = $Folder.GetDetailsOf($File, 20)  # Author Metadata
            $LastSavedBy = $Folder.GetDetailsOf($File, 13)  # Last Saved By Metadata
        } else {
            $Authors = ""
            $LastSavedBy = ""
        }

        # Get file owner
        $FileSecurity = $Item.GetAccessControl()
        $Owner = $FileSecurity.Owner
    } catch {
        Write-Host "Error retrieving metadata for: $FilePath" -ForegroundColor Red
        $Owner = ""
        $Authors = ""
        $LastSavedBy = ""
    }

    return [PSCustomObject]@{
        Owner = $Owner
        Authors = $Authors
        LastSavedBy = $LastSavedBy
    }
}

# Function to scan files efficiently with streaming and split large CSV files
function Get-Files {
    param (
        [string]$Path,
        [string]$OutputFileBase,
        [string]$OutputFolder,
        [string]$ServerName,
        [int]$MaxFileSizeMB
    )

    Write-Host "Scanning: $Path" -ForegroundColor Cyan

    if (!(Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -LiteralPath $OutputFolder -Force | Out-Null
    }

    $Files = Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne ".lnk" }
    
    $FileCounter = 1
    $CurrentOutputFile = "$OutputFolder\${OutputFileBase}_Part$FileCounter.csv"
    $Stream = [System.IO.StreamWriter]::new($CurrentOutputFile, $true, [System.Text.UTF8Encoding]::new($false))
    $Stream.WriteLine("ServerName|FullName|Date Created|Date Modified|Owner|Authors|Last Saved By|Length|Extension|Attributes|DirectoryName|Name")
    
    try {
        foreach ($File in $Files) {
            $Metadata = Get-FileMetadata -FilePath $File.FullName
            $Data = ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}" -f 
				$ServerName, $File.FullName, $File.CreationTime, $File.LastWriteTime, $Metadata.Owner, 
				$Metadata.Authors, $Metadata.LastSavedBy, $File.Length, $File.Extension, 
				$File.Attributes, $File.DirectoryName, $File.Name) -replace '\u200B', '' -replace '\uFEFF', ''

            
            $Stream.WriteLine($Data)
        }
    } finally {
        $Stream.Close()
        Add-BOM -FilePath $CurrentOutputFile
    }
}

$GrandTotalSizeBytes = 0
foreach ($FolderPath in $ServerFolderPathsToScan) {
    if (Test-Path -LiteralPath $FolderPath) {
        $FolderSize = (Get-ChildItem -LiteralPath $FolderPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $GrandTotalSizeBytes += $FolderSize
        Get-Files -Path $FolderPath -OutputFileBase $OutputCSVBase -OutputFolder $OutputFolder -ServerName $ServerName -MaxFileSizeMB $MaxFileSizeMB
    } else {
        Write-Host "Path not found: $FolderPath" -ForegroundColor Red
    }
}

$GrandTotalSizeGB = [math]::Round($GrandTotalSizeBytes / 1GB, 2)
$EndTime = Get-Date
$Duration = $EndTime - $StartTime

Write-Host "----------------------------------------------------" -ForegroundColor Green
Write-Host "File discovery completed. Output saved to: $OutputFolder\$OutputCSV" -ForegroundColor Green
Write-Host "Total Data Size for All Folders: $GrandTotalSizeGB GB" -ForegroundColor Cyan
Write-Host "Total Runtime: $($Duration.Hours)h $($Duration.Minutes)m $($Duration.Seconds)s" -ForegroundColor Yellow
Write-Host "----------------------------------------------------" -ForegroundColor Green
