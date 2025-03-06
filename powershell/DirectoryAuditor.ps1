<#
.SYNOPSIS
This script logs all folders and files recursively from a specified UNC path and outputs the results to a text file.

.DESCRIPTION
The script scans a UNC path, logs each file and folder, and handles access issues by noting them in the log. It also calculates the total number of files, folders, and their size, and records the scan duration.

.PARAMETER uncPath
The UNC path to scan. Example: "\\server\share"

.PARAMETER outputDir
The directory where the output file will be saved. Example: "C:\Data"

.EXAMPLE
.\DirectoryAuditor.ps1 -uncPath "\\server\share" -outputDir "C:\Data"
#>

param (
    [string]$uncPath = "\\server\share",
    [string]$outputDir = "C:\Data"
)

# Extract the server hostname from the UNC path
$serverHostname = $uncPath -replace '^\\\\([^\\]+).*', '$1'

# Define the output file path with the server hostname and UNC path
$outputPath = Join-Path -Path $outputDir -ChildPath "$($uncPath.Replace('\', '-').TrimStart('-'))-$(Get-Date -Format 'yyyyMMdd').txt"

# Initialize counters
$totalFiles = 0
$totalFolders = 0
$totalSize = 0

# Record the start time
$startTime = Get-Date

# Function to log folders and files recursively
function Log-FoldersAndFiles {
    param (
        [string]$path,
        [ref]$totalFiles,
        [ref]$totalFolders,
        [ref]$totalSize
    )

    try {
        # Get all items in the current directory
        $items = Get-ChildItem -Path $path -ErrorAction Stop

        foreach ($item in $items) {
            # Output the full name of the item
            $item.FullName | Out-File -FilePath $outputPath -Append

            # Write progress to host
            Write-Host "Scanning: $($item.FullName)"

            # If the item is a directory, recursively log its contents
            if ($item.PSIsContainer) {
                $totalFolders.Value++
                Log-FoldersAndFiles -path $item.FullName -totalFiles $totalFiles -totalFolders $totalFolders -totalSize $totalSize
            }
            else {
                $totalFiles.Value++
                $totalSize.Value += $item.Length
            }
        }
    }
    catch {
        # If access is denied, log the path and specify that contents could not be accessed
        "Access Denied: $path" | Out-File -FilePath $outputPath -Append
        Write-Host "Access Denied: $path"
    }
}

# Create the output directory if it doesn't exist
if (-Not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Log folders and files starting from the UNC path
Log-FoldersAndFiles -path $uncPath -totalFiles ([ref]$totalFiles) -totalFolders ([ref]$totalFolders) -totalSize ([ref]$totalSize)

# Record the end time
$endTime = Get-Date

# Calculate the total scan time in minutes
$scanDuration = ($endTime - $startTime).TotalMinutes

# Convert total size from bytes to gigabytes
$totalSizeGB = [math]::round($totalSize / 1GB, 2)

# Append summary information to the output file
"Total Files: $totalFiles" | Out-File -FilePath $outputPath -Append
"Total Folders: $totalFolders" | Out-File -FilePath $outputPath -Append
"Total Size: $totalSizeGB GB" | Out-File -FilePath $outputPath -Append
"Total Scan Time: $scanDuration minutes" | Out-File -FilePath $outputPath -Append

Write-Output "Logging completed. Output saved to $outputPath"