# Script Name: Reset-SPOFolderAndFilesPermissions.ps1
# Description: This script resets unique permissions on folders and files in a SharePoint document library.
# This version uses PnP.PowerShell for modern authentication and robust execution.

param (
    [Parameter(Mandatory=$false)]
    [string]$SiteURL = "https://SITECORP.sharepoint.com/sites/SITENAME"
)

# --- Variables ---
$ListName = "Documents"
$RelativeFolderPath = "" # e.g. "Shared Documents/Subfolder". Leave blank for library root.

# --- Execution ---

# Connect to SharePoint Online (Modern Auth)
Try {
    Write-Host "-----------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host " AUTHENTICATION REQUIRED" -ForegroundColor Yellow
    Write-Host " Please log in with a Global Administrator or SharePoint Admin account" -ForegroundColor Yellow
    Write-Host " in the browser window that appears." -ForegroundColor Yellow
    Write-Host "-----------------------------------------------------------------------`n" -ForegroundColor Yellow

    Write-Host "Connecting to $SiteURL..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteURL -Interactive -ErrorAction Stop
}
Catch {
    Write-Host "Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Get items from the list/folder
Try {
    Write-Host "Fetching items from '$ListName'..." -ForegroundColor Cyan
    $items = Get-PnPListItem -List $ListName -FolderServerRelativeUrl $RelativeFolderPath -Recursive -PageSize 1000 -Includes "HasUniqueRoleAssignments","FileLeafRef"
}
Catch {
    Write-Host "Failed to retrieve items: $($_.Exception.Message)" -ForegroundColor Red
    return
}

$totalCount = $items.Count
Write-Host "Found $totalCount items to check." -ForegroundColor Yellow

$counter = 0
ForEach ($item in $items) {
    $counter++
    $itemName = $item["FileLeafRef"]

    # Progress update in console
    Write-Progress -Activity "Resetting Permissions" -Status "Processing: $itemName ($counter/$totalCount)" -PercentComplete (($counter / $totalCount) * 100)

    Try {
        # Check if item has unique permissions
        If ($item.HasUniqueRoleAssignments) {
            Write-Host "[$counter/$totalCount] Resetting unique permissions: $($itemName)" -ForegroundColor Green
            $item.ResetRoleInheritance()
            # Invoke-PnPQuery handles 429/503 throttling automatically with retries
            Invoke-PnPQuery -RetryCount 10
        }
        Else {
            # Already inheriting, skip
            # Write-Host "[$counter/$totalCount] Skipping (Already inheriting): $itemName" -ForegroundColor Gray
        }
    }
    Catch {
        Write-Host "[$counter/$totalCount] Error processing $($itemName): $($_.Exception.Message)" -ForegroundColor Red

        # Explicit throttling check as a backup
        If ($_.Exception.Message -like "*429*" -or $_.Exception.Message -like "*503*") {
            Write-Host "Throttling detected. Waiting 10 seconds before continuing..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
        }
    }
}

Write-Host "`nTask Complete! Processed $totalCount items." -ForegroundColor Green
pause
