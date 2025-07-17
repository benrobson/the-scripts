# Script Name: Reset-SPOFolderAndFilesPermissions.ps1
# Description: This script resets unique permissions on folders and files in a SharePoint document library.

# Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 
# Function to Reset Permissions of all Sub-folders and Files in a Folder
Function Reset-SPOFolderAndFilesPermissions([Microsoft.SharePoint.Client.Folder]$Folder, [System.Collections.Generic.HashSet[String]]$ProcessedItems)
{
    Try {
        # Get all Sub Folders and Files
        $Ctx.Load($Folder.Folders)
        $Ctx.Load($Folder.Files)
        $Ctx.ExecuteQuery()

        # Iterate through each sub-folder and file of the folder
        Foreach ($Item in $Folder.Folders + $Folder.Files | Where {$_.Name -ne "Forms" -and $_.Name -ne "Document"})
        {
            # Check if item has already been processed
            If ($ProcessedItems.Contains($Item.ServerRelativeUrl)) {
                Write-host "Item already processed. Skipping:"$Item.ServerRelativeUrl
                Continue
            }

            Write-host "Processing Item:"$Item.ServerRelativeUrl

            # Get the "Has Unique Permissions" Property
            $Item.ListItemAllFields.Retrieve("HasUniqueRoleAssignments")
            $Ctx.ExecuteQuery()
   
            If($Item.ListItemAllFields.HasUniqueRoleAssignments -eq $True)
            {
                # Reset Item Permissions
                $Item.ListItemAllFields.ResetRoleInheritance()
                $Ctx.ExecuteQuery()
                Write-host -f Green "`tItem's Unique Permissions are Removed!"
            }

            # Add item to the processed items set
            $ProcessedItems.Add($Item.ServerRelativeUrl)

            # If it's a folder, recurse into it
            If ($Item -is [Microsoft.SharePoint.Client.Folder]) {
                Reset-SPOFolderAndFilesPermissions $Item $ProcessedItems
            }

            # Introduce a delay to avoid hitting rate limits
            Start-Sleep -Seconds $SleepTime
        }
    }
    Catch {
        write-host -f Red "Error Resetting Item Permissions!" $_.Exception.Message

        # If it's a rate limit error, wait and then retry
        if ($_.Exception.Message -like '*429*') {
            Write-Host "Encountered 429 error. Retrying in 5 seconds..."
            Start-Sleep -Seconds 5
            Reset-SPOFolderAndFilesPermissions $Folder $ProcessedItems
        }
    }
}

# Variables
$SiteURL = "https://SITECORP.sharepoint.com/sites/SITENAME"
$ListName = "Documents"
$SleepTime = 1

# Get Credentials to connect
$Username = "SITEADMIN@SITETENANT.onmicrosoft.com"
$SecurePassword = ConvertTo-SecureString -String "APPPASSWORD" -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)


# Setup the context
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)

# Get the Library
$List = $Ctx.web.Lists.GetByTitle($ListName)
$Ctx.Load($List.RootFolder)
$Ctx.ExecuteQuery()

# Initialize the HashSet to store processed items
$ProcessedItems = New-Object System.Collections.Generic.HashSet[String]

# Call the function to reset permissions of all folders and files of the document library
Reset-SPOFolderAndFilesPermissions $List.RootFolder $ProcessedItems

pause
