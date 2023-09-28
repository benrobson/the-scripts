#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 
#Function to Reset Permissions of all Sub-folders in a Folder
Function Reset-SPOSubFolderPermissions([Microsoft.SharePoint.Client.Folder]$Folder)
{
    Try {
        #Get all Sub Folders
        $Ctx.Load($Folder.Folders)
        $Ctx.ExecuteQuery()
  
        #Iterate through each sub-folder of the folder
        Foreach ($Folder in $Folder.Folders | Where {$_.Name -ne "Forms" -and $_.Name -ne "Document"})
        {
            Write-host "Processing Folder:"$Folder.ServerRelativeUrl
 
            #Get the "Has Unique Permissions" Property
            $Folder.ListItemAllFields.Retrieve("HasUniqueRoleAssignments")
            $Ctx.ExecuteQuery()
   
            If($Folder.ListItemAllFields.HasUniqueRoleAssignments -eq $True)
            {
                #Reset Folder Permissions
                $Folder.ListItemAllFields.ResetRoleInheritance()
                $Ctx.ExecuteQuery()
                Write-host -f Green "`tFolder's Unique Permissions are Removed!"
            }
 
            #Call the function recursively
            Reset-SPOSubFolderPermissions $Folder
        }
    }
    Catch {
        write-host -f Red "Error Resetting Folder Permissions!" $_.Exception.Message
    }
}
  
#Variables
$SiteURL = "https://crescent.sharepoint.com/sites/marketing"
$ListName = "Documents"
  
#Get Credentials to connect
$Cred= Get-Credential
  
#Setup the context
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
  
#Get the Library
$List = $Ctx.web.Lists.GetByTitle($ListName)
$Ctx.Load($List.RootFolder)
$Ctx.ExecuteQuery()
  
#call the function to reset permissions of all folders of the document library
Reset-SPOSubFolderPermissions $List.RootFolder