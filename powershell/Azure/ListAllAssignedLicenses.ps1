# The provided script installs the "AzureADPreview" module, connects to Azure Active Directory, 
# retrieves user data including their licenses, and presents the information in a formatted table.

# Install required modules
Install-Module -Name AzureADPreview

# Import modules
Import-Module AzureADPreview

# Connect to Azure AD
Connect-AzureAD

# List users and their licenses
$users = Get-AzureADUser -All $true | Where-Object { $_.UserType -eq "Member" }

$usersData = $users | ForEach-Object {
    $user = $_.UserPrincipalName
    $licenses = Get-AzureADUserLicenseDetail -ObjectId $_.ObjectId | Select-Object -ExpandProperty SkuPartNumber

    [PSCustomObject]@{
        User = $user
        Licenses = $licenses -join ', '
    }
}

$usersData | Format-Table -Property User, Licenses -AutoSize
