# The provided script installs the "AzureADPreview" module, connects to Azure Active Directory, 
# retrieves user data including their first names, last names, emails, and presents the information in a formatted table.
# It also exports the data to a CSV file.

# Install required modules
Install-Module -Name AzureADPreview

# Import modules
Import-Module AzureADPreview

# Connect to Azure AD
Connect-AzureAD

# List users
$users = Get-AzureADUser -All $true | Where-Object { $_.UserType -eq "Member" }

$usersData = $users | ForEach-Object {
    $firstName = $_.GivenName
    $lastName = $_.Surname
    $email = $_.Mail

    [PSCustomObject]@{
        FirstName = $firstName
        LastName = $lastName
        Email = $email
    }
}

# Output the data to a formatted table
$usersData | Format-Table -Property FirstName, LastName, Email -AutoSize

# Export the data to a CSV file
$csvFilePath = "C:\TenantDataExport.csv"  # Replace with your desired file path
$usersData | Export-Csv -Path $csvFilePath -NoTypeInformation

# Get all security groups
$securityGroups = Get-AzureADGroup -All $true | Where-Object { $_.SecurityEnabled -eq $true }

# Iterate through each security group
foreach ($group in $securityGroups) {
    $groupDisplayName = $group.DisplayName
    $groupMembers = Get-AzureADGroupMember -ObjectId $group.ObjectId | Where-Object { $_.ObjectType -eq "User" }

    # Output group members
    Write-Host "Users in group $($group.DisplayName):"
    $groupMembers | ForEach-Object {
        Write-Host "- $($_.UserPrincipalName)"
    }
}
