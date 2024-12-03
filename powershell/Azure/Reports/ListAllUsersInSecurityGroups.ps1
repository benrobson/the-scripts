# Connect to Azure AD
Connect-AzureAD

# Toggle for filtering by keyword
$useKeyword = $true  # Set to $false to list all groups without filtering
$keyword = "SEC_"    # Specify your desired keyword for filtering

# Get all users
$users = Get-AzureADUser -All $true | Where-Object { $_.UserType -eq "Member" }

# Initialize an array to store data
$userGroupData = @()

# Iterate through each user to get their group memberships
foreach ($user in $users) {
    $userName = "$($user.GivenName) $($user.Surname)"
    $userEmail = $user.Mail
    $userId = $user.ObjectId

    # Get groups the user is a member of
    $userGroups = Get-AzureADUserMembership -ObjectId $userId | Where-Object {
        $_.ObjectType -eq "Group" -and ($useKeyword -eq $false -or $_.DisplayName -like "*$keyword*")
    }
    $groupNames = $userGroups.DisplayName -join "; "  # Join group names with semicolons

    # Create a custom object for the user if they are in any matching groups
    if ($groupNames) {
        $userGroupData += [PSCustomObject]@{
            Name           = $userName
            Email          = $userEmail
            SecurityGroups = $groupNames
        }
    }
}

# Define the file path for CSV export
$csvFilePath = if ($useKeyword) {
    "C:\Data\TenantUserGroups_Filtered.csv"  # Replace with your desired file path
}
else {
    "C:\Data\TenantUserGroups_All.csv"       # Replace with your desired file path
}

# Export the data to CSV
$userGroupData | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "Data export complete. File saved at $csvFilePath"
