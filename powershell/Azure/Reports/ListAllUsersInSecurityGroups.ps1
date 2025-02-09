### PowerShell Script to Export Azure AD User Group Memberships
# This script connects to Azure AD, retrieves all users, and exports their security group memberships to a CSV file.
# The script allows filtering by a keyword in the group name.

# Connect to Azure AD
Connect-AzureAD  # Ensure you have the AzureAD module installed and are authenticated

# Toggle for filtering by keyword
$useKeyword = $true  # Set to $false to list all groups without filtering
$keyword = "SEC_"    # Specify your desired keyword for filtering (e.g., groups containing 'SEC_')

# Get all users from Azure AD
$users = Get-AzureADUser -All $true | Where-Object { $_.UserType -eq "Member" }

# Initialize an array to store user and group data
$userGroupData = @()

# Iterate through each user to retrieve their group memberships
foreach ($user in $users) {
    $userName = "$($user.GivenName) $($user.Surname)"  # Construct full name
    $userEmail = $user.Mail  # Get user email address
    $userId = $user.ObjectId  # Get user unique identifier

    # Retrieve groups the user is a member of, filtering based on keyword if enabled
    $userGroups = Get-AzureADUserMembership -ObjectId $userId | Where-Object {
        $_.ObjectType -eq "Group" -and ($useKeyword -eq $false -or $_.DisplayName -like "*$keyword*")
    }
    
    # Extract and join group names into a single string
    $groupNames = $userGroups.DisplayName -join "; "  

    # Store user and their security groups in an object if they belong to any matching groups
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
    "C:\Data\TenantUserGroups_Filtered.csv"  # File path for filtered results
}
else {
    "C:\Data\TenantUserGroups_All.csv"       # File path for all results
}

# Export the data to a CSV file without type information for easy readability
$userGroupData | Export-Csv -Path $csvFilePath -NoTypeInformation

# Output message to indicate completion
Write-Host "Data export complete. File saved at $csvFilePath"
