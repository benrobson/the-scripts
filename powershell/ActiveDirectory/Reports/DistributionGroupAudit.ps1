# This script retrieves all distribution groups from Active Directory,
# including their name, email address, and the number of members in each group.
# The data is displayed as a formatted table and exported to a CSV file for auditing purposes.

Import-Module ActiveDirectory

# Retrieve all distribution groups
$distributionGroups = Get-ADGroup -Filter { GroupCategory -eq "Distribution" } -Properties Mail, Members

# Create a collection to store group details
$groupInfo = @()

foreach ($group in $distributionGroups) {
    $memberCount = if ($group.Members) { $group.Members.Count } else { 0 }
    
    $groupInfo += [PSCustomObject]@{
        GroupName    = $group.Name
        EmailAddress = $group.Mail
        MemberCount  = $memberCount
    }
}

# Display results in table format
$groupInfo | Format-Table -AutoSize

# Export results to CSV file
$timestamp = Get-Date -Format "yyyyMMdd"
$groupInfo | Export-Csv -Path "C:\Data\DistributionGroups_$timestamp.csv" -NoTypeInformation
