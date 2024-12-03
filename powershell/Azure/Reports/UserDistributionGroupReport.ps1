# Import the ImportExcel module (ensure it's installed first)
Import-Module ImportExcel -ErrorAction Stop

# Connect to Azure AD
Connect-AzureAD

# Initialize a hashtable to store group data
$groupData = @{}

# Get all mail-enabled groups (distribution lists only)
$distributionGroups = Get-AzureADGroup -All $true | Where-Object {
    $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false -and ($_.GroupTypes -notcontains "Unified")
}

# Iterate through each distribution group
foreach ($group in $distributionGroups) {
    $groupName = $group.DisplayName
    $groupEmail = $group.Mail
    $groupId = $group.ObjectId

    # Get members of the group (users and contacts)
    $groupMembers = Get-AzureADGroupMember -ObjectId $groupId | Where-Object {
        $_.ObjectType -in @("User", "Contact")
    }

    # Create a list of members with their names and email addresses
    $groupMembersData = $groupMembers | ForEach-Object {
        $email = if ($_.ObjectType -eq "User") {
            $_.Mail
        }
        elseif ($_.ObjectType -eq "Contact") {
            if ($_.Mail) {
                $_.Mail
            }
            else {
                $_.EmailAddress
            }
        }
        else {
            $null
        }

        # Return a PSCustomObject for each member
        [PSCustomObject]@{
            Name  = if ($_.ObjectType -eq "User") {
                "$($_.GivenName) $($_.Surname)"
            }
            else {
                $_.DisplayName
            }
            Email = $email
            Type  = $_.ObjectType  # Include the type (User or Contact)
        }
    }

    # Add data to the hashtable
    $groupData[$groupName] = [PSCustomObject]@{
        Email   = $groupEmail
        Members = $groupMembersData
    }
}

# Define the Excel file path
$excelFilePath = "C:\Data\DistributionListsWithAllContactsReport.xlsx"  # Replace with your desired file path

# Export data to Excel
foreach ($groupName in $groupData.Keys) {
    $groupEmail = $groupData[$groupName].Email
    # Create a worksheet name with the group name and email in parentheses
    $worksheetName = "$groupName ($groupEmail)"

    # Truncate worksheet name if it's longer than 31 characters
    if ($worksheetName.Length -gt 31) {
        $worksheetName = $worksheetName.Substring(0, 31)
    }

    # Sanitize the worksheet name to remove invalid characters
    $worksheetName = $worksheetName -replace '[\/:*?"<>|]', ''

    # Append each group's data to a new worksheet
    $groupData[$groupName].Members | Export-Excel -Path $excelFilePath -WorksheetName $worksheetName -Append:$true -AutoSize
}

Write-Host "Data export complete. Excel file saved at $excelFilePath"
