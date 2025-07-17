Import-Module ActiveDirectory

<#
.DESCRIPTION
This script checks all users in a specified Organizational Unit (OU) in Active Directory to verify if their SMTP addresses include any of the specified domains.

.PARAMETER $ou
The distinguished name (DN) of the Organizational Unit (OU) to search.

.PARAMETER $acceptedDomains
An array of domains to verify against user email addresses.

.OUTPUTS
- A table output of users without matching SMTP addresses.
- An optional CSV file containing the same information.

.NOTES
Ensure the Active Directory module is installed and the script is run with sufficient permissions to query AD users.

.EXAMPLE
# Example usage:
# Define the OU and domains, then run the script
$ou = "OU=YourOU,DC=YourDomain,DC=com"
$acceptedDomains = @("example.com", "example.org")
Run the script to check SMTP addresses and export results to CSV.
#>

# Define the OU to search
$ou = "OU=YourOU,DC=YourDomain,DC=com"

# Define the accepted domains
$acceptedDomains = @("example.com", "example.org")

# Function to check if a user has an accepted SMTP address
function Has-AcceptedSMTP {
    param (
        [string[]]$EmailAddresses,
        [string[]]$Domains
    )
    foreach ($email in $EmailAddresses) {
        foreach ($domain in $Domains) {
            if ($email -match "@${domain}$") {
                return $true
            }
        }
    }
    return $false
}

# Query all users in the specified OU
$users = Get-ADUser -Filter * -SearchBase $ou -Properties EmailAddress, ProxyAddresses

# Check each user
$results = @()
foreach ($user in $users) {
    $emailAddresses = @()

    # Collect email addresses
    if ($user.EmailAddress) {
        $emailAddresses += $user.EmailAddress
    }
    if ($user.ProxyAddresses) {
        $emailAddresses += $user.ProxyAddresses -replace "SMTP:|smtp:", ""
    }

    # Check if all required domains are covered
    $missingDomains = @()
    foreach ($domain in $acceptedDomains) {
        if (-not ($emailAddresses -match "@${domain}$")) {
            $missingDomains += $domain
        }
    }

    if ($missingDomains.Count -gt 0) {
        $results += [PSCustomObject]@{
            UserName       = $user.SamAccountName
            DisplayName    = $user.Name
            EmailAddresses = $emailAddresses -join ", "
            MissingDomains = $missingDomains -join ", "
        }
    }
}

# Output the results
if ($results.Count -eq 0) {
    Write-Output "All users in the OU have SMTP addresses matching the accepted domains."
}
else {
    Write-Output "The following users do not have SMTP addresses matching all the accepted domains:"
    $results | Format-Table -AutoSize

    # Optionally export to a CSV
    $exportPath = "C:\Export\NonMatchingSMTPUsers.csv"
    $results | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Output "Results have been exported to $exportPath"
}
