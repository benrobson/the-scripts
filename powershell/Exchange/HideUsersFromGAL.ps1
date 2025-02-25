# Define the Active Directory group to process
$groupDistinguishedName = "CN=Offboarded,OU=Security Groups,OU=Groups,DC=example,DC=com"

try {
    # Retrieve members of the specified Active Directory group
    $groupMembers = Get-ADGroupMember -Identity $groupDistinguishedName | Select-Object SamAccountName

    # Check if there are any members in the group
    if ($groupMembers) {
        foreach ($member in $groupMembers) {
            try {
                # Retrieve the specific attribute for the user
                $cloudExtensionAttribute = Get-ADUser -Identity $member.SamAccountName -Properties msDS-cloudExtensionAttribute1 | Select-Object -ExpandProperty msDS-cloudExtensionAttribute1

                # Check if the attribute does not already contain the value "HideFromGAL"
                if (-not ($cloudExtensionAttribute -match "HideFromGAL")) {
                    # Set the attribute to hide the user from the GAL
                    Set-ADUser -Identity $member.SamAccountName -Replace @{'msDS-cloudExtensionAttribute1' = "HideFromGAL" } -ErrorAction Stop
                    Write-Host "Hidden $($member.SamAccountName) from GAL"
                }
                else {
                    Write-Host "$($member.SamAccountName) already hidden"
                }
            }
            catch {
                Write-Error "Error processing user $($member.SamAccountName): $_"
            }
        }
    }
    else {
        Write-Warning "No users found in the group '$groupDistinguishedName'."
    }
}
catch {
    Write-Error "Error getting members of group '$groupDistinguishedName': $_"
}
