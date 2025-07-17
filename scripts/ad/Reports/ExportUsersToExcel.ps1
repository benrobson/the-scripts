# Import the Active Directory module
Import-Module ActiveDirectory

# Set the output file path
$OutputFile = "C:\Users\Public\UserExport.xlsx"

# Set the starting OU
$StartingOU = "OU=Users,DC=example,DC=com" # ***REPLACE WITH YOUR ACTUAL OU PATH***

# Hashtable to store users categorized by OU
$UsersByOU = @{}

try {
    $ADUsers = Get-ADUser -Filter * -SearchBase $StartingOU -Properties GivenName, Surname, SamAccountName, DistinguishedName, MemberOf -ResultSetSize $null | Where-Object { $_.Enabled -eq $true }
}
catch {
    Write-Error "Error retrieving users from Active Directory: $($_.Exception.Message)"
    return
}

foreach ($User in $ADUsers) {
    # Extract relative OU
    $OU = ($User.DistinguishedName -replace "^CN=[^,]+,", "") -replace [regex]::Escape($StartingOU) + ",", "" -replace "^OU=", "" -replace ",DC=.*", ""

    # Extract group names (removing the distinguished name parts)
    $Groups = @()
    if ($User.MemberOf) {
        foreach ($GroupDN in $User.MemberOf) {
            try {
                $Group = Get-ADGroup $GroupDN -Properties Name | Select-Object -ExpandProperty Name
                $Groups += $Group
            }
            catch {
                Write-Warning "Could not retrieve group information for DN: $GroupDN for user $($User.SamAccountName)"
            }
        }
    }
    $GroupString = ($Groups -join "; ") # Join groups with semicolon

    # Create user data object
    $UserData = [PSCustomObject]@{
        FirstName = $User.GivenName
        LastName  = $User.Surname
        Username  = $User.SamAccountName
        Groups    = $GroupString # Add groups to output
    }

    # Add user to the appropriate OU category in the hashtable
    if (-not $UsersByOU.ContainsKey($OU)) {
        $UsersByOU[$OU] = @()
    }
    $UsersByOU[$OU] += $UserData
}

# Export to Excel with separate worksheets for each OU
if ($UsersByOU.Count -gt 0) {
    try {
        if (!(Get-Module -ListAvailable ImportExcel)) {
            Write-Warning "ImportExcel module is not installed. Installing..."
            Install-Module ImportExcel -Force
        }

        foreach ($OU in $UsersByOU.Keys) {
            $UsersByOU[$OU] | Export-Excel -Path $OutputFile -AutoSize -WorksheetName $OU -FreezePane "2,2" -BoldTopRow
        }
        Write-Host "User data exported to: $OutputFile"
    }
    catch {
        Write-Error "Error exporting to Excel: $($_.Exception.Message)"
    }
}
else {
    Write-Warning "No users found in Active Directory under the specified OU."
}
