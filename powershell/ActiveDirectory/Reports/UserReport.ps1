# This script exports user information from an Active Directory domain, including the user's name, username, email address, status, and groups they belong to. 
# It retrieves all user objects and iterates over them, collecting group membership details for each user. 
# The script then creates custom objects for each user, adding the relevant properties, and appends them to an array. 
# Finally, it sorts and displays the report on the screen and exports it to a CSV file. 
# This script is useful for generating a comprehensive user report for analysis or documentation purposes.

# Export the Users Name, Username, Email Address, Status, and Groups the user is in.

$Report = @()

# Collect all users
$Users = Get-ADUser -Filter * -Properties Name, GivenName, SurName, SamAccountName, UserPrincipalName, EmailAddress, MemberOf, Enabled -ResultSetSize $Null

# Use ForEach loop, as we need group membership for every account that is collected.
# MemberOf property of User object has the list of groups and is available in DN format.
Foreach($User in $users) {
    $UserGroupCollection = $User.MemberOf

    # This Array will hold Group Names to which the user belongs.
    $UserGroupMembership = @()

    # To get the Group Names from DN format we will again use Foreach loop to query every DN and retrieve the Name property of Group.
    Foreach($UserGroup in $UserGroupCollection){
    $GroupDetails = Get-ADGroup -Identity $UserGroup

    # Here we will add each group Name to UserGroupMembership array
    $UserGroupMembership += $GroupDetails.Name
    }

    # As the UserGroupMembership is array we need to join element with ‘,’ as the seperator
    $Groups = $UserGroupMembership -join ', '

    # Creating custom objects
    $Out = New-Object PSObject
    $Out | Add-Member -MemberType noteproperty -Name Name -Value $User.Name
    $Out | Add-Member -MemberType noteproperty -Name UserName -Value $User.SamAccountName
    $Out | Add-Member -MemberType noteproperty -Name EmailAddress -Value $User.EmailAddress
    $Out | Add-Member -MemberType noteproperty -Name Status -Value $User.Enabled
    $Out | Add-Member -MemberType noteproperty -Name Groups -Value $Groups
    $Report += $Out
}

# Output to screen as well as csv file.
$Report | Sort-Object Name | FT -AutoSize
$Report | Sort-Object Name | Export-Csv -Path 'C:\Data\users.csv' -NoTypeInformation