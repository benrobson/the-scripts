# Retrieves a list of users and contacts from Active Directory along with their distribution group memberships. 
# It distinguishes between users and contacts, includes their first name, last name, email address, and indicates whether they are a user or a contact.

Import-Module ActiveDirectory

$users = Get-ADObject -LDAPFilter "(&(objectClass=user)(objectCategory=person))" -Properties GivenName, sn, Mail, MemberOf
$contacts = Get-ADObject -LDAPFilter "(objectClass=contact)" -Properties GivenName, sn, Mail, MemberOf
$allUsers = $users + $contacts
$userMemberships = @()

foreach ($user in $allUsers) {
    $userGroups = $user.MemberOf | ForEach-Object {
        $group = Get-ADGroup $_ -Properties GroupCategory
        if ($group.GroupCategory -eq 'Distribution') {
            $group.Name
        }
    }

    if ($userGroups) {
        $entryType = if ($user.ObjectClass -eq 'user') {
            'User'
        } else {
            'Contact'
        }

        $lastName = if ($entryType -eq 'User') {
            $user.sn
        } else {
            $user.'sn'
        }

        $userMemberships += [PSCustomObject]@{
            EntryType = $entryType
            FirstName = $user.GivenName
            LastName = $lastName
            EmailAddress = $user.Mail
            Groups = $userGroups -join ', '
        }
    }
}

$userMemberships | Format-Table -AutoSize

$userMemberships | Export-Csv -Path "c:\CLIENTDistUsersAuditDDMMYYYY.csv" -NoTypeInformation