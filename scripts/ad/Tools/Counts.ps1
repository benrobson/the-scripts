# This script provides a summary of the counts for users, groups, and computers in an Active Directory domain.
# It uses PowerShell cmdlets to retrieve and count the number of objects in each category.
# This information can be useful for administrators to get an overview of the Active Directory domain's size and distribution of objects.

# Count number of users
(Get-ADUser -Filter *).Count

# Count number of groups
(Get-ADGroup -Filter *).Count

# Count number of computers
(Get-ADComputer -Filter *).Count