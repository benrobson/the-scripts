# Count number of users
(Get-ADUser -Filter *).Count

# Count number of groups
(Get-ADGroup -Filter *).Count

# Count number of computers
(Get-ADComputer -Filter *).Count