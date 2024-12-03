# This PowerShell script retrieves information about user accounts in an Active Directory domain.
# It separates users into two categories based on their login activity over the past 90 days: those who haven't logged in and those who have. 
# The script specifies the OU path and export path for the resulting data. 
# It uses the Get-ADComputer cmdlet to filter and retrieve user information, including samaccountname, Name, and LastLogonDate properties. 
# The data is sorted by LastLogonDate and exported to a CSV file. This script helps administrators identify inactive users and manage their accounts effectively.

# Use -90 for who hasn’t logged in over the past 90 days
# Use 90 for who has logged in over the past 90 days

$OUpath = 'ou=CLIENTOUUSERS,dc=DOMAIN,dc=local'
$ExportPath = 'c:\Data\CLIENTComputersDDMMYYYY.csv'

# Getting users who haven't logged in in over 90 days
$Date = (Get-Date).AddDays(-90)
 
# Filtering All enabled users who haven't logged in.
Get-ADComputer -Filter {LastLogonTimeStamp -lt $Date} -ResultPageSize 2000 -resultSetSize $null -Properties LastLogonDate | select samaccountname, Name, LastLogonDate | Sort-Object LastLogonDate | Export-Csv -NoType $ExportPath