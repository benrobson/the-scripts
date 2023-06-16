# This script defines a function that retrieves user information from an Active Directory domain based on their login activity.
# It filters and selects enabled users who haven't logged in over the past 90 days within a specified organizational unit. 
# The resulting user information, including their username, name, and last login date, is sorted and exported to a CSV file. 
# This function is useful for identifying inactive user accounts within a specific organizational unit.

# Use -90 for who hasn’t logged in over the past 90 days
# Use 90 for who has logged in over the past 90 days

$OUpath = 'ou=CLIENTOUUSERS,dc=DOMAIN,dc=local'
$ExportPath = 'c:\TEMP\CLIENTUsersDDMMYYYY.csv'

# Getting users who haven't logged in in over 90 days
$Date = (Get-Date).AddDays(-90)
 
# Filtering All enabled users who haven't logged in.
Get-ADUser -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -SearchBase $OUpath -Properties LastLogonDate | select samaccountname, Name, LastLogonDate | Sort-Object LastLogonDate | Export-Csv -NoType $ExportPath