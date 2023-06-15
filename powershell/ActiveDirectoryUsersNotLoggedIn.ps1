# Use -90 for who hasn’t logged in over the past 90 days
# Use 90 for who has logged in over the past 90 days

$OUpath = 'ou=CLIENTOUUSERS,dc=DOMAIN,dc=local'
$ExportPath = 'c:\TEMP\CLIENTUsersDDMMYYYY.csv'

# Getting users who haven't logged in in over 90 days
$Date = (Get-Date).AddDays(-90)
 
# Filtering All enabled users who haven't logged in.
Get-ADUser -Filter {((Enabled -eq $true) -and (LastLogonDate -lt $date))} -SearchBase $OUpath -Properties LastLogonDate | select samaccountname, Name, LastLogonDate | Sort-Object LastLogonDate | Export-Csv -NoType $ExportPath