# This script exports the names of security groups in an Active Directory domain to a CSV file.
# It begins by setting the export path for the CSV file. 
# Then, it uses the Get-ADGroup cmdlet with a filter to retrieve only security groups. 
# The script selects the group names and sorts them alphabetically. 
# Finally, it exports the results to the specified CSV file. 
# This script is useful for obtaining a list of security groups in the domain for various administrative purposes.

$ExportPath = 'c:\Data\CLIENTSecGroupsDDMMYYYY.csv'

Get-ADGroup -filter {groupCategory -eq 'Security'} | Select Name | Sort-Object Name | Export-Csv -NoType $ExportPath