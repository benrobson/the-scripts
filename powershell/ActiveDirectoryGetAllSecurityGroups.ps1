$ExportPath = 'c:\TEMP\CLIENTSecGroupsDDMMYYYY.csv'

Get-ADGroup -filter {groupCategory -eq 'Security'} | Select Name | Sort-Object Name | Export-Csv -NoType $ExportPath