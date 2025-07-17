# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@example.com

# Define the user email and date range
$userEmail = "user@example.com"
$startDate = (Get-Date).AddDays(-90)
$endDate = Get-Date

# Retrieve deleted items from the Recoverable Items folder
$recoverableItems = Get-RecoverableItems -Identity $userEmail -FilterStartTime $startDate -FilterEndTime $endDate -FilterItemType IPM.Note

# Retrieve mailbox audit logs for delete operations
$mailboxAuditLogs = Search-MailboxAuditLog -Identity $userEmail -StartDate $startDate -EndDate $endDate -LogonTypes Owner, Admin, Delegate -ShowDetails |
    Where-Object {$_.Operation -match "Delete"}

# Combine results
$allDeletedEmails = @()
$allDeletedEmails += $recoverableItems
$allDeletedEmails += $mailboxAuditLogs | Select-Object LastAccessed, OperationResult, LogonUserDisplayName, SourceItemSubjectsList, LogonType, Operation, FolderPathName

# Export combined results to a CSV file
$allDeletedEmails | Export-Csv -Path "C:\Data\AllDeletedEmails.csv" -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
