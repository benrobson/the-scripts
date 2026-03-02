#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Online Permissions Tool (GUI)
    - Detects unique permissions and counts items (Dry Run)
    - Detects site-level external sharing settings
    - Optionally resets unique permissions (inheritance)
    - Compatible with PowerShell 5.1 and uses CSOM (Username + App Password)

.DESCRIPTION
    This script provides a GUI to manage SharePoint Online permissions. It can perform a 'Dry Run' to
    audit unique permissions and item counts, or a 'Reset' run to remove unique permissions from
    folders and files in a specified library.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

# ---- CSOM DLLs ----
# Default path for SharePoint Client components
$csomBase = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI"

if (-not (Test-Path (Join-Path $csomBase "Microsoft.SharePoint.Client.dll"))) {
    [System.Windows.Forms.MessageBox]::Show("SharePoint Client DLLs not found at $csomBase. Please ensure SharePoint Client Components are installed.", "Error", "OK", "Error") | Out-Null
    return
}

Add-Type -Path (Join-Path $csomBase "Microsoft.SharePoint.Client.dll")
Add-Type -Path (Join-Path $csomBase "Microsoft.SharePoint.Client.Runtime.dll")

# ---------------- GUI Initialization ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SPO Permissions Manager (CSOM)"
$form.Size = New-Object System.Drawing.Size(980, 800)
$form.StartPosition = "CenterScreen"

$font = New-Object System.Drawing.Font("Segoe UI", 9)

function New-Label($text, $x, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.AutoSize = $true
    $l.Font = $font
    $l
}

function New-TextBox($x, $y, $w, $text = "") {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y)
    $t.Size = New-Object System.Drawing.Size($w, 22)
    $t.Font = $font
    $t.Text = $text
    $t
}

function UiLog([string]$msg) {
    $time = (Get-Date).ToString("HH:mm:ss")
    $txtLog.AppendText("[$time] $msg`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
}

# Inputs
$form.Controls.Add((New-Label "Site URL" 10 10))
$txtSite = New-TextBox 10 30 945 ""
$form.Controls.Add($txtSite)

$form.Controls.Add((New-Label "Library Title" 10 60))
$txtLib = New-TextBox 10 80 240 "Documents"
$form.Controls.Add($txtLib)

$form.Controls.Add((New-Label "Username (UPN)" 270 60))
$txtUser = New-TextBox 270 80 330 ""
$form.Controls.Add($txtUser)

$form.Controls.Add((New-Label "App Password" 610 60))
$txtPass = New-TextBox 610 80 345 ""
$txtPass.UseSystemPasswordChar = $true
$form.Controls.Add($txtPass)

$form.Controls.Add((New-Label "Tenant Admin URL (Optional, for External Sharing check)" 10 110))
$txtAdmin = New-TextBox 10 130 945 ""
$form.Controls.Add($txtAdmin)

$form.Controls.Add((New-Label "Output Folder" 10 160))
$txtOut = New-TextBox 10 180 780 (Join-Path $PWD "SPO-Permissions-Output")
$form.Controls.Add($txtOut)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(800, 178)
$btnBrowse.Size = New-Object System.Drawing.Size(155, 25)
$form.Controls.Add($btnBrowse)

# Mode Selection
$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "Execution Mode"
$grpMode.Location = New-Object System.Drawing.Point(10, 210)
$grpMode.Size = New-Object System.Drawing.Size(300, 60)
$form.Controls.Add($grpMode)

$rbDryRun = New-Object System.Windows.Forms.RadioButton
$rbDryRun.Text = "Dry Run (Scan Only)"
$rbDryRun.Location = New-Object System.Drawing.Point(10, 25)
$rbDryRun.Checked = $true
$rbDryRun.AutoSize = $true
$grpMode.Controls.Add($rbDryRun)

$rbReset = New-Object System.Windows.Forms.RadioButton
$rbReset.Text = "Reset Permissions"
$rbReset.Location = New-Object System.Drawing.Point(150, 25)
$rbReset.AutoSize = $true
$grpMode.Controls.Add($rbReset)

$form.Controls.Add((New-Label "Sleep (sec) between items" 320 210))
$numSleep = New-Object System.Windows.Forms.NumericUpDown
$numSleep.Location = New-Object System.Drawing.Point(320, 235)
$numSleep.Size = New-Object System.Drawing.Size(100, 22)
$numSleep.Minimum = 0
$numSleep.Maximum = 10
$numSleep.Value = 1
$form.Controls.Add($numSleep)

# Buttons
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Process"
$btnStart.Location = New-Object System.Drawing.Point(10, 280)
$btnStart.Size = New-Object System.Drawing.Size(155, 32)
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(175, 280)
$btnCancel.Size = New-Object System.Drawing.Size(155, 32)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(340, 286)
$progress.Size = New-Object System.Drawing.Size(615, 20)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Idle"
$lblStatus.Location = New-Object System.Drawing.Point(340, 308)
$lblStatus.Size = New-Object System.Drawing.Size(615, 20)
$form.Controls.Add($lblStatus)

# Log
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 340)
$txtLog.Size = New-Object System.Drawing.Size(945, 400)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

$folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog

$btnBrowse.Add_Click({
    if ($folderDlg.ShowDialog() -eq "OK") {
        $txtOut.Text = $folderDlg.SelectedPath
    }
})

function SetUiRunning([bool]$running) {
    $btnStart.Enabled = -not $running
    $btnCancel.Enabled = $running
    $txtSite.Enabled = -not $running
    $txtLib.Enabled = -not $running
    $txtUser.Enabled = -not $running
    $txtPass.Enabled = -not $running
    $txtAdmin.Enabled = -not $running
    $txtOut.Enabled  = -not $running
    $btnBrowse.Enabled = -not $running
    $numSleep.Enabled = -not $running
    $grpMode.Enabled = -not $running
}

# ---------------- Background Worker ----------------
$worker = New-Object System.ComponentModel.BackgroundWorker
$worker.WorkerReportsProgress = $true
$worker.WorkerSupportsCancellation = $true

# Register-ObjectEvent is used for compatibility with StrictMode and PS 5.1
Get-EventSubscriber | Where-Object { $_.SourceObject -eq $worker } | Unregister-Event -Force -ErrorAction SilentlyContinue

Register-ObjectEvent -InputObject $worker -EventName DoWork -Action {
    $a = $EventArgs.Argument
    $siteUrl  = $a.SiteUrl
    $libTitle = $a.Library
    $user     = $a.Username
    $pass     = $a.Password
    $adminUrl = $a.AdminUrl
    $out      = $a.Output
    $sleepSec = [int]$a.SleepSec
    $isDryRun = [bool]$a.IsDryRun

    New-Item -ItemType Directory -Force -Path $out | Out-Null
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $flagCsv   = Join-Path $out "permissions-flagged-items-$ts.csv"
    $countCsv  = Join-Path $out "folder-counts-$ts.csv"
    $summaryJs = Join-Path $out "summary-$ts.json"

    $worker.ReportProgress(0, @{type="log"; msg="[INFO] Connecting to SharePoint..."})

    $sharingCapability = "Unknown"
    $secure = ConvertTo-SecureString -String $pass -AsPlainText -Force

    try {
        # Check External Sharing Capability if Admin URL is provided
        if (-not [string]::IsNullOrWhiteSpace($adminUrl)) {
            $worker.ReportProgress(0, @{type="log"; msg="[INFO] Connecting to Tenant Admin to check Sharing Capability..."})
            try {
                $adminCtx = New-Object Microsoft.SharePoint.Client.ClientContext($adminUrl)
                $adminCtx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)
                $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($adminCtx)
                $siteProperties = $tenant.GetSitePropertiesByUrl($siteUrl, $true)
                $adminCtx.Load($siteProperties)
                $adminCtx.ExecuteQuery()
                $sharingCapability = $siteProperties.SharingCapability.ToString()
                $worker.ReportProgress(0, @{type="log"; msg=("[INFO] Site Sharing Capability: " + $sharingCapability)})
            } catch {
                $worker.ReportProgress(0, @{type="log"; msg=("[WARN] Failed to check Sharing Capability: " + $_.Exception.Message)})
                # Continue anyway, as this is an optional check
            }
        }

        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
        $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)

        $web = $ctx.Web
        $ctx.Load($web, "Title", "Url")
        $list = $web.Lists.GetByTitle($libTitle)
        $ctx.Load($list, "Title", "ItemCount")
        $ctx.Load($list.RootFolder, "ServerRelativeUrl")
        $ctx.ExecuteQuery()

        $rootUrl = $list.RootFolder.ServerRelativeUrl.TrimEnd("/")
        $worker.ReportProgress(2, @{type="log"; msg=("[INFO] Connected to {0}" -f $web.Title)})
        $worker.ReportProgress(2, @{type="log"; msg=("[INFO] Library '{0}' has approximately {1} items." -f $libTitle, $list.ItemCount)})
    } catch {
        $worker.ReportProgress(0, @{type="log"; msg=("[ERROR] Connection failed: " + $_.Exception.Message)})
        throw $_.Exception
    }

    $worker.ReportProgress(5, @{type="status"; msg="Retrieving items..."})

    $allItems = New-Object System.Collections.Generic.List[object]
    $position = $null
    do {
        if ($worker.CancellationPending) { $EventArgs.Cancel = $true; return }
        $qry = New-Object Microsoft.SharePoint.Client.CamlQuery
        $qry.ViewXml = "<View Scope='RecursiveAll'><RowLimit>2000</RowLimit></View>"
        $qry.ListItemCollectionPosition = $position
        $items = $list.GetItems($qry)
        # We also try to load if it's shared externally.
        # Note: CSOM doesn't have a simple 'IsShared' property on ListItem, but we can look for guest users if needed.
        # For now, we'll stick to unique permissions as the primary flag.
        $ctx.Load($items, "Include(FileRef, FileDirRef, FileLeafRef, HasUniqueRoleAssignments, FileSystemObjectType)")
        $ctx.ExecuteQuery()
        $position = $items.ListItemCollectionPosition
        foreach ($it in $items) { $allItems.Add($it) }
        $worker.ReportProgress(5, @{type="log"; msg=("[INFO] Retrieved {0} items..." -f $allItems.Count)})
    } while ($position -ne $null)

    $total = $allItems.Count
    $folderCounts = @{}
    $flagged = New-Object System.Collections.Generic.List[object]

    $worker.ReportProgress(10, @{type="status"; msg=("Processing {0} items..." -f $total)})

    for ($i=0; $i -lt $total; $i++) {
        if ($worker.CancellationPending) { $EventArgs.Cancel = $true; return }
        $it = $allItems[$i]

        $fileRef = "" + $it["FileRef"]
        $fileDir = ("" + $it["FileDirRef"]).TrimEnd("/")
        $leaf    = "" + $it["FileLeafRef"]
        $isFolder = $it.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder

        # Skip internal SharePoint folders like "Forms"
        if ($leaf -eq "Forms" -or $leaf -eq "_t" -or $leaf -eq "_w") { continue }

        # Top folder logic
        $topFolder = "(root)"
        if ($fileDir -eq $rootUrl -or [string]::IsNullOrWhiteSpace($fileDir)) {
            $topFolder = "(root)"
        } elseif ($fileDir.StartsWith($rootUrl + "/")) {
            $rest = $fileDir.Substring($rootUrl.Length + 1)
            $topFolder = $rest.Split("/")[0]
            if ([string]::IsNullOrWhiteSpace($topFolder)) { $topFolder = "(root)" }
        } else {
            $topFolder = "(outside-root?)"
        }

        if (-not $folderCounts.ContainsKey($topFolder)) { $folderCounts[$topFolder] = 0 }
        $folderCounts[$topFolder]++

        $hasUnique = [bool]$it.HasUniqueRoleAssignments
        if ($hasUnique) {
            if (-not $isDryRun) {
                $worker.ReportProgress(-1, @{type="log"; msg=("[ACTION] Resetting permissions on: $fileRef")})
                $retryCount = 0
                $maxRetries = 3
                $success = $false
                while (-not $success -and $retryCount -lt $maxRetries) {
                    try {
                        $it.ResetRoleInheritance()
                        $ctx.ExecuteQuery()
                        $success = $true
                    } catch {
                        if ($_.Exception.Message -like '*429*') {
                            $retryCount++
                            $worker.ReportProgress(-1, @{type="log"; msg=("[WARN] Throttled (429). Retrying in 5 seconds... (Attempt $retryCount)")})
                            Start-Sleep -Seconds 5
                        } else {
                            $worker.ReportProgress(-1, @{type="log"; msg=("[ERROR] Failed to reset $fileRef : " + $_.Exception.Message)})
                            break
                        }
                    }
                }
            }

            $flagged.Add([pscustomobject]@{
                Url = $fileRef
                Name = $leaf
                Type = if ($isFolder) { "Folder" } else { "File" }
                TopFolder = $topFolder
                HasUniquePermissions = $true
                Status = if ($isDryRun) { "Detected" } else { "Reset" }
            })
        }

        if ($sleepSec -gt 0 -and $hasUnique) { Start-Sleep -Seconds $sleepSec }

        if (($i % 100) -eq 0 -or $i -eq ($total - 1)) {
            $pct = 10 + [int](85 * (($i+1) / [double]$total))
            if ($pct -gt 95) { $pct = 95 }
            $worker.ReportProgress($pct, @{type="status"; msg=("Processed {0}/{1}" -f ($i+1), $total)})
        }
    }

    $worker.ReportProgress(96, @{type="status"; msg="Exporting reports..."})
    $flagged | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $flagCsv
    $folderCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            TopLevelFolder = $_.Name
            ItemCount = $_.Value
        }
    } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $countCsv

    $summary = [pscustomobject]@{
        Timestamp = (Get-Date).ToString("o")
        SiteUrl = $siteUrl
        Library = $libTitle
        SharingCapability = $sharingCapability
        TotalItemsScanned = $total
        UniquePermissionsFound = $flagged.Count
        Mode = if ($isDryRun) { "Dry Run" } else { "Reset" }
        Outputs = @{
            FlaggedItemsCsv = $flagCsv
            FolderCountsCsv = $countCsv
        }
    }
    $summary | ConvertTo-Json -Depth 6 | Out-File -Encoding UTF8 -FilePath $summaryJs

    $worker.ReportProgress(100, @{type="log"; msg=("[INFO] Done. Items flagged/processed: {0}" -f $flagged.Count)})
    $EventArgs.Result = $summary
} | Out-Null

Register-ObjectEvent -InputObject $worker -EventName ProgressChanged -Action {
    $payload = $EventArgs.UserState
    $pct = $EventArgs.ProgressPercentage
    $form.BeginInvoke([Action]{
        if ($payload -and $payload.type -eq "log") { UiLog $payload.msg }
        if ($payload -and $payload.type -eq "status") { $lblStatus.Text = $payload.msg }
        if ($pct -ge 0 -and $pct -le 100) { $progress.Value = $pct }
    }) | Out-Null
} | Out-Null

Register-ObjectEvent -InputObject $worker -EventName RunWorkerCompleted -Action {
    $form.BeginInvoke([Action]{
        SetUiRunning $false
        if ($EventArgs.Error) {
            $lblStatus.Text = "Failed"
            $progress.Value = 0
            UiLog ("[ERROR] " + $EventArgs.Error.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show("Process failed:`r`n$($EventArgs.Error.Exception.Message)", "Error", "OK", "Error") | Out-Null
        } elseif ($EventArgs.Cancelled) {
            $lblStatus.Text = "Cancelled"
            $progress.Value = 0
            UiLog "[INFO] Cancelled by user."
        } else {
            $summary = $EventArgs.Result
            $lblStatus.Text = "Completed"
            $progress.Value = 100
            UiLog "[INFO] Completed successfully."
            [System.Windows.Forms.MessageBox]::Show("Process completed.`r`nSharing Capability: $($summary.SharingCapability)`r`nItems with unique permissions: $($summary.UniquePermissionsFound)`r`nCheck output folder for reports.", "Done", "OK", "Information") | Out-Null
        }
    }) | Out-Null
} | Out-Null

# ---------------- Event Handlers ----------------
$btnStart.Add_Click({
    $site = $txtSite.Text.Trim()
    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text
    if ([string]::IsNullOrWhiteSpace($site) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
        [System.Windows.Forms.MessageBox]::Show("Please fill Site URL, Username, and App Password.", "Input Missing") | Out-Null
        return
    }

    $txtLog.Clear()
    $progress.Value = 0
    $lblStatus.Text = "Starting..."
    SetUiRunning $true
    UiLog "[INFO] Starting process..."

    $worker.RunWorkerAsync(@{
        SiteUrl = $site
        Library = $txtLib.Text.Trim()
        Username = $user
        Password = $pass
        AdminUrl = $txtAdmin.Text.Trim()
        Output   = $txtOut.Text.Trim()
        SleepSec = [int]$numSleep.Value
        IsDryRun = $rbDryRun.Checked
    })
})

$btnCancel.Add_Click({
    if ($worker.IsBusy) {
        UiLog "[INFO] Cancel requested..."
        $worker.CancelAsync()
    }
})

SetUiRunning $false
[void]$form.ShowDialog()
