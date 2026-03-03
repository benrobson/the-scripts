#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Online Permissions Tool (GUI)
    - Detects unique permissions and counts items (Dry Run)
    - Detects site-level external sharing settings
    - Optionally resets unique permissions (inheritance)
    - Compatible with PowerShell 5.1 and uses CSOM (Username + App Password)

.DESCRIPTION
    This script provides a GUI to manage SharePoint Online permissions. It uses Runspaces
    in STA mode to ensure UI responsiveness and compatibility with SPO authentication.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

# ---- CSOM DLL Loader ----
function Load-SharePointAssemblies {
    $found = $false
    $searchPaths = @(
        "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI",
        "C:\Program Files (x86)\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI",
        "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI",
        "C:\Program Files (x86)\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI"
    )

    foreach ($path in $searchPaths) {
        $clientDll = Join-Path $path "Microsoft.SharePoint.Client.dll"
        if (Test-Path $clientDll) {
            try {
                Add-Type -Path $clientDll -ErrorAction SilentlyContinue
                Add-Type -Path (Join-Path $path "Microsoft.SharePoint.Client.Runtime.dll") -ErrorAction SilentlyContinue
                $tenantDll = Join-Path $path "Microsoft.Online.SharePoint.Client.Tenant.dll"
                if (Test-Path $tenantDll) { Add-Type -Path $tenantDll -ErrorAction SilentlyContinue }
                $found = $true
                return $true
            } catch { }
        }
    }

    if (-not $found) {
        try {
            Add-Type -AssemblyName "Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" -ErrorAction SilentlyContinue
            Add-Type -AssemblyName "Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" -ErrorAction SilentlyContinue
            $found = $true
            return $true
        } catch { }
    }
    return $false
}

if (-not (Load-SharePointAssemblies)) {
    [System.Windows.Forms.MessageBox]::Show("SharePoint Client DLLs not found. Please install SharePoint Online Client Components SDK.", "Error", "OK", "Error") | Out-Null
    return
}

# ---------------- GUI Initialization ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SPO Permissions Manager (CSOM)"
$form.Size = New-Object System.Drawing.Size(980, 850)
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

# Shared Data for Runspace Communication
$logQueue = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$syncHash = [hashtable]::Synchronized(@{
    Status = "Idle"
    Progress = 0
    CancelRequested = $false
    IsRunning = $false
    Result = $null
    Error = $null
    Completed = $false
})

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

$form.Controls.Add((New-Label "Sleep (sec) between items (only if unique)" 320 210))
$numSleep = New-Object System.Windows.Forms.NumericUpDown
$numSleep.Location = New-Object System.Drawing.Point(320, 235)
$numSleep.Size = New-Object System.Drawing.Size(100, 22)
$numSleep.Minimum = 0
$numSleep.Maximum = 10
$numSleep.Value = 1
$form.Controls.Add($numSleep)

# Buttons
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Test Connection"
$btnTest.Location = New-Object System.Drawing.Point(10, 280)
$btnTest.Size = New-Object System.Drawing.Size(155, 32)
$form.Controls.Add($btnTest)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Process"
$btnStart.Location = New-Object System.Drawing.Point(175, 280)
$btnStart.Size = New-Object System.Drawing.Size(155, 32)
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(10, 320)
$btnCancel.Size = New-Object System.Drawing.Size(155, 32)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(175, 326)
$progress.Size = New-Object System.Drawing.Size(780, 20)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Idle"
$lblStatus.Location = New-Object System.Drawing.Point(175, 348)
$lblStatus.Size = New-Object System.Drawing.Size(780, 20)
$form.Controls.Add($lblStatus)

# Log
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 380)
$txtLog.Size = New-Object System.Drawing.Size(945, 410)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

$folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog

$btnBrowse.Add_Click({
    if ($folderDlg.ShowDialog() -eq "OK") { $txtOut.Text = $folderDlg.SelectedPath }
})

function SetUiRunning([bool]$running) {
    $btnStart.Enabled = -not $running
    $btnTest.Enabled = -not $running
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

# ---------------- Background Runspace Logic ----------------

$backgroundScript = {
    param($data, $syncHash, $logQueue)

    $ErrorActionPreference = "Stop"

    function Log($msg) {
        $logQueue.Add("[$((Get-Date).ToString('HH:mm:ss'))] $msg`r`n")
    }

    function Execute-QueryWithRetry($context) {
        $retry = $true
        $retryCount = 0
        while ($retry -and $retryCount -lt 3) {
            try {
                $context.ExecuteQuery()
                $retry = $false
            } catch {
                if ($_.Exception.Message -like "*429*" -or $_.Exception.Message -like "*503*") {
                    $retryCount++
                    Log "[WARN] Throttled (429/503). Waiting 5 seconds... (Attempt $retryCount)"
                    Start-Sleep -Seconds 5
                } else { throw }
            }
        }
    }

    # Runspaces do not inherit loaded assemblies.
    function Load-Assemblies-Internal {
        $searchPaths = @(
            "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI",
            "C:\Program Files (x86)\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI",
            "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI",
            "C:\Program Files (x86)\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI"
        )
        foreach ($path in $searchPaths) {
            $clientDll = Join-Path $path "Microsoft.SharePoint.Client.dll"
            if (Test-Path $clientDll) {
                Add-Type -Path $clientDll -ErrorAction SilentlyContinue
                Add-Type -Path (Join-Path $path "Microsoft.SharePoint.Client.Runtime.dll") -ErrorAction SilentlyContinue
                $tenantDll = Join-Path $path "Microsoft.Online.SharePoint.Client.Tenant.dll"
                if (Test-Path $tenantDll) { Add-Type -Path $tenantDll -ErrorAction SilentlyContinue }
                return $true
            }
        }
        Add-Type -AssemblyName "Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" -ErrorAction SilentlyContinue
        Add-Type -AssemblyName "Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" -ErrorAction SilentlyContinue
        return $true
    }

    try {
        $syncHash.IsRunning = $true
        Log "[INFO] Background process started."
        Load-Assemblies-Internal | Out-Null

        $siteUrl  = $data.SiteUrl
        $libTitle = $data.Library
        $user     = $data.Username
        $pass     = $data.Password
        $adminUrl = $data.AdminUrl
        $out      = $data.Output
        $sleepSec = [int]$data.SleepSec
        $isDryRun = [bool]$data.IsDryRun
        $isTest   = [bool]$data.IsTest

        $secure = ConvertTo-SecureString -String $pass -AsPlainText -Force

        if ($isTest) {
            Log "[INFO] Testing connection to $siteUrl..."
            $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
            $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)
            $ctx.Load($ctx.Web)
            Execute-QueryWithRetry $ctx
            Log "[SUCCESS] Connected to site: $($ctx.Web.Title)"
            $syncHash.Status = "Test Successful"
            return
        }

        New-Item -ItemType Directory -Force -Path $out | Out-Null
        $ts = Get-Date -Format "yyyyMMdd-HHmmss"
        $flagCsv   = Join-Path $out "permissions-flagged-items-$ts.csv"
        $countCsv  = Join-Path $out "folder-counts-$ts.csv"
        $summaryJs = Join-Path $out "summary-$ts.json"

        $sharingCapability = "Unknown"
        if (-not [string]::IsNullOrWhiteSpace($adminUrl)) {
            Log "[INFO] Checking Sharing Capability via Tenant Admin..."
            try {
                $adminCtx = New-Object Microsoft.SharePoint.Client.ClientContext($adminUrl)
                $adminCtx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)
                $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($adminCtx)
                $siteProperties = $tenant.GetSitePropertiesByUrl($siteUrl, $true)
                $adminCtx.Load($siteProperties)
                Execute-QueryWithRetry $adminCtx
                $sharingCapability = $siteProperties.SharingCapability.ToString()
                Log "[INFO] Site Sharing Capability: $sharingCapability"
            } catch { Log "[WARN] Failed to check Sharing Capability: $($_.Exception.Message)" }
        }

        Log "[INFO] Connecting to Library: $libTitle..."
        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
        $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)
        $list = $ctx.Web.Lists.GetByTitle($libTitle)
        $ctx.Load($list)
        $ctx.Load($list.RootFolder)
        Execute-QueryWithRetry $ctx
        $rootUrl = $list.RootFolder.ServerRelativeUrl.TrimEnd("/")

        Log "[INFO] Enumerate items (Recursive)..."
        $allItems = New-Object System.Collections.Generic.List[object]
        $position = $null
        do {
            if ($syncHash.CancelRequested) { Log "[INFO] Cancellation requested."; return }
            $qry = New-Object Microsoft.SharePoint.Client.CamlQuery
            $qry.ViewXml = "<View Scope='RecursiveAll'><RowLimit>2000</RowLimit></View>"
            $qry.ListItemCollectionPosition = $position
            $items = $list.GetItems($qry)

            # Using Include syntax in a single string is standard for PowerShell CSOM property loading
            $ctx.Load($items, "Include(FileRef, FileDirRef, FileLeafRef, HasUniqueRoleAssignments, FileSystemObjectType)")
            Execute-QueryWithRetry $ctx

            $position = $items.ListItemCollectionPosition
            foreach ($it in $items) { $allItems.Add($it) }
            Log "[INFO] Retrieved $($allItems.Count) items..."
            $syncHash.Status = "Retrieved $($allItems.Count) items"
        } while ($position -ne $null)

        $total = $allItems.Count
        $folderCounts = @{}
        $flagged = New-Object System.Collections.Generic.List[object]

        Log "[INFO] Processing $total items..."
        for ($i=0; $i -lt $total; $i++) {
            if ($syncHash.CancelRequested) { Log "[INFO] Cancellation requested."; return }
            $it = $allItems[$i]

            $fileRef = "" + $it["FileRef"]
            $fileDir = ("" + $it["FileDirRef"]).TrimEnd("/")
            $leaf    = "" + $it["FileLeafRef"]
            # FSObjType: 0=File, 1=Folder
            $isFolder = ($it.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder) -or ($it["FileSystemObjectType"] -eq 1)

            if ($leaf -eq "Forms" -or $leaf -eq "_t" -or $leaf -eq "_w") { continue }

            $topFolder = "(root)"
            $relPath = ""
            if ($fileDir.Length -gt $rootUrl.Length) {
                $relPath = $fileDir.Substring($rootUrl.Length).TrimStart("/")
            }

            if ([string]::IsNullOrWhiteSpace($relPath)) {
                $topFolder = "(root)"
            } else {
                $topFolder = $relPath.Split("/")[0]
            }

            if (-not $folderCounts.ContainsKey($topFolder)) { $folderCounts[$topFolder] = 0 }
            $folderCounts[$topFolder]++

            if ([bool]$it.HasUniqueRoleAssignments) {
                if (-not $isDryRun) {
                    Log "[ACTION] Resetting Inheritance: $fileRef"
                    $it.ResetRoleInheritance()
                    Execute-QueryWithRetry $ctx
                }
                $flagged.Add([pscustomobject]@{ Url = $fileRef; Name = $leaf; Type = if ($isFolder) { "Folder" } else { "File" }; TopFolder = $topFolder; HasUniquePermissions = $true; Status = if ($isDryRun) { "Detected" } else { "Reset" } })
            }

            if ($sleepSec -gt 0 -and $it.HasUniqueRoleAssignments) { Start-Sleep -Seconds $sleepSec }

            $syncHash.Progress = [int](10 + (90 * ($i / $total)))
            $syncHash.Status = "Processing $i / $total"
        }

        Log "[INFO] Exporting reports..."
        $flagged | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $flagCsv
        $folderCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { [pscustomobject]@{ TopLevelFolder = $_.Name; ItemCount = $_.Value } } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $countCsv

        $summary = @{ Timestamp = (Get-Date).ToString("o"); SiteUrl = $siteUrl; Library = $libTitle; SharingCapability = $sharingCapability; TotalItemsScanned = $total; UniquePermissionsFound = $flagged.Count; Mode = if ($isDryRun) { "Dry Run" } else { "Reset" }; Outputs = @{ FlaggedItemsCsv = $flagCsv; FolderCountsCsv = $countCsv } }
        $summary | ConvertTo-Json -Depth 6 | Out-File -Encoding UTF8 -FilePath $summaryJs
        $syncHash.Result = $summary
        Log "[INFO] Completed successfully."
    } catch {
        Log "[ERROR] $($_.Exception.Message)"
        $syncHash.Error = $_.Exception.Message
    } finally {
        $syncHash.IsRunning = $false
        $syncHash.Completed = $true
    }
}

# ---------------- UI Timer to update from SyncHash ----------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200
$timer.Add_Tick({
    if ($logQueue.Count -gt 0) {
        [System.Threading.Monitor]::Enter($logQueue.SyncRoot)
        try {
            $logs = ""
            foreach ($l in $logQueue) { $logs += $l }
            $logQueue.Clear()
            $txtLog.AppendText($logs)
        } finally {
            [System.Threading.Monitor]::Exit($logQueue.SyncRoot)
        }
    }

    $lblStatus.Text = $syncHash.Status
    $progress.Value = $syncHash.Progress

    if ($syncHash.Completed) {
        $timer.Stop()
        $syncHash.Completed = $false
        SetUiRunning $false
        if ($syncHash.Error) {
            [System.Windows.Forms.MessageBox]::Show("Error: $($syncHash.Error)", "Error", "OK", "Error") | Out-Null
        } elseif ($syncHash.Status -eq "Test Successful") {
            [System.Windows.Forms.MessageBox]::Show("Test Connection Successful!", "Success") | Out-Null
        } elseif ($syncHash.Result) {
            [System.Windows.Forms.MessageBox]::Show("Process completed.`n`nFound $($syncHash.Result.UniquePermissionsFound) items with unique permissions.", "Done") | Out-Null
        }
    }
})

function Start-Runspace-Logic($data) {
    $syncHash.Progress = 0
    $syncHash.Status = "Starting..."
    $syncHash.CancelRequested = $false
    $syncHash.Result = $null
    $syncHash.Error = $null
    $syncHash.Completed = $false
    $logQueue.Clear()
    SetUiRunning $true

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
    $runspace.SessionStateProxy.SetVariable("logQueue", $logQueue)

    $powershell = [PowerShell]::Create().AddScript($backgroundScript).AddArgument($data).AddArgument($syncHash).AddArgument($logQueue)
    $powershell.Runspace = $runspace
    $powershell.BeginInvoke()
    $timer.Start()
}

$btnStart.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text) -or [string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please fill all credentials.") | Out-Null; return
    }
    $txtLog.Clear()
    $data = @{ SiteUrl = $txtSite.Text.Trim(); Library = $txtLib.Text.Trim(); Username = $txtUser.Text.Trim(); Password = $txtPass.Text; AdminUrl = $txtAdmin.Text.Trim(); Output = $txtOut.Text.Trim(); SleepSec = $numSleep.Value; IsDryRun = $rbDryRun.Checked; IsTest = $false }
    Start-Runspace-Logic $data
})

$btnTest.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text) -or [string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please fill Site URL, Username, and App Password.") | Out-Null; return
    }
    $data = @{ SiteUrl = $txtSite.Text.Trim(); Username = $txtUser.Text.Trim(); Password = $txtPass.Text; IsTest = $true }
    Start-Runspace-Logic $data
})

$btnCancel.Add_Click({
    $syncHash.CancelRequested = $true
    $lblStatus.Text = "Cancellation Requested..."
})

[void]$form.ShowDialog()
