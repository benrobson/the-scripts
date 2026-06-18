#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Online & Atlas RBAC Tool (GUI)
    - Detects unique permissions and counts items (Dry Run)
    - Detects site-level external sharing settings
    - Optionally resets unique permissions (inheritance)
    - Ability to provision Atlas groups and invite users (RBAC)
    - Compatible with PowerShell 5.1 and uses CSOM (Username + App Password)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

# ---- CSOM DLL Loader ----
function Load-SharePointAssemblies {
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
                return $true
            } catch { }
        }
    }
    try {
        Add-Type -AssemblyName "Microsoft.SharePoint.Client, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" -ErrorAction SilentlyContinue
        Add-Type -AssemblyName "Microsoft.SharePoint.Client.Runtime, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" -ErrorAction SilentlyContinue
        return $true
    } catch { return $false }
}

if (-not (Load-SharePointAssemblies)) {
    [System.Windows.Forms.MessageBox]::Show("SharePoint Client DLLs not found. Please install SharePoint Online Client Components SDK.", "Error", "OK", "Error") | Out-Null
    return
}

# ---------------- GUI Initialization ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SPO & Atlas Management Suite"
$form.Size = New-Object System.Drawing.Size(1100, 950)
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

# Common Inputs (Static)
$form.Controls.Add((New-Label "Site URL" 10 10))
$txtSite = New-TextBox 10 30 1065 ""
$form.Controls.Add($txtSite)

$form.Controls.Add((New-Label "Username (UPN)" 10 60))
$txtUser = New-TextBox 10 80 500 ""
$form.Controls.Add($txtUser)

$form.Controls.Add((New-Label "App Password" 530 60))
$txtPass = New-TextBox 530 80 545 ""
$txtPass.UseSystemPasswordChar = $true
$form.Controls.Add($txtPass)

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 115)
$tabControl.Size = New-Object System.Drawing.Size(1065, 230)
$form.Controls.Add($tabControl)

# --- Tab 1: Permissions Audit/Reset ---
$tabPerms = New-Object System.Windows.Forms.TabPage
$tabPerms.Text = "Permissions & Audit"
$tabControl.TabPages.Add($tabPerms)

$tabPerms.Controls.Add((New-Label "Library Title" 10 10))
$txtLib = New-TextBox 10 30 240 "Documents"
$tabPerms.Controls.Add($txtLib)

$tabPerms.Controls.Add((New-Label "Tenant Admin URL (Optional, for Sharing check)" 270 10))
$txtAdmin = New-TextBox 270 30 780 ""
$tabPerms.Controls.Add($txtAdmin)

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "Execution Mode"
$grpMode.Location = New-Object System.Drawing.Point(10, 65)
$grpMode.Size = New-Object System.Drawing.Size(300, 60)
$tabPerms.Controls.Add($grpMode)

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

$tabPerms.Controls.Add((New-Label "Sleep (sec) between unique items" 320 65))
$numSleep = New-Object System.Windows.Forms.NumericUpDown
$numSleep.Location = New-Object System.Drawing.Point(320, 90)
$numSleep.Size = New-Object System.Drawing.Size(100, 22)
$numSleep.Minimum = 0
$numSleep.Maximum = 10
$numSleep.Value = 1
$tabPerms.Controls.Add($numSleep)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Scan/Reset"
$btnStart.Location = New-Object System.Drawing.Point(10, 140)
$btnStart.Size = New-Object System.Drawing.Size(155, 32)
$tabPerms.Controls.Add($btnStart)

$tabPerms.Controls.Add((New-Label "Output Folder" 270 65))
$txtOut = New-TextBox 270 90 600 (Join-Path $PWD "SPO-Permissions-Output")
$tabPerms.Controls.Add($txtOut)

# --- Tab 2: Atlas RBAC & Invitations ---
$tabAtlas = New-Object System.Windows.Forms.TabPage
$tabAtlas.Text = "Atlas RBAC & Invites"
$tabControl.TabPages.Add($tabAtlas)

$tabAtlas.Controls.Add((New-Label "Atlas User Emails (One per line)" 10 10))
$txtAtlasUsers = New-Object System.Windows.Forms.TextBox
$txtAtlasUsers.Location = New-Object System.Drawing.Point(10, 30)
$txtAtlasUsers.Size = New-Object System.Drawing.Size(400, 100)
$txtAtlasUsers.Multiline = $true
$txtAtlasUsers.ScrollBars = "Vertical"
$tabAtlas.Controls.Add($txtAtlasUsers)

$tabAtlas.Controls.Add((New-Label "Target RBAC Role" 430 10))
$cmbAtlasRole = New-Object System.Windows.Forms.ComboBox
$cmbAtlasRole.Location = New-Object System.Drawing.Point(430, 30)
$cmbAtlasRole.Size = New-Object System.Drawing.Size(250, 22)
$cmbAtlasRole.DropDownStyle = "DropDownList"
[void]$cmbAtlasRole.Items.Add("Atlas Admins (Full Control)")
[void]$cmbAtlasRole.Items.Add("Atlas Contributors (Edit)")
[void]$cmbAtlasRole.Items.Add("Atlas Readers (Read)")
$cmbAtlasRole.SelectedIndex = 1
$tabAtlas.Controls.Add($cmbAtlasRole)

$tabAtlas.Controls.Add((New-Label "Custom Group Name Prefix (Optional)" 430 65))
$txtAtlasPrefix = New-TextBox 430 85 250 "Atlas"
$tabAtlas.Controls.Add($txtAtlasPrefix)

$btnAtlas = New-Object System.Windows.Forms.Button
$btnAtlas.Text = "Provision Atlas RBAC"
$btnAtlas.Location = New-Object System.Drawing.Point(430, 140)
$btnAtlas.Size = New-Object System.Drawing.Size(200, 32)
$tabAtlas.Controls.Add($btnAtlas)

# Common Bottom Controls
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Test Connection + Diags"
$btnTest.Location = New-Object System.Drawing.Point(10, 355)
$btnTest.Size = New-Object System.Drawing.Size(200, 32)
$form.Controls.Add($btnTest)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(220, 355)
$btnCancel.Size = New-Object System.Drawing.Size(155, 32)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(385, 361)
$progress.Size = New-Object System.Drawing.Size(690, 20)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Idle"
$lblStatus.Location = New-Object System.Drawing.Point(385, 383)
$lblStatus.Size = New-Object System.Drawing.Size(690, 20)
$form.Controls.Add($lblStatus)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 410)
$txtLog.Size = New-Object System.Drawing.Size(1065, 480)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

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
                $msg = $_.Exception.Message
                if ($msg -like "*429*" -or $msg -like "*503*") {
                    $retryCount++
                    Log "[WARN] Throttled (429/503). Waiting 5 seconds... (Attempt $retryCount)"
                    Start-Sleep -Seconds 5
                } else { throw }
            }
        }
    }

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
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

        $siteUrl  = $data.SiteUrl
        $user     = $data.Username
        $pass     = $data.Password
        $secure = ConvertTo-SecureString -String $pass -AsPlainText -Force

        if ($data.Task -eq "Test") {
            Log "[INFO] Testing connection to $siteUrl..."
            $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
            $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)
            $ctx.Load($ctx.Web)
            Execute-QueryWithRetry $ctx
            Log "[SUCCESS] Connected to site: $($ctx.Web.Title)"
            $syncHash.Status = "Test Successful"
            return
        }

        if ($data.Task -eq "Permissions") {
            $libTitle = $data.Library
            $adminUrl = $data.AdminUrl
            $sleepSec = $data.SleepSec
            $isDryRun = $data.IsDryRun
            $out      = $data.Output

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

            Log "[INFO] Starting Permissions Task for $libTitle..."
            $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
            $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)

            $list = $ctx.Web.Lists.GetByTitle($libTitle)
            $ctx.Load($list)
            $ctx.Load($list.RootFolder)
            Execute-QueryWithRetry $ctx
            $rootUrl = $list.RootFolder.ServerRelativeUrl.TrimEnd("/")

            Log "[INFO] Enumerating items..."
            $allItems = New-Object System.Collections.Generic.List[object]
            $position = $null
            do {
                if ($syncHash.CancelRequested) { Log "[INFO] Cancellation requested."; return }
                $qry = New-Object Microsoft.SharePoint.Client.CamlQuery
                $qry.ViewXml = "<View Scope='RecursiveAll'><RowLimit>2000</RowLimit></View>"
                $qry.ListItemCollectionPosition = $position
                $items = $list.GetItems($qry)

                # Single-parameter Load is compatible with PS 5.1 CSOM
                $ctx.Load($items)
                Execute-QueryWithRetry $ctx

                # Batch retrieval of required properties for all items in page
                foreach ($it in $items) {
                    # Note: We can't use Include in string in Load($items) easily in raw CSOM,
                    # so we queue properties for each item.
                    $ctx.Load($it, "FileRef", "FileDirRef", "FileLeafRef", "HasUniqueRoleAssignments", "FileSystemObjectType")
                }
                Execute-QueryWithRetry $ctx

                $position = $items.ListItemCollectionPosition
                foreach ($it in $items) { $allItems.Add($it) }
                Log "[INFO] Retrieved $($allItems.Count) items..."
                $syncHash.Status = "Retrieved $($allItems.Count) items"
            } while ($position -ne $null)

            $total = $allItems.Count
            $folderCounts = @{}
            $flagged = New-Object System.Collections.Generic.List[object]

            for ($i=0; $i -lt $total; $i++) {
                if ($syncHash.CancelRequested) { Log "[INFO] Cancellation requested."; return }
                $it = $allItems[$i]
                $fileRef = "" + $it["FileRef"]
                $fileDir = ("" + $it["FileDirRef"]).TrimEnd("/")
                $leaf    = "" + $it["FileLeafRef"]
                $isFolder = ($it.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder) -or ($it["FileSystemObjectType"] -eq 1)

                if ($leaf -eq "Forms") { continue }

                $relPath = if ($fileDir.Length -gt $rootUrl.Length) { $fileDir.Substring($rootUrl.Length).TrimStart("/") } else { "" }
                $topFolder = if ([string]::IsNullOrWhiteSpace($relPath)) { "(root)" } else { $relPath.Split("/")[0] }

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

            Log "[INFO] Task completed. Items flagged: $($flagged.Count)"
            $syncHash.Result = $summary
        }

        if ($data.Task -eq "Atlas") {
            Log "[INFO] Starting Atlas RBAC Provisioning..."
            $emails = $data.Users -split "`r`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $roleName = $data.Role
            $prefix = $data.Prefix

            $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
            $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure)
            $web = $ctx.Web
            $ctx.Load($web)
            Execute-QueryWithRetry $ctx

            $groupName = "$prefix - $roleName"
            Log "[INFO] Targeting SharePoint Group: $groupName"

            $group = $null
            try {
                $group = $web.SiteGroups.GetByName($groupName)
                $ctx.Load($group)
                Execute-QueryWithRetry $ctx
                Log "[INFO] Group already exists."
            } catch {
                Log "[INFO] Creating group..."
                $groupInfo = New-Object Microsoft.SharePoint.Client.GroupCreationInformation
                $groupInfo.Title = $groupName
                $group = $web.SiteGroups.Add($groupInfo)
                $ctx.Load($group)
                Execute-QueryWithRetry $ctx
                Log "[SUCCESS] Group created."
            }

            # Assign Permission Level
            Log "[INFO] Ensuring permissions for group..."
            $roleDefName = if ($roleName -match "Admins") { "Full Control" } elseif ($roleName -match "Contributors") { "Edit" } else { "Read" }
            $roleDef = $web.RoleDefinitions.GetByName($roleDefName)
            $roleAssign = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
            $roleAssign.Add($roleDef)
            $web.RoleAssignments.Add($group, $roleAssign)
            Execute-QueryWithRetry $ctx
            Log "[SUCCESS] Permissions set to $roleDefName"

            # Add Users
            foreach ($email in $emails) {
                if ($syncHash.CancelRequested) { return }
                try {
                    Log "[INFO] Inviting user: $email"
                    $ensureUser = $web.EnsureUser($email.Trim())
                    $ctx.Load($ensureUser)
                    Execute-QueryWithRetry $ctx

                    $group.Users.AddUser($ensureUser)
                    Execute-QueryWithRetry $ctx
                    Log "[SUCCESS] User $email added to group."
                } catch {
                    Log "[ERROR] Failed to add user $email : $($_.Exception.Message)"
                }
            }
            Log "[INFO] Atlas Provisioning Completed."
            $syncHash.Status = "Atlas Provisioning Completed"
        }

    } catch {
        Log "[ERROR] $($_.Exception.Message)"
        $syncHash.Error = $_.Exception.Message
    } finally {
        $syncHash.IsRunning = $false
        $syncHash.Completed = $true
    }
}

# ---------------- UI Timer ----------------
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
        $timer.Stop(); $syncHash.Completed = $false; SetUiRunning $false
        if ($syncHash.Error) { [System.Windows.Forms.MessageBox]::Show("Error: $($syncHash.Error)", "Error") | Out-Null }
        elseif ($syncHash.Status -eq "Test Successful") { [System.Windows.Forms.MessageBox]::Show("Connected!", "Success") | Out-Null }
        else { [System.Windows.Forms.MessageBox]::Show("Operation Completed.", "Done") | Out-Null }
    }
})

function Start-Task($data) {
    SetUiRunning $true
    $syncHash.Progress = 0; $syncHash.Status = "Starting..."; $syncHash.CancelRequested = $false; $syncHash.Completed = $false
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"; $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
    $runspace.SessionStateProxy.SetVariable("logQueue", $logQueue)
    $powershell = [PowerShell]::Create().AddScript($backgroundScript).AddArgument($data).AddArgument($syncHash).AddArgument($logQueue)
    $powershell.Runspace = $runspace
    $powershell.BeginInvoke()
    $timer.Start()
}

$btnStart.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text)) { return }
    $data = @{ Task="Permissions"; SiteUrl=$txtSite.Text; Username=$txtUser.Text; Password=$txtPass.Text; Library=$txtLib.Text; AdminUrl=$txtAdmin.Text; SleepSec=$numSleep.Value; IsDryRun=$rbDryRun.Checked; Output=$txtOut.Text }
    Start-Task $data
})

$btnAtlas.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text) -or [string]::IsNullOrWhiteSpace($txtAtlasUsers.Text)) { return }
    $data = @{ Task="Atlas"; SiteUrl=$txtSite.Text; Username=$txtUser.Text; Password=$txtPass.Text; Users=$txtAtlasUsers.Text; Role=$cmbAtlasRole.Text; Prefix=$txtAtlasPrefix.Text }
    Start-Task $data
})

$btnTest.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text)) { return }
    Start-Task @{ Task="Test"; SiteUrl=$txtSite.Text; Username=$txtUser.Text; Password=$txtPass.Text }
})

$btnCancel.Add_Click({ $syncHash.CancelRequested = $true; $lblStatus.Text = "Cancelling..." })

SetUiRunning $false
[void]$form.ShowDialog()
