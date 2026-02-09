<#
.SYNOPSIS
    SPO Permission Resetter GUI (v2)
.DESCRIPTION
    A WPF-based PowerShell script to reset unique permissions on folders and files in a SharePoint Online document library.
    Uses PnP.PowerShell for modern authentication and CSOM for permission reset.
.NOTES
    Requirement: PnP.PowerShell module
    Run in STA mode: powershell.exe -STA -File .\Reset-SPOFolderAndFilesPermissions-GUI.ps1
#>

param (
    [Parameter(Mandatory=$false)]
    [string]$SiteURL = "https://tenant.sharepoint.com/sites/yoursite"
)

# Ensure we are running in STA mode for WPF
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Re-launching in STA mode..." -ForegroundColor Yellow
    if ($PSBoundParameters.ContainsKey('SiteURL')) {
        powershell.exe -STA -File "$PSCommandPath" -SiteURL "$SiteURL"
    } else {
        powershell.exe -STA -File "$PSCommandPath"
    }
    return
}

# --- Load Assemblies ---
Try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Xaml
    Add-Type -AssemblyName System.Windows.Forms
} Catch {
    # Fallback for some environments
    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')
    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')
    [void][System.Reflection.Assembly]::LoadWithPartialName('WindowsBase')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Xaml')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
}

# --- XAML UI Definition ---
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2000/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2001/xaml"
        Title="SPO Permission Resetter v2" Height="650" Width="850" Background="#F0F0F0">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Connection -->
            <RowDefinition Height="Auto"/> <!-- Selection -->
            <RowDefinition Height="Auto"/> <!-- Stats -->
            <RowDefinition Height="*"/>    <!-- Log -->
            <RowDefinition Height="Auto"/> <!-- Progress -->
            <RowDefinition Height="Auto"/> <!-- Buttons -->
        </Grid.RowDefinitions>

        <!-- Connection Section -->
        <GroupBox Header="1. Connection" Grid.Row="0" Margin="0,0,0,10" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Content="Site URL:" VerticalAlignment="Center"/>
                <TextBox Name="txtSiteUrl" Grid.Column="1" Margin="10,0" VerticalContentAlignment="Center" Height="25" Text="https://tenant.sharepoint.com/sites/yoursite"/>
                <Button Name="btnConnect" Grid.Column="2" Content="Connect" Width="100" Height="25"/>
            </Grid>
        </GroupBox>

        <!-- Selection Section -->
        <GroupBox Header="2. Library &amp; Folder Selection" Grid.Row="1" Margin="0,0,0,10" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Label Content="Library:" VerticalAlignment="Center"/>
                <ComboBox Name="cmbLibrary" Grid.Column="1" Margin="10,0" Height="25" IsEnabled="False"/>

                <Label Content="Folder Path:" Grid.Column="2" VerticalAlignment="Center"/>
                <TextBox Name="txtFolderPath" Grid.Column="3" Margin="10,0" VerticalContentAlignment="Center" Height="25" IsEnabled="False" ToolTip="Relative path (e.g. Shared Documents/Subfolder). Leave blank for root."/>
            </Grid>
        </GroupBox>

        <!-- Stats Section -->
        <GroupBox Header="3. Scan Results" Grid.Row="2" Margin="0,0,0,10" Padding="10">
            <UniformGrid Columns="4">
                <StackPanel HorizontalAlignment="Center">
                    <Label Content="Total Folders" HorizontalAlignment="Center"/>
                    <TextBlock Name="lblFolders" Text="0" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
                </StackPanel>
                <StackPanel HorizontalAlignment="Center">
                    <Label Content="Total Files" HorizontalAlignment="Center"/>
                    <TextBlock Name="lblFiles" Text="0" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
                </StackPanel>
                <StackPanel HorizontalAlignment="Center">
                    <Label Content="Total Size" HorizontalAlignment="Center"/>
                    <TextBlock Name="lblSize" Text="0 MB" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
                </StackPanel>
                <StackPanel HorizontalAlignment="Center">
                    <Label Content="Est. Time" HorizontalAlignment="Center"/>
                    <TextBlock Name="lblEstTime" Text="0s" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
                </StackPanel>
            </UniformGrid>
        </GroupBox>

        <!-- Log Window -->
        <TextBox Name="txtLog" Grid.Row="3" IsReadOnly="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True"
                 Background="Black" Foreground="#00FF00" FontFamily="Consolas" FontSize="12" Margin="0,0,0,10"/>

        <!-- Progress Bar -->
        <ProgressBar Name="progressBar" Grid.Row="4" Height="20" Margin="0,0,0,10" Minimum="0" Maximum="100"/>

        <!-- Action Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <CheckBox Name="chkDryRun" Content="Dry Run (Log only)" VerticalAlignment="Center" IsChecked="True" Margin="0,0,20,0"/>
            <Button Name="btnScan" Content="Scan Folder Tree" Width="120" Height="30" Margin="0,0,10,0" IsEnabled="False"/>
            <Button Name="btnExecute" Content="Execute Reset" Width="120" Height="30" Background="#FF4B4B" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
        </StackPanel>
    </Grid>
</Window>
"@

# --- Load XAML ---
Try {
    $xaml = $xaml -replace 'x:Name', 'Name'
    [xml]$xml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $window = [Windows.Markup.XamlReader]::Load($reader)
} Catch {
    Write-Host "Error loading XAML: $($_.Exception.Message)" -ForegroundColor Red
    If ($_.Exception.InnerException) { Write-Host "Inner Error: $($_.Exception.InnerException.Message)" -ForegroundColor Red }
    return
}

# --- UI Elements ---
$txtSiteUrl = $window.FindName("txtSiteUrl")
$txtSiteUrl.Text = $SiteURL
$btnConnect = $window.FindName("btnConnect")
$cmbLibrary = $window.FindName("cmbLibrary")
$txtFolderPath = $window.FindName("txtFolderPath")
$lblFolders = $window.FindName("lblFolders")
$lblFiles = $window.FindName("lblFiles")
$lblSize = $window.FindName("lblSize")
$lblEstTime = $window.FindName("lblEstTime")
$txtLog = $window.FindName("txtLog")
$progressBar = $window.FindName("progressBar")
$chkDryRun = $window.FindName("chkDryRun")
$btnScan = $window.FindName("btnScan")
$btnExecute = $window.FindName("btnExecute")

# --- State ---
$global:ItemsToProcess = New-Object System.Collections.Generic.List[PSObject]

$window.Add_Loaded({
    Write-Log "Welcome! Please enter a Site URL and click 'Connect' to begin."
    Write-Log "Authentication as a Global Admin or SharePoint Admin is required."
})

# --- Helpers ---
Function Start-Connection {
    $url = $txtSiteUrl.Text.Trim()
    If (-not $url) { [System.Windows.MessageBox]::Show("Please enter a Site URL."); return }

    Write-Log "-----------------------------------------------------------------------"
    Write-Log " AUTHENTICATION REQUIRED"
    Write-Log " Please log in with a Global Administrator or SharePoint Admin account"
    Write-Log " in the browser window that appears."
    Write-Log "-----------------------------------------------------------------------"

    Write-Log "Connecting to $url..."
    Try {
        Connect-PnPOnline -Url $url -Interactive -ErrorAction Stop
        Write-Log "Connected successfully!" -Type "SUCCESS"

        $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }
        $cmbLibrary.ItemsSource = $libraries
        $cmbLibrary.DisplayMemberPath = "Title"
        $cmbLibrary.IsEnabled = $True
        $txtFolderPath.IsEnabled = $True
        $btnScan.IsEnabled = $True
    }
    Catch {
        Write-Log "Connection failed: $($_.Exception.Message)" -Type "ERROR"
        [System.Windows.MessageBox]::Show("Failed to connect.`n`n$($_.Exception.Message)")
    }
}

Function Write-Log {
    Param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $window.Dispatcher.Invoke({
        $txtLog.AppendText("[$timestamp] [$Type] $Message`r`n")
        $txtLog.ScrollToEnd()
    })
}

Function Format-Size {
    Param([long]$Bytes)
    If ($Bytes -gt 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) }
    ElseIf ($Bytes -gt 1MB) { "{0:N2} MB" -f ($Bytes / 1MB) }
    ElseIf ($Bytes -gt 1KB) { "{0:N2} KB" -f ($Bytes / 1KB) }
    Else { "$Bytes Bytes" }
}

# --- Connection ---
$btnConnect.Add_Click({
    Start-Connection
})

# --- Scan ---
$btnScan.Add_Click({
    If (-not $cmbLibrary.SelectedItem) { [System.Windows.MessageBox]::Show("Please select a library."); return }

    $lib = $cmbLibrary.SelectedItem
    $folderPath = $txtFolderPath.Text.Trim()

    Write-Log "Scanning '$($lib.Title)'..."

    $global:ItemsToProcess.Clear()
    $folderCount = 0
    $fileCount = 0
    $totalSize = 0

    $progressBar.IsIndeterminate = $True

    Try {
        $items = Get-PnPListItem -List $lib -FolderServerRelativeUrl $folderPath -Recursive -PageSize 1000 -Includes "HasUniqueRoleAssignments","FileLeafRef","FileRef","File_x0020_Size","FileSystemObjectType"

        ForEach ($item in $items) {
            $global:ItemsToProcess.Add($item)
            If ($item.FileSystemObjectType -eq "Folder") {
                $folderCount++
            } Else {
                $fileCount++
                $totalSize += [long]$item["File_x0020_Size"]
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        $lblFolders.Text = $folderCount
        $lblFiles.Text = $fileCount
        $lblSize.Text = Format-Size $totalSize

        $estSecs = $global:ItemsToProcess.Count * 0.5
        $lblEstTime.Text = "$([Math]::Round($estSecs, 0))s"

        Write-Log "Scan complete. Found $($global:ItemsToProcess.Count) items." -Type "SUCCESS"
        $btnExecute.IsEnabled = $True
    }
    Catch {
        Write-Log "Scan error: $($_.Exception.Message)" -Type "ERROR"
    }
    $progressBar.IsIndeterminate = $False
})

# --- Execution ---
$btnExecute.Add_Click({
    If ($global:ItemsToProcess.Count -eq 0) { return }

    $isDryRun = $chkDryRun.IsChecked
    $modeText = If ($isDryRun) { "DRY RUN" } Else { "REAL EXECUTION" }

    $msg = "Start reset for $($global:ItemsToProcess.Count) items in $modeText mode?"
    $confirm = [System.Windows.MessageBox]::Show($msg, "Confirm Action", "YesNo", "Question")
    If ($confirm -ne "Yes") { return }

    $btnScan.IsEnabled = $False
    $btnExecute.IsEnabled = $False
    $progressBar.Value = 0
    $progressBar.Maximum = $global:ItemsToProcess.Count

    Write-Log "Starting reset ($modeText)..."

    $counter = 0
    ForEach ($item in $global:ItemsToProcess) {
        $counter++
        $itemName = $item["FileLeafRef"]
        $itemUrl = $item["FileRef"]

        Try {
            $hasUnique = $item.HasUniqueRoleAssignments

            If ($hasUnique) {
                If (-not $isDryRun) {
                    Write-Log "Resetting permissions: $($itemName)"
                    $item.ResetRoleInheritance()
                    Invoke-PnPQuery -RetryCount 10
                } Else {
                    Write-Log "[DRY RUN] Would reset permissions: $($itemName)"
                }
            }
        }
        Catch {
            Write-Log "Error on $($itemName): $($_.Exception.Message)" -Type "ERROR"
            If ($_.Exception.Message -like "*429*" -or $_.Exception.Message -like "*503*") {
                Write-Log "Throttling detected. Waiting 10 seconds..." -Type "WARNING"
                Start-Sleep -Seconds 10
            }
        }

        $progressBar.Value = $counter
        [System.Windows.Forms.Application]::DoEvents()
    }

    Write-Log "Execution finished." -Type "SUCCESS"
    $btnScan.IsEnabled = $True
    $btnExecute.IsEnabled = $True
    [System.Windows.MessageBox]::Show("Finished!")
})

# --- Show Window ---
# If window was shown via Auto-Connect, we can't call ShowDialog.
# Instead, we'll use ShowDialog by default, but if we need to auto-start,
# we'll use a 'ContentRendered' event to trigger it.

$window.Add_ContentRendered({
    If ($PSBoundParameters.ContainsKey('SiteURL')) {
        Start-Connection
    }
})

$window.ShowDialog() | Out-Null
