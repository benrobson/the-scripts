# 
# Computer and Application Setup Script
# AKA TechSetup
# 
# Docs: 


# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a logging form
$loggingForm = New-Object System.Windows.Forms.Form
$loggingForm.Text = "Script Log"
$loggingForm.Size = New-Object System.Drawing.Size(600, 400)

# Create a RichTextBox for logging
$logTextBox = New-Object System.Windows.Forms.RichTextBox
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Both"
$logTextBox.WordWrap = $false
$logTextBox.Dock = "Fill"

# Add the RichTextBox to the form
$loggingForm.Controls.Add($logTextBox)

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )

    # Append the message to the RichTextBox
    $logTextBox.AppendText("$message`r`n")

    # Scroll to the end of the text
    $logTextBox.ScrollToCaret()
}

function Show-LogScreen {
    # Create a new Windows Form for the welcome screen
    $welcomeForm = New-Object System.Windows.Forms.Form
    $welcomeForm.Text = "Welcome to Windows Update Script"
    $welcomeForm.Size = New-Object System.Drawing.Size(400, 200)

    # Add a label with a welcome message
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Welcome to the Windows Update Script! Click 'OK' to start."
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)

    # Add an OK button to close the form
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(150, 100)
    $okButton.Add_Click({ $welcomeForm.Close() })

    # Add controls to the form
    $welcomeForm.Controls.Add($label)
    $welcomeForm.Controls.Add($okButton)

    # Activate and show the form
    $welcomeForm.Add_Shown({ $welcomeForm.Activate() })
    $welcomeForm.ShowDialog()
}

# Show the log screen
Show-LogScreen

# Administration Verification
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Log-Message "Script terminated: Not running as administrator."
    $loggingForm.ShowDialog()
    exit
}

# Set the Execution Policy
Log-Message "Execution policy set to RemoteSigned."
Set-ExecutionPolicy RemoteSigned -Scope Process -Force

# Windows Update
# Check and list regular updates
Write-Host "Checking and listing regular updates..."
$securityUpdates = Get-WindowsUpdate -MicrosoftUpdate -Category "Security Updates", "Critical Updates"
$securityUpdates | Format-Table -Property Title, KB, Description -AutoSize

# Check and list optional updates
Write-Host "Checking and listing optional updates..."
$optionalUpdates = Get-WindowsUpdate -MicrosoftUpdate -Category "Optional"
$optionalUpdates | Format-Table -Property Title, KB, Description -AutoSize

# Force install all updates
Write-Host "Installing updates..."
$securityUpdates | Install-WindowsUpdate -AcceptAll
$optionalUpdates | Install-WindowsUpdate -AcceptAll

Write-Host "Updates installation complete."
Log-Message "Updates installation complete."

# Chocolatey
Log-Message "Installing Chocolatey..."
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

# Wait for Chocolatey to install (optional, depending on your system)
Start-Sleep -Seconds 5

Log-Message "Installing TeamViewer via Chocolatey..."
choco install teamviewer -y --x64

Log-Message "Installing Adobe Reader x64 via Chocolatey..."
choco install adobereader -y -params '"/DesktopIcon"' --x64

Log-Message "Installing Google Chrome x64 via Chocolatey..."
choco install googlechrome -y --x64 --ignore-checksums

# Show the logging form
$loggingForm.ShowDialog()

