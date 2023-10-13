# Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Define a function to generate a random number between 10 and 100
Function GenerateRandomNumber {
    return Get-Random -Minimum 10 -Maximum 100
}

# Define a list of words for password generation
$wordList = @(
    "apple", "banana", "cherry", "grape", "orange",
    "lemon", "lime", "kiwi", "mango", "peach"
)

# Function to generate a new password
Function GenerateNewPassword {
    $word1 = $wordList | Get-Random
    $word2 = $wordList | Get-Random

    $number = GenerateRandomNumber
    $symbols = "!", "?", "*"
    $symbol = $symbols | Get-Random

    return "$word1$number$symbol$word2"
}

# Automatically generate initial password
$password = GenerateNewPassword

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Password Generator"
$form.Size = New-Object System.Drawing.Size(300,200) # Adjusted height
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog" # Set form to not resizable

# Create TextBox to display password
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Location = New-Object System.Drawing.Point(10,20)
$textbox.Width = $form.ClientSize.Width - 20 # Adjust width of textbox to fit the form
$textbox.Text = $password
$textbox.ReadOnly = $true

# Create Button to copy password to clipboard
$buttonCopy = New-Object System.Windows.Forms.Button
$buttonCopy.Location = New-Object System.Drawing.Point(20,60) # Adjusted location
$buttonCopy.Size = New-Object System.Drawing.Size(120,40)
$buttonCopy.Text = "Copy to Clipboard"

$buttonCopy.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($password)
    [System.Windows.Forms.MessageBox]::Show("Password copied to clipboard.", "Copy Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Create Button to generate a new password
$buttonNew = New-Object System.Windows.Forms.Button
$buttonNew.Location = New-Object System.Drawing.Point(160,60) # Adjusted location
$buttonNew.Size = New-Object System.Drawing.Size(120,40)
$buttonNew.Text = "Generate New Password"

$buttonNew.Add_Click({
    $password = GenerateNewPassword
    $textbox.Text = $password
})

# Add controls to form
$form.Controls.Add($textbox)
$form.Controls.Add($buttonCopy)
$form.Controls.Add($buttonNew)

# Display the form
$form.ShowDialog()
