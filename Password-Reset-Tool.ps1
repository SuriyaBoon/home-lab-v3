Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

$logFile = "C:\Scripts\password-reset-log.txt"

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Add-Content $logFile
}

$form = New-Object Windows.Forms.Form
$form.Text = "IT Self-Service - Password Reset Tool"
$form.Size = New-Object Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = "FixedDialog"

$lblHeader = New-Object Windows.Forms.Label
$lblHeader.Text = "🔐 Password Reset Tool"
$lblHeader.Font = New-Object Drawing.Font("Segoe UI", 16, [Drawing.FontStyle]::Bold)
$lblHeader.ForeColor = [System.Drawing.Color]::DarkBlue
$lblHeader.Location = New-Object Drawing.Point(20, 20)
$lblHeader.Size = New-Object Drawing.Size(450, 40)

$lblUser = New-Object Windows.Forms.Label
$lblUser.Text = "Username:"
$lblUser.Location = New-Object Drawing.Point(20, 80)
$lblUser.Size = New-Object Drawing.Size(120, 25)
$lblUser.Font = New-Object Drawing.Font("Segoe UI", 10)

$txtUser = New-Object Windows.Forms.TextBox
$txtUser.Location = New-Object Drawing.Point(150, 78)
$txtUser.Size = New-Object Drawing.Size(300, 25)
$txtUser.Font = New-Object Drawing.Font("Segoe UI", 10)

$lblPass = New-Object Windows.Forms.Label
$lblPass.Text = "New Password:"
$lblPass.Location = New-Object Drawing.Point(20, 120)
$lblPass.Size = New-Object Drawing.Size(120, 25)
$lblPass.Font = New-Object Drawing.Font("Segoe UI", 10)

$txtPass = New-Object Windows.Forms.TextBox
$txtPass.Location = New-Object Drawing.Point(150, 118)
$txtPass.Size = New-Object Drawing.Size(300, 25)
$txtPass.PasswordChar = '*'
$txtPass.Font = New-Object Drawing.Font("Segoe UI", 10)

$lblConfirm = New-Object Windows.Forms.Label
$lblConfirm.Text = "Confirm Password:"
$lblConfirm.Location = New-Object Drawing.Point(20, 160)
$lblConfirm.Size = New-Object Drawing.Size(120, 25)
$lblConfirm.Font = New-Object Drawing.Font("Segoe UI", 10)

$txtConfirm = New-Object Windows.Forms.TextBox
$txtConfirm.Location = New-Object Drawing.Point(150, 158)
$txtConfirm.Size = New-Object Drawing.Size(300, 25)
$txtConfirm.PasswordChar = '*'
$txtConfirm.Font = New-Object Drawing.Font("Segoe UI", 10)

$lblReq = New-Object Windows.Forms.Label
$lblReq.Text = "Requirements: Min 8 chars, Uppercase, Lowercase, Number, Special Char"
$lblReq.Location = New-Object Drawing.Point(20, 195)
$lblReq.Size = New-Object Drawing.Size(450, 20)
$lblReq.Font = New-Object Drawing.Font("Segoe UI", 8, [Drawing.FontStyle]::Italic)
$lblReq.ForeColor = [System.Drawing.Color]::Gray

$lblStatus = New-Object Windows.Forms.Label
$lblStatus.Location = New-Object Drawing.Point(20, 230)
$lblStatus.Size = New-Object Drawing.Size(450, 50)
$lblStatus.Font = New-Object Drawing.Font("Segoe UI", 9)

$btnReset = New-Object Windows.Forms.Button
$btnReset.Text = "🔄 Reset Password"
$btnReset.Location = New-Object Drawing.Point(150, 290)
$btnReset.Size = New-Object Drawing.Size(150, 40)
$btnReset.BackColor = [System.Drawing.Color]::DarkBlue
$btnReset.ForeColor = [System.Drawing.Color]::White
$btnReset.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$btnReset.FlatStyle = "Flat"

$btnClear = New-Object Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object Drawing.Point(310, 290)
$btnClear.Size = New-Object Drawing.Size(80, 40)
$btnClear.Font = New-Object Drawing.Font("Segoe UI", 10)

$btnReset.Add_Click({
    $username = $txtUser.Text.Trim()
    $newPass = $txtPass.Text
    $confirmPass = $txtConfirm.Text

    if ([string]::IsNullOrWhiteSpace($username)) {
        $lblStatus.Text = "❌ Error: Please enter a Username"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }

    if ($newPass -ne $confirmPass) {
        $lblStatus.Text = "❌ Error: Passwords do not match"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }

    if ($newPass.Length -lt 8) {
        $lblStatus.Text = "❌ Error: Password must be at least 8 characters long"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }

    if (-not ($newPass -match '[A-Z]' -and $newPass -match '[a-z]' -and $newPass -match '[0-9]' -and $newPass -match '[^a-zA-Z0-9]')) {
        $lblStatus.Text = "❌ Error: Password does not meet complexity requirements"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }

    try {
        $user = Get-ADUser -Identity $username -ErrorAction Stop

        Set-ADAccountPassword -Identity $username `
            -NewPassword (ConvertTo-SecureString $newPass -AsPlainText -Force) `
            -Reset

        Set-ADUser -Identity $username -ChangePasswordAtLogon $true

        $lblStatus.Text = "✅ Reset password Success! User: $username"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green

        Write-Log "SUCCESS - Password reset for user: $username"

        [System.Windows.Forms.MessageBox]::Show(
            "Password reset successful!`nUser: $username`nUser is required to change password at next login.",
            "Success",
            "OK",
            "Information"
        )

        $txtUser.Clear()
        $txtPass.Clear()
        $txtConfirm.Clear()

    } catch {
        $lblStatus.Text = "❌ Error: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "FAILED - User: $username - Error: $($_.Exception.Message)"
    }
})

$btnClear.Add_Click({
    $txtUser.Clear()
    $txtPass.Clear()
    $txtConfirm.Clear()
    $lblStatus.Text = ""
})

$form.Controls.AddRange(@(
    $lblHeader, $lblUser, $txtUser,
    $lblPass, $txtPass, $lblConfirm, $txtConfirm,
    $lblReq, $lblStatus, $btnReset, $btnClear
))

$form.ShowDialog() | Out-Null