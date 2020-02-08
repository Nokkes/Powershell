# Adding assemblies
Add-Type -AssemblyName System.Windows.Forms

# Creating Forms Form
$TaskKiller = New-Object System.Windows.Forms.Form
$TaskKiller.Text = "Taskkiller for <ciscojabber.exe> - $env:COMPUTERNAME - $env:USERNAME"
$TaskKiller.BackColor = "#D1C336"
$TaskKiller.TopMost = $true
$TaskKiller.Width = 400
$TaskKiller.Height = 120

# Creating Label
$Label = New-Object system.windows.Forms.Label 
$Label.Width = 600
$Label.Height = 20
$Label.Location = New-Object System.Drawing.Point(10,10)
$Label.Font = "Times New Roman,12"
$Label.Text = "Please enter your Hostname below"

# Creating TextBox
$HostField = New-Object System.Windows.Forms.TextBox
$HostField.Multiline = $false
$HostField.BackColor = "#ffffff"
$HostField.Width = 120
$HostField.Height = 20
$HostField.Location = New-Object System.Drawing.Point(13,38)
$HostField.Font = "Times New Roman,12"

# Creating Button
$KillButton = New-Object System.Windows.Forms.Button
$KillButton.BackColor = "#929195"
$KillButton.Text = "Kill"
$KillButton.Width = 50
$KillButton.Height = 26
$KillButton.Location = New-Object System.Drawing.Point(135,38)
$KillButton.Font = "Times New Roman,12"

# Add Objects to Form
$TaskKiller.Controls.Add($Label)
$TaskKiller.Controls.Add($HostField)
$TaskKiller.Controls.Add($KillButton)

# Action for Kill button
$Button_Click = {
    taskkill.exe /IM ciscojabber.exe /s $HostField.Text
    Start-Sleep 2
}

# Focus on Text Field & perform actions
$HostField.Focus()
$KillButton.Add_Click($Button_Click)

# Starting & Stopping Form
[void]$TaskKiller.ShowDialog()
$TaskKiller.Dispose()


<# Credits: easycalculation.com
            poshgui.com
#>