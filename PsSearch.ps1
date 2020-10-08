# Add Assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

### Add Forms ###
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "PsSearch"
$MainForm.Size = New-Object System.Drawing.Size(800,400)
$MainForm.KeyPreview = $True
$MainForm.FormBorderStyle = "1"
$MainForm.MaximizeBox = $false
$MainForm.StartPosition = "CenterScreen"
$Icon = [System.Drawing.SystemIcons]::Question
$MainForm.Icon = $Icon

### Add Text boxes ####
$textboxSearch = New-Object System.Windows.Forms.TextBox
$textboxSearch.Size = New-Object System.Drawing.Size(50,25)
$textboxSearch.Location = New-Object System.Drawing.Size(400,25)
$textboxSearch.Text = ""

$textboxPath = New-Object System.Windows.Forms.TextBox
$textboxPath.Size = New-Object System.Drawing.Size(300,25)
$textboxPath.Location = New-Object System.Drawing.Size(10,25)
$textboxPath.Text = ""
$fbdPath = New-Object System.Windows.Forms.FolderBrowserDialog
$textboxPath.add_click({$fbdPath.ShowDialog();$textboxPath.Text = $fbdPath.SelectedPath})

$textboxExtension = New-Object System.Windows.Forms.TextBox
$textboxExtension.Size = New-Object System.Drawing.Size(75,25)
$textboxExtension.Location = New-Object System.Drawing.Size(10,70)
$textboxExtension.Text = "*"
#$textboxExtension.add_click({[void] $textboxExtension.Clear()})
$textboxExtension_FocusLost = {
    if ($textboxExtension.Text -eq ""){
        $textboxExtension.Text = "*"
    }
}
$textboxExtension.add_Leave($textboxExtension_FocusLost)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePickerFormat = New-Object System.Windows.Forms.DateTimePickerFormat
$datePickerFormat = 
$datePicker.Format = "Custom"
$datePicker.CustomFormat = "dd/MM/yyyy HH:mm:ss"
$datePicker.Size = New-Object System.Drawing.Size(200,50)
$datePicker.Location = New-Object System.Drawing.Size(170,70)

### Add Labels ###
$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Size = New-Object System.Drawing.Size(75,25)
$lblPath.Location = New-Object System.Drawing.Size(10,10)
$lblPath.Text = "Path"

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Size = New-Object System.Drawing.Size(150,25)
$lblSearch.Location = New-Object System.Drawing.Size(400,10)
$lblSearch.Text = "Search String"

$lblExt = New-Object System.Windows.Forms.Label
$lblExt.Size = New-Object System.Drawing.Size(150,25)
$lblExt.Location = New-Object System.Drawing.Size(10,55)
$lblExt.Text = "File Extension (e.g. .txt)"

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Size = New-Object System.Drawing.Size(75,25)
$lblDate.Location = New-Object System.Drawing.Size(170,55)
$lblDate.Text = "From Date"

### Add Dgv ###
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Size = New-Object System.Drawing.Size(750,200)
$dgv.Location = New-Object System.Drawing.Size(10,140)

### Add ProgressBar ###
$pb = New-Object System.Windows.Forms.ProgressBar
$pb.Location = New-Object System.Drawing.Size(100,100)
$pb.Size = New-Object System.Drawing.Size(350,25)

### Add Buttons ###
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Size = New-Object System.Drawing.Size(75,25)
$btnSearch.Location = New-Object System.Drawing.Size(10,100)
$btnSearch.Text = "Search"

$btnSearch.add_click({ 
    $dt = [datetime]::ParseExact($datePicker.Text, $datePicker.CustomFormat, $null)
    $result = gci -Path $textboxPath.Text -Filter "*$($textboxSearch.Text)*" -Recurse -Force -ea SilentlyContinue  | ? {$_.LastWriteTime -gt $dt -and $_.Extension -like $textboxExtension.Text} | Select-Object Name, FullName, LastWriteTime, CreationTime, Extension, Directory
    
    if ($result.Count -le 0){
        [System.Windows.Forms.MessageBox]::Show("No Results Found", "Message from $($MainForm.Text)")
    }
    else {
        $list = New-Object System.Collections.ArrayList
        $list.AddRange($result)
        $dgv.DataSource = $list
        $dgv.AutoResizeColumns()
    }
})



### Add Controls to mainform ###
$MainForm.Controls.Add($textboxSearch)
$MainForm.Controls.Add($textboxPath)
$MainForm.Controls.Add($textboxExtension)
$MainForm.Controls.Add($datePicker)
$MainForm.Controls.Add($lblSearch)
$MainForm.Controls.Add($lblPath)
$MainForm.Controls.Add($lblExt)
$MainForm.Controls.Add($lblDate)
$MainForm.Controls.Add($dgv)
$MainForm.Controls.Add($btnSearch)
$MainForm.Controls.Add($pb)


### Activate Form ###
$MainForm.Add_Shown({$MainForm.Activate()})
[void] $MainForm.ShowDialog()