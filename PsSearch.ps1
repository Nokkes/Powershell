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

### Add Search boxes ####
$strSearchBox = New-Object System.Windows.Forms.TextBox
$strSearchBox.Size = New-Object System.Drawing.Size(50,25)
$strSearchBox.Location = New-Object System.Drawing.Size(400,25)
$strSearchBox.Text = ""

$pathSearchBox = New-Object System.Windows.Forms.TextBox
$pathSearchBox.Size = New-Object System.Drawing.Size(300,25)
$pathSearchBox.Location = New-Object System.Drawing.Size(10,25)
$pathSearchBox.Text = ""
$fdSearchBox = New-Object System.Windows.Forms.FolderBrowserDialog
$pathSearchBox.add_click({$fdSearchBox.ShowDialog();$pathSearchBox.Text = $fdSearchBox.SelectedPath})



$extSearchBox = New-Object System.Windows.Forms.TextBox
$extSearchBox.Size = New-Object System.Drawing.Size(75,25)
$extSearchBox.Location = New-Object System.Drawing.Size(10,70)
$extSearchBox.Text = "*"
#$extSearchBox.add_click({[void] $extSearchBox.Clear()})
$extSearchBox_FocusLost = {
    if ($extSearchBox.Text -eq ""){
        $extSearchBox.Text = "*"
    }
}
$extSearchBox.add_Leave($extSearchBox_FocusLost)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePickerFormat = New-Object System.Windows.Forms.DateTimePickerFormat
$datePickerFormat = 
$datePicker.Format = "Custom"
$datePicker.CustomFormat = "dd/MM/yyyy HH:mm:ss"
$datePicker.Size = New-Object System.Drawing.Size(200,50)
$datePicker.Location = New-Object System.Drawing.Size(170,70)

### Add Labels ###
$pathText = New-Object System.Windows.Forms.Label
$pathText.Size = New-Object System.Drawing.Size(75,25)
$pathText.Location = New-Object System.Drawing.Size(10,10)
$pathText.Text = "Path"

$strText = New-Object System.Windows.Forms.Label
$strText.Size = New-Object System.Drawing.Size(150,25)
$strText.Location = New-Object System.Drawing.Size(400,10)
$strText.Text = "Search String"

$extText = New-Object System.Windows.Forms.Label
$extText.Size = New-Object System.Drawing.Size(150,25)
$extText.Location = New-Object System.Drawing.Size(10,55)
$extText.Text = "File Extension (e.g. .txt)"

$dateText = New-Object System.Windows.Forms.Label
$dateText.Size = New-Object System.Drawing.Size(75,25)
$dateText.Location = New-Object System.Drawing.Size(170,55)
$dateText.Text = "From Date"


### Add Dgv ###
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Size = New-Object System.Drawing.Size(750,200)
$dgv.Location = New-Object System.Drawing.Size(10,140)


### Add Buttons ###
$Search = New-Object System.Windows.Forms.Button
$Search.Size = New-Object System.Drawing.Size(75,25)
$Search.Location = New-Object System.Drawing.Size(10,100)
$Search.Text = "Search"

$Search.add_click({ 
    $dt = [datetime]::ParseExact($datePicker.Text, $datePicker.CustomFormat, $null)
    $result = gci -Path $pathSearchBox.Text -Filter "*$($strSearchBox.Text)*" -Recurse -Force -ea SilentlyContinue  | ? {$_.LastWriteTime -gt $dt -and $_.Extension -like $extSearchBox.Text} | Select-Object Name, FullName, LastWriteTime, CreationTime, Extension, Directory
    
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

### Add Controls ###
$MainForm.Controls.Add($strSearchBox)
$MainForm.Controls.Add($strText)
$MainForm.Controls.Add($pathSearchBox)
$MainForm.Controls.Add($extSearchBox)
$MainForm.Controls.Add($datePicker)
$MainForm.Controls.Add($pathText)
$MainForm.Controls.Add($extText)
$MainForm.Controls.Add($dateText)
$MainForm.Controls.Add($dgv)
$MainForm.Controls.Add($Search)


### Activate Form ###
$MainForm.Add_Shown({$MainForm.Activate()})
[void] $MainForm.ShowDialog()