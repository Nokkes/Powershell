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

### Add Search boxes ####
$pathSearchBox = New-Object System.Windows.Forms.TextBox
$pathSearchBox.Size = New-Object System.Drawing.Size(300,25)
$pathSearchBox.Location = New-Object System.Drawing.Size(10,25)
$pathSearchBox.Text = ""
$pathSearchBox.add_click({[void] $pathSearchBox.Clear()})
$MainForm.Controls.Add($pathSearchBox)

$extSearchBox = New-Object System.Windows.Forms.TextBox
$extSearchBox.Size = New-Object System.Drawing.Size(75,25)
$extSearchBox.Location = New-Object System.Drawing.Size(10,70)
$extSearchBox.Text = "*.*"
$extSearchBox.add_click({[void] $extSearchBox.Clear()})
$extSearchBox_FocusLost= {
    if ($extSearchBox.Text -eq ""){
        $extSearchBox.Text = "*.*"
    }
}
$extSearchBox.add_Leave($extSearchBox_FocusLost)
$MainForm.Controls.Add($extSearchBox)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePickerFormat = New-Object System.Windows.Forms.DateTimePickerFormat
$datePickerFormat = 
$datePicker.Format = "Custom"
$datePicker.CustomFormat = "dd/MM/yyyy HH:mm:ss"
$datePicker.Size = New-Object System.Drawing.Size(200,50)
$datePicker.Location = New-Object System.Drawing.Size(110,70)

$MainForm.Controls.Add($datePicker)


### Add Labels ###
$pathText = New-Object System.Windows.Forms.Label
$pathText.Size = New-Object System.Drawing.Size(75,25)
$pathText.Location = New-Object System.Drawing.Size(10,10)
$pathText.Text = "Path"
$MainForm.Controls.Add($pathText)

$extText = New-Object System.Windows.Forms.Label
$extText.Size = New-Object System.Drawing.Size(75,25)
$extText.Location = New-Object System.Drawing.Size(10,55)
$extText.Text = "File Extension"
$MainForm.Controls.Add($extText)

$dateText = New-Object System.Windows.Forms.Label
$dateText.Size = New-Object System.Drawing.Size(75,25)
$dateText.Location = New-Object System.Drawing.Size(110,55)
$dateText.Text = "From Date"
$MainForm.Controls.Add($dateText)

### Add Dgv ###
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Size = New-Object System.Drawing.Size(750,200)
$dgv.Location = New-Object System.Drawing.Size(10,140)
$MainForm.Controls.Add($dgv)

### Add Buttons ###
$Search = New-Object System.Windows.Forms.Button
$Search.Size = New-Object System.Drawing.Size(75,25)
$Search.Location = New-Object System.Drawing.Size(10,100)
$Search.Text = "Search"
$MainForm.Controls.Add($Search)
$Search.add_click({ 
    $dt = [datetime]::ParseExact($datePicker.Text, $datePicker.CustomFormat, $null)
    $result = gci -Recurse -Force -ea SilentlyContinue -Path $pathSearchBox.Text | ? {$_.LastWriteTime -gt $dt -and $_.Extension -eq $extSearchBox.Text} | Select-Object Name, FullName, LastWriteTime, CreationTime, Extension, Directory
    
    if ($result.Count -eq 0){
        [System.Windows.Forms.MessageBox]::Show("No Results Found", "Message")
    }
    else {
        $list = New-Object System.Collections.ArrayList
        $list.AddRange($result)
        $dgv.DataSource = $list
        $dgv.AutoResizeColumns()
    }
})

### Activate Form ###
$MainForm.Add_Shown({$MainForm.Activate()})
[void] $MainForm.ShowDialog()