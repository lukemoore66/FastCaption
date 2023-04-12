# import functions and assemblies
. ('{0}\{1}' -f $PSScriptRoot, 'caplib.ps1')

# hide the console window
Hide-Console

# declare global variables
$Script:toolBoxFunctions = @('Append', 'Prepend', 'Replace', 'Remove', 'Regex Replace', 'Set All Text', 'Clear All Text')
$Script:presetControls = [ordered]@{}
$Script:txtBox = $null
	
$Script:InputDir = $null
$Script:InputDirName = $null
$Script:cfgPath = $null
$Script:capCfgPath = [System.IO.Path]::ChangeExtension($MyInvocation.MyCommand.Path, '.cfg')

$Script:imageFormats = Get-ImageFormats
$Script:imageFiles = $null
$Script:selectedImageFile = $null
$Script:currentImagePath = $null
$Script:currentTextPath = $null
$Script:editedList = New-Object System.Collections.Generic.List[string]

#false = filter by filename, true = filter by content
$Script:filterMode = $false

# open a folder
Open-Folder $null

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Fast Caption"
$form.Width = 1205
$form.Height = [System.Math]::Floor([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height * 0.9)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.KeyPreview = $true

# Add a tool strip menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&File"
$openMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openMenuItem.Text = "&Open Folder"
$openMenuItem.Add_Click({
	Open-Folder $form
})
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "E&xit"
$exitMenuItem.Add_Click({
    $form.Close()
})
$fileMenu.DropDownItems.Add($openMenuItem) | Out-Null
$fileMenu.DropDownItems.Add($exitMenuItem) | Out-Null
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Items.Add($fileMenu) | Out-Null
$form.Controls.Add($menuStrip)

# Create a picture box to display the selected image
$picBox = New-Object System.Windows.Forms.PictureBox
$picBox.SizeMode = 'Zoom'
$picBox.Width = 512
$picBox.Height = 512
$picBox.Top = 25
$picBox.Left = 10
$form.Controls.Add($picBox)

#add all of the preset controls
Add-PresetControls $form $picBox

# create a flow layout panel
$flowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowLayoutPanel.Top = $picBox.Top
$flowLayoutPanel.Left = ($picBox.Left + $picBox.Width + 10)
$flowLayoutPanel.Width = $form.Width - $flowLayoutPanel.Left - 16
$flowLayoutPanel.Height = $picBox.Height - 5
$flowLayoutPanel.FlowDirection = "LeftToRight"
$flowLayoutPanel.WrapContents = $true
$flowLayoutPanel.AutoScroll = $true
$form.Controls.Add($flowLayoutPanel)

# Add a "Show in explorer" button
$showButton = New-Object System.Windows.Forms.Button
$showButton.Text = "Show In Explorer"
$showButton.Width = $form.Width - $picBox.width - 45
$showButton.Height = 40
$showButton.Top  = $form.Height - ($showButton.Height * 2) - 5
$showButton.Left = $picBox.Left + $picBox.Width + 10
$showButton.Enabled = $true
$form.Controls.Add($showButton)
$showButton.Add_Click({
	$textfile = (Resolve-Path -LiteralPath ('{0}\{1}' -f $Script:InputDir, [System.IO.Path]::ChangeExtension($Script:selectedImageFile, '.txt'))).Path
	Start-Process -FilePath "explorer.exe" -ArgumentList "/select, $textfile"
})

# Create a clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = 'Clear Text'
$clearButton.Width = ($showButton.Width / 2) - 5
$clearButton.Height = 40
$clearButton.Left = $picBox.Left + $picBox.Width + 10
$clearButton.Top = $showButton.Top - $clearButton.Height - 5
$form.Controls.Add($clearButton)
$clearButton.Add_Click({
	$Script:txtBox.Text = ''
})

# Add an undo button
$undoButton = New-Object System.Windows.Forms.Button
$undoButton.Text = "Undo"
$undoButton.Width = ($showButton.Width / 2)
$undoButton.Height = 40
$undoButton.Top  = $clearButton.Top
$undoButton.Left = $clearButton.Left + $clearButton.Width + 5
$undoButton.Enabled = $true
$Form.Controls.Add($undoButton)
$undoButton.Add_Click({
	$Script:txtBox.Undo()
})

# Create another text box to display the filename
$fileNameTxtBox = New-Object System.Windows.Forms.TextBox
$fileNameTxtBox.Multiline = $false
$fileNameTxtBox.ReadOnly = $true
$fileNameTxtBox.Width = $showButton.Width
$fileNameTxtBox.Height = $fileNameTxtBox.PreferredHeight
$fileNameTxtBox.Left = $showButton.Left
$fileNameTxtBox.Top = $clearButton.Top - $fileNameTxtBox.Height - 10
$fileNameTxtBox.MultiLine = $true
$fileNameTxtBox.AcceptsReturn = $true
$fileNameTxtBox.Text = $null
$form.Controls.Add($fileNameTxtBox)

# Add a button to save the edited text file
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Width = ($showButton.Width / 2) - 5
$saveButton.Height = 40
$saveButton.Top = $fileNameTxtBox.Top - $saveButton.Height - 10
$saveButton.Left = $showButton.Left
$form.Controls.Add($saveButton)
$saveButton.Add_Click({
    $selectedTextFile = '{0}\{1}' -f $InputDir, [System.IO.Path]::ChangeExtension($selectedImageFile, '.txt')
	$cleanedText = Clean-Text $txtBox.Text
	$txtBox.Text = $cleanedText
	Set-Content -LiteralPath $selectedTextFile -Value $cleanedText
})

# Add a "Mark as Edited" button
$markAsEditedButton = New-Object System.Windows.Forms.Button
$markAsEditedButton.Text = "Mark as Edited"
$markAsEditedButton.Width = $showButton.Width - $saveButton.Width - 60
$markAsEditedButton.Height = $saveButton.Height
$markAsEditedButton.Top = $saveButton.Top
$markAsEditedButton.Left = $saveButton.Left + $saveButton.Width + 5
$markAsEditedButton.Enabled = $true
$Form.Controls.Add($markAsEditedButton)
$markAsEditedButton.Add_Click({
	Prompt-Save $form
	
	ForEach ($button in $flowLayoutPanel.Controls) {
		if ($button.Tag -ne $Script:selectedImageFile) {
			continue
		}
		
		If ($Script:editedList -contains $button.Tag) {
			$Script:editedList.Remove($button.Tag)
			$button.FlatAppearance.BorderSize = 0
			$markAsEditedButton.Text = "Mark as Edited"
		}
		Else {
			$Script:editedList.Add($button.Tag)
			$button.FlatAppearance.BorderSize = 5
			$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
			$markAsEditedButton.Text = "Unmark as Edited"
		}
		
		break
	}
})

# Add a "Previous" button
$prevButton = New-Object System.Windows.Forms.Button
$prevButton.Text = "<"
$prevButton.Width = 20
$prevButton.Height = $saveButton.Height
$prevButton.Top = $saveButton.Top
$prevButton.Left = $markAsEditedButton.Left + $markAsEditedButton.Width + 5
$prevButton.Enabled = $true
$Form.Controls.Add($prevButton)
$prevButton.Add_Click({
	$i = 0
	foreach ($button in $flowLayoutPanel.Controls) {
		if ($button.Tag -eq $Script:selectedImageFile) {
			If (-not $i) {return}
			$control = $flowLayoutPanel.Controls[$i - 1]
			$control.Select()
			$control.PerformClick()
			break
		}
		$i++
	}
})

# Add a "Next" button
$nextButton = New-Object System.Windows.Forms.Button
$nextButton.Text = ">"
$nextButton.Width = 20
$nextButton.Height = $saveButton.Height
$nextButton.Top = $saveButton.Top
$nextButton.Left = $prevButton.Left + $prevButton.Width + 5
$nextButton.Enabled = $true
$Form.Controls.Add($nextButton)
$nextButton.Add_Click({
	$i = 0
	foreach ($button in $flowLayoutPanel.Controls) {
		if ($button.Tag -eq $Script:selectedImageFile) {
			If ($i -ge ($flowLayoutPanel.Controls.Count - 1)) {return}
			$control = $flowLayoutPanel.Controls[$i + 1]
			$control.Select()
			$control.PerformClick()
			break
		}
		$i++
	}
})

# create a search combo box
$searchCombo = New-Object System.Windows.Forms.ComboBox
$searchCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$searchCombo.DropDownHeight = 243
$searchCombo.Width = 200
$searchCombo.Height = 20
$searchCombo.Top = $flowLayoutPanel.Top + $flowLayoutPanel.Height + 5
$searchCombo.Left = $form.Width - $searchCombo.Width - 25
$form.Controls.Add($searchCombo)
$searchCombo.Add_SelectedIndexChanged({
	$wasFocused = $searchCombo.Focused
	
	#the indexes for the selected index for the combo box and the image list have parity
	$control = $flowLayoutPanel.Controls[$searchCombo.SelectedIndex]
	If ($control) {
		$control.Select()
		$control.PerformClick()
		
		If ($wasFocused) {
			$searchCombo.Focus()
		}
	}
})

$searchCombo.Add_KeyDown({
	If (($this.SelectedIndex -eq $this.Items.Count - 1) -and ($_.KeyCode -eq 'Down')) {
		$_.Handled = $true
	}
	
	If (($this.SelectedIndex -eq 0) -and ($_.KeyCode -eq 'Up')) {
		$_.Handled = $true
	}
})

#add a search box
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Width = 200
$searchBox.Height = 20
$searchBox.Top = $searchCombo.Top
$searchBox.Left = $searchCombo.Left - $searchBox.Width - 5
$form.Controls.Add($searchBox)
$searchBox.add_KeyDown({
    if ($_.KeyCode -eq 'Enter') {
		Prompt-Save
				
		$newImageFiles = Get-ImageFiles $Script:InputDir		
		If ($newImageFiles) {
			Init-Form $Script:InputDir $form $newImageFiles
		}
		Else {
			Show-PopUp $searchBox 'No Images Found! Aborting...'
		}
		
		If ($searchCombo.Items.Count -gt 0) {$searchCombo.SelectedIndex = 0}
    }
})

#add radio buttons for search options
$radioLabel = New-Object System.Windows.Forms.Label
$radioLabel.Text = "Filter by:"
$radioLabel.Width = 52
$radioLabel.Height = 20
$radioLabel.Top = $searchCombo.Top
$radioLabel.Left = $picBox.Left + $picBox.Width + 10
$form.Controls.Add($radioLabel)

$radioButton1 = New-Object System.Windows.Forms.RadioButton
$radioButton1.Text = "File Name"
$radioButton1.Width = 77
$radioButton1.Height = 20
$radioButton1.Top = $radioLabel.Top
$radioButton1.Left = $radioLabel.Left + $radioLabel.Width
$form.Controls.Add($radioButton1)
$radioButton1.Add_CheckedChanged({
	If ($this.Checked) {
		$Script:filterMode = $False
	}
	Else {
		$Script:filterMode = $true
	}
	
	Set-FilterMode
})

$radioButton2 = New-Object System.Windows.Forms.RadioButton
$radioButton2.Text = "Caption Content"
$radioButton2.Width = 112
$radioButton2.Height = 20
$radioButton2.Top = $radioLabel.Top
$radioButton2.Left = $radioButton1.Left + $radioButton1.Width
$form.Controls.Add($radioButton2)
$radioButton2.Add_CheckedChanged({
	If ($this.Checked) {
		$Script:filterMode = $true
	}
	Else {
		$Script:filterMode = $false
	}
	
	Set-FilterMode
})

#create a checkbox for all files
$toolCheckbox = New-Object System.Windows.Forms.CheckBox
$toolCheckbox.Text = 'All Files'
$toolCheckbox.Width = 80
$toolCheckbox.Height = 20
$toolCheckbox.Top = $saveButton.Top - $toolCheckbox.Height - 5
$toolCheckbox.Left = $form.Width - $toolCheckbox.Width - 5
$form.Controls.Add($toolCheckbox)

# Create an apply button
$toolButton = New-Object System.Windows.Forms.Button
$toolButton.Text = "Apply"
$toolButton.Width = 45
$toolButton.Height = $toolCheckbox.Height + 2
$toolButton.Top = $toolCheckbox.Top - 2
$toolButton.Left = $toolCheckbox.Left - $toolButton.Width - 5
$toolButton.Enabled = $true
$Form.Controls.Add($toolButton)
$toolButton.Add_Click({
	Handle-ToolBoxButton $toolCombo.SelectedItem $toolCheckbox.Checked
})

#create a tool combo box
$toolCombo = New-Object System.Windows.Forms.ComboBox
$toolCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$toolCombo.DropDownHeight = 243
$toolCombo.Width = 105
$toolCombo.Height = 20
$toolCombo.Top = $toolCheckbox.Top - 2
$toolCombo.Left = $toolButton.Left - $toolCombo.Width - 5
$toolCombo.Items.AddRange($Script:toolboxFunctions)
$form.Controls.Add($toolCombo)
$toolCombo.SelectedIndex = 0
$toolCombo.Add_SelectedIndexChanged({
	Handle-ToolBoxChange $toolCombo.SelectedItem
})

# Create tool text box 1
$toolTextBox1 = New-Object System.Windows.Forms.TextBox
$toolTextBox1.Width = (($toolCombo.Left - $showButton.Left) / 2) - 5
$toolTextBox1.Height = $toolTextBox1.PreferredHeight
$toolTextBox1.Top = $toolCheckbox.Top - 2
$toolTextBox1.Left = $showButton.Left
$form.Controls.Add($toolTextBox1)

# Create tool text box 2
$toolTextBox2 = New-Object System.Windows.Forms.TextBox
$toolTextBox2.Width = $toolTextBox1.Width
$toolTextBox2.Height = $toolTextBox2.PreferredHeight
$toolTextBox2.Top = $toolCheckbox.Top - 2
$toolTextBox2.Left = $showButton.Left + $toolTextBox1.Width + 5
$form.Controls.Add($toolTextBox2)

# Create a text box to display the text file contents
$txtBox = New-Object System.Windows.Forms.TextBox
$txtBox.Text = $null
$txtBox.Multiline = $true
$txtBox.Left = $showButton.Left
$txtBox.Top = $searchCombo.Top + $searchCombo.Height + 5
$txtBox.Width = $showButton.Width
$txtBox.Height = $toolCheckbox.Top - $txtBox.Top - 5
$form.Controls.Add($txtBox)

$form.Add_Shown({
	Populate-Form $form
})

$form.Add_FormClosing({
	Prompt-Save $form
	
	Write-Config $form
	
	Write-CapCfg 'FolderPath' $Script:InputDir
	
	#release all handles to image / text files
	$flowLayoutPanel.Controls | % {$_.Dispose()}
	$form.Controls | % {$_.Dispose()}
	$picBox.Image.Dispose()
	$form.Dispose()
})

$form.ShowDialog() | Out-Null
