# import functions and assemblies
. ('{0}\{1}' -f $PSScriptRoot, 'caplib.ps1')

# hide the console window
Hide-Console

# declare global variables
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
$flowLayoutPanel.Add_KeyPress({
	if ($eventArgs.KeyChar -eq 'M' -and $sender.Focused) {
		Write-Host "Doing the thing"
		$button = $sender.GetNextControl($sender.ActiveControl, $true)
		# Set the border colour of the button
		If ((Get-SectionLines 'EditedFiles' $Script:cfgPath -ErrorAction SilentlyContinue) -contains $imagefile) {
			$button.FlatAppearance.BorderSize = 5
			$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
		}
		Else {
			$button.FlatAppearance.BorderSize = 1
			$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
		}
	}
})

# Create a text box to display the text file contents
$txtBox = New-Object System.Windows.Forms.TextBox
$txtBox.Text = $null
$txtBox.Multiline = $true
$txtBox.Width = $form.Width - $picBox.width - 45
$txtBox.Height = ($form.Height - ($picBox.Top + $picBox.Height + 10) - 225) #225 is the height of the saveButton + fileNameTxtBox + clearButton + showButton
$txtBox.Left = $picBox.Left + $picBox.Width + 10
$txtBox.Top = $flowLayoutPanel.Top + $flowLayoutPanel.Height + 10
$form.Controls.Add($txtBox)

# Add a button to save the edited text file
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Width = $txtBox.Width / 2 - 5
$saveButton.Height = 40
$savebutton.Top = $txtBox.Top + $txtBox.Height + 10
$saveButton.Left = $txtBox.Left
$form.Controls.Add($saveButton)
$saveButton.Add_Click({
    $selectedTextFile = '{0}\{1}' -f $InputDir, [System.IO.Path]::ChangeExtension($selectedImageFile, '.txt')
	$cleanedText = Clean-Text $txtBox.Text
	$txtBox.Text = $cleanedText
	Set-Content -LiteralPath $selectedTextFile -Value $cleanedText
})

# Create another text box to display the filename
$fileNameTxtBox = New-Object System.Windows.Forms.TextBox
$fileNameTxtBox.Multiline = $false
$fileNameTxtBox.ReadOnly = $true
$fileNameTxtBox.Width = $txtBox.Width
$fileNameTxtBox.Height = $fileNameTxtBox.PreferredHeight
$fileNameTxtBox.Left = $txtBox.Left
$fileNameTxtBox.Top = $saveButton.Top + $saveButton.Height + 10
$fileNameTxtBox.MultiLine = $true
$fileNameTxtBox.AcceptsReturn = $true
$fileNameTxtBox.Text = $null
$form.Controls.Add($fileNameTxtBox)

# Create a clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = 'Clear Text'
$clearButton.Width = ($txtBox.Width / 2) - 5
$clearButton.Height = 40
$clearButton.Left = $picBox.Left + $picBox.Width + 10
$clearButton.Top = $fileNameTxtBox.Top + $fileNameTxtBox.PreferredHeight + 5
$form.Controls.Add($clearButton)
$clearButton.Add_Click({
	$Script:txtBox.Text = ''
})

# Add a "Show" button
$undoButton = New-Object System.Windows.Forms.Button
$undoButton.Text = "Undo"
$undoButton.Width = ($fileNameTxtBox.Width / 2)
$undoButton.Height = $saveButton.Height
$undoButton.Top  = $clearButton.Top
$undoButton.Left = $clearButton.Left + $clearButton.Width + 5
$undoButton.Enabled = $true
$Form.Controls.Add($undoButton)
$undoButton.Add_Click({
	$Script:txtBox.Undo()
})

# Add a "Show in explorer" button
$showButton = New-Object System.Windows.Forms.Button
$showButton.Text = "Show In Explorer"
$showButton.Width = $fileNameTxtBox.Width
$showButton.Height = $clearButton.Height
$showButton.Top  = $clearButton.Top + $clearButton.Height + 15
$showButton.Left = $clearButton.Left
$showButton.Enabled = $true
$Form.Controls.Add($showButton)
$showButton.Add_Click({
	$textfile = Resolve-Path -LiteralPath ('{0}\{1}{2}' -f $InputDir, [System.IO.Path]::GetFileNameWithoutExtension($Script:selectedImageFile)), '.txt'
	Start-Process -FilePath "explorer.exe" -ArgumentList "/select, $textfile"
})

# Add a "Mark as Edited" button
$MarkAsEditedButton = New-Object System.Windows.Forms.Button
$MarkAsEditedButton.Text = "Mark as Edited"
$MarkAsEditedButton.Width = $txtBox.Width - $saveButton.Width - 60
$MarkAsEditedButton.Height = $saveButton.Height
$MarkAsEditedButton.Top = $saveButton.Top
$MarkAsEditedButton.Left = $saveButton.Left + $saveButton.Width + 5
$MarkAsEditedButton.Enabled = $true
$Form.Controls.Add($MarkAsEditedButton)
$MarkAsEditedButton.Add_Click({
	Prompt-Save $form
	
	foreach ($button in $flowLayoutPanel.Controls) {
		if ($button.Tag -ne $Script:selectedImageFile) {
			continue
		}
		
		#if the button is green, we remove the file from the edited list
		If ($button.FlatAppearance.BorderColor -eq [System.Drawing.Color]::Green) {			
			$button.FlatAppearance.BorderSize = 1
			$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
			$MarkAsEditedButton.Text = "Mark as Edited"
		}
		#otherwise we add it to the edited list
		Else {			
			$button.FlatAppearance.BorderSize = 5
			$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
			$MarkAsEditedButton.Text = "Unmark as Edited"
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
$prevButton.Left = $MarkAsEditedButton.Left + $MarkAsEditedButton.Width + 5
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

$form.Add_Shown({
	Populate-Form $form
})

$form.Add_FormClosing({
	Prompt-Save $form
	
	Write-Config $form
	
	Write-CapCfg 'FolderPath' $Script:InputDir
	
	#release all handles to image / text files
	$flowLayoutPanel.Controls | % {$_.Dispose()}
	$flowLayoutPanel.Controls.Clear()
	$flowLayoutPanel.Dispose()
	$picBox.Dispose()
	$form.Dispose()
	$Script:txtBox.Dispose()
})

$form.ShowDialog() | Out-Null
