# import assemblies
Add-Type -AssemblyName System.Windows.Forms

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

# functions
Function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
}

Function Get-ImageFormats () {
	$formats = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders().FilenameExtension
	$formats = $formats.Replace('*', '')
	$formats = $formats.Split(';')
	$formats = $formats.ToLower()
	$formats = $formats.Replace('.', '*.')
	$formats = $formats.Trim().Split()
	
	return $formats
}

Function Prompt-Save ($form) {
	if (-not $form) {return}
	
	foreach ($button in $flowLayoutPanel.Controls) {
		if ($button.Tag -ne $Script:selectedImageFile) {
			continue
		}
		
		if ((Get-Content -LiteralPath $Script:currentTextPath) -eq $txtBox.Text) {
			break
		}
		
		$result = [System.Windows.Forms.MessageBox]::Show("You have not saved your changes. Do you wish to save them?", "Confirmation", `
		[System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

		if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
			$saveButton.PerformClick()
		}
	}
}

Function Clean-Text ($text) {
	$cleanedText = $text -replace '\r\n', ', '
	$cleanedText = $cleanedText -replace ',+', ', '
	$cleanedText = $cleanedText -replace '\s+', ' '
	$cleanedText = $cleanedText.Trim(' ,').Trim().ToLower()
	
	Return $cleanedText
}

Function Add-PresetControl ($panel, $top, $left, $width, $height) {
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Width = $width - 45
	$textBox.Height = $height
	$textBox.Location = New-Object System.Drawing.Point($left, $top)
	$panel.Controls.Add($textBox)
	
	$button = New-Object System.Windows.Forms.Button
	$button.Text = 'Add'
	$button.Width = 40
	$button.Height = $height - 5
	$buttonLeft = $textBox.Left + $textBox.Width + 5
	$button.Location = New-Object System.Drawing.Point($buttonLeft, $top)
	$panel.Controls.Add($button)
	
	$Script:presetControls[$button] = $textbox
		
	$button.Add_Click({
		param($sender)
		$currentTextBox = $Script:presetControls[$sender]
		If ($currentTextBox.Text) {
			$joinString = ', '
			
			$Script:txtBox.Refresh()	
			$text = $Script:txtBox.Text
			
			$cursorPos = $Script:txtBox.SelectionStart
			
			$startText = $text.SubString(0, $cursorPos)
			$midText = ($Script:presetControls[$sender]).Text
			$endText = $text.SubString($cursorPos)
			
			$cleanStartText = (Clean-Text $startText) + $joinString + (Clean-Text $midtext) + $joinString
			$cleanEndText = (Clean-Text $endText) + $joinString
			$cleanText = $cleanStartText + $cleanEndText
			
			$Script:txtBox.Text = Clean-Text $cleanText

			If ($Script:txtBox.SelectionStart -gt ($Script:txtBox.Text.Length - 1)) {
				$Script:txtBox.Text = Clean-Text $Script:txtBox.Text
				$Script:txtBox.SelectionStart = $Script:txtBox.Text.Length - 1
			}
			Else {	
				$Script:txtBox.SelectionStart = $cleanStartText.Length
			}
		}
	})
}

Function Add-PresetControls ($form, $picBox) {
	#get the top, left, width and height
	$top = $picBox.Top + $picBox.Height + 10
	$left = $picBox.Left
	
	#get the default control height
	$getControlHeight = New-Object System.Windows.Forms.TextBox
	$controlHeight = $getControlHeight.PreferredHeight + 5
	$getControlHeight = $null
	
	#get width and height of the panel
	$totalWidth = $picBox.Width
	$totalHeight = $form.Height - $top - 40
	
	#make a panel to store the controls
	$presetPanel = New-Object System.Windows.Forms.Panel
	$presetPanel.Top = $top
	$presetPanel.Left = $left
	$presetPanel.Width = $totalWidth
	$presetPanel.Height = $totalHeight
	$presetPanel.AutoScroll = $true
	$form.Controls.Add($presetPanel)
	
	#calculate the amount of controls that can be fit into the form's height
	#$numOfControlsHeight = ($totalHeight / $controlHeight)
	#$numOfControlsHeight = [System.Math]::Floor($numOfControlsHeight)
	$numOfControlsHeight = 32
	
	#set the number of controls for the given width
	$numOfControlsWidth = 2
	
	#calculate the width that each button and textbox pair has to work with
	$controlWidth = [System.Math]::Floor($totalWidth / $numOfControlsWidth) - 18
	
	$topOffset = 0
	$leftOffset = 0
	$controlCount = 0
	For ($i = 0; $i -lt $numOfControlsWidth; $i++) {
		$controlLeft = $left + $leftOffset
		For ($j = 0; $j -lt $numOfControlsHeight; $j++) {
			$controlTop = $topOffset
			Add-PresetControl $presetPanel $controlTop $controlLeft $controlWidth $controlHeight
			$topOffset += $controlHeight
			$controlCount++
		}
		
		$topOffset = 0
		$leftOffset += $controlWidth + 5
	}
}

Function Populate-PresetControls () {
	# get preset lines
	$presetLines = @()
	$presetLines = Get-SectionLines 'PresetTexts' $Script:cfgPath
	
	#pad the array as needed
	$paddingAmount = $Script:presetControls.Count - $presetLines.Size
	$presetLines += [string[]]::new($paddingAmount)
	
	$i = 0
	ForEach ($entry in $Script:presetControls.GetEnumerator()) {
		$entry.Value.Text = $presetLines[$i]
		$i++
	}
}

Function Get-SectionLines ($section, $configPath) {
	$cfgContent = Get-Content -LiteralPath $configPath
	
	$sectionLines = @()
    $isInSection = $false

     ForEach ($line in $cfgContent) {
        if ($line -match "^\[($section)\]$") {
            $isInSection = $true
        } elseif ($line -match "^\[.+\]$") {
            $isInSection = $false
        } elseif ($isInSection) {
            $sectionLines += $line
        }
    }
	
	Return $sectionLines
}

Function Create-Cfg () {
	If (-not (Test-Path -LiteralPath $Script:cfgPath)) {		
		$content = @()
		$content += @('[EditedFiles]')
		$content += @('[PresetTexts]')
		$content += Get-SectionLines 'PresetTexts' $Script:capCfgPath
		
		Set-Content -LiteralPath $script:cfgPath -Value $content
	}
}

Function Populate-FlowLayoutPanel () {
	$flowLayoutPanel.Controls.Clear()
	
	ForEach ($imageFile in $Script:imageFiles) {
		# Load the thumbnail image and add it to the flow layout panel
		$button = New-Object System.Windows.Forms.Button
		$button.Width = 100
		$button.Height = 100
		$thumb = [System.Drawing.Image]::FromFile($imageFile.FullName)
		$thumb = $thumb.GetThumbnailImage($button.Width - 8, $button.Height - 8, $null, [System.IntPtr]::Zero)
		$button.FlatStyle = "Flat"
		$button.FlatAppearance.BorderSize = 1
		$button.Image = $thumb
		$button.ImageAlign = "MiddleCenter"
		$button.Tag = $imageFile.Name
		$flowLayoutPanel.Controls.Add($button)
		
		If ($Script:editedList -contains $button.Tag) {
			$button.FlatAppearance.BorderSize = 5
			$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
		}
		Else {
			$button.FlatAppearance.BorderSize = 0
		}
		
		$button.Add_Click({
			Prompt-Save $form
			
			# Update the picture box and text box with the selected image and text file
			$Script:selectedImageFile = $this.Tag
			$selectedImageFile = $this.Tag
			$Script:currentImagePath = '{0}\{1}' -f $Script:InputDir, $selectedImageFile
			
			#dispose of the previous picture box image if there is one
			If ($picBox.Image) {
				$picBox.Image.Dispose()
			}
			
			$picBox.Image = [System.Drawing.Image]::FromFile($Script:currentImagePath)
			$selectedTextFile = '{0}\{1}' -f $Script:InputDir, [System.IO.Path]::ChangeExtension($selectedImageFile, '.txt')
			
			# create the text file if it DNE
			If (-not (Test-Path -LiteralPath $selectedTextFile)) {
				Set-Content -LiteralPath $selectedTextFile -Value ''
			}
			
			$Script:currentTextPath = $selectedTextFile
			$Script:txtBox.Text = Get-Content -LiteralPath $selectedTextFile
			$Script:txtBox.SelectionStart = $txtBox.Text.Length
			$fileNameTxtBox.Text = $Script:selectedImageFile
			
			If ($Script:editedList -contains $this.Tag) {
				$markAsEditedButton.Text = "Unmark As Edited"
			}
			Else {
				$markAsEditedButton.Text = "Mark As Edited"
			}
			
			#update the combo box selection
			$searchCombo.SelectedIndex = $flowLayoutPanel.Controls.IndexOf($this)
		})
		
		$button.Add_KeyPress({
			param($sender, $e)
						
		    If ($e.KeyChar -ceq 'M') {
				ForEach ($button in $flowLayoutPanel.Controls) {
					If ($Script:editedList -notcontains $button.Tag) {
						$Script:editedList.Add($button.Tag)
						$button.FlatAppearance.BorderSize = 5
						$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
					}
				}
				$markAsEditedButton.Text = "Unmark as Edited"
				Return
			}
			
			If ($e.KeyChar -ceq 'U') {
				ForEach ($button in $flowLayoutPanel.Controls) {
					If ($Script:editedList -contains $button.Tag) {
						$Script:editedList.Remove($button.Tag)
						$button.FlatAppearance.BorderSize = 0
						$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
					}
				}
				$markAsEditedButton.Text = "Mark as Edited"
				Return
			}

			If ($e.KeyChar -ceq 'm') {
				If ($Script:editedList -notcontains $this.Tag) {
						$Script:editedList.Add($this.Tag)
						$this.FlatAppearance.BorderSize = 5
						$this.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
				}
				$markAsEditedButton.Text = "Unmark as Edited"
				Return
			}
			
			If ($e.KeyChar -ceq 'u') {
				If ($Script:editedList -contains $this.Tag) {
						$Script:editedList.Remove($this.Tag)
						$this.FlatAppearance.BorderSize = 0
						$this.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
				}
				$markAsEditedButton.Text = "Mark as Edited"
				Return
			}
		})
	}
}

Function Show-PopUp ($control, $message) {
	# set the popup's location and size
	$controlLocation = $control.PointToScreen([System.Drawing.Point]::Empty)
	$popupFormX = [System.Math]::Floor($controlLocation.X - ($controlLocation.X * 0.0025))
	$popupFormY = [System.Math]::Floor($controlLocation.Y - ($controlLocation.Y * 0.015))
	$popupFormWidth = [System.Math]::Floor($control.ClientSize.Width + ($control.ClientSize.Width * 0.075))
	$popupFormHeight = [System.Math]::Floor($control.ClientSize.Height + ($control.ClientSize.Height * 0.05))
	
	# Create a new form with a label control that displays the popup message
	$popupForm = New-Object System.Windows.Forms.Form
	$popupForm.StartPosition = 'Manual'
    $popupForm.Location = New-Object System.Drawing.Point($popupFormX, $popupFormY)
	$popupForm.Width = $popupFormWidth
	$popupForm.Height = $popupFormHeight
	$popupForm.FormBorderStyle = 'None'
	$popupForm.BackColor = [System.Drawing.Color]::WhiteSmoke
	$popupForm.Anchor = 'Top'

	$popupLabel = New-Object System.Windows.Forms.Label
	$popupLabel.Text = $message
	$popupLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
	$popupLabel.TextAlign = 'MiddleCenter'
	$popupLabel.Dock = 'Fill'
	$popupForm.Controls.Add($popupLabel)
	
	# Set a timer to toggle the visibility of the label to make the text flash
	$flashtimer = New-Object System.Windows.Forms.Timer
	$flashtimer.Interval = 500 # Set the interval to half a second
	$flashtimer.add_Tick({
		$popupLabel.Visible = !$popupLabel.Visible # Toggle the visibility of the label
	})
	$flashtimer.Start()

	# Set a timer to close the form after a few seconds
	$timer = New-Object System.Windows.Forms.Timer
	$timer.Interval = 3000
	$timer.add_Tick({
		$flashtimer.Dispose()
		$popupForm.Close()
		$timer.Dispose()
	})
	$timer.Start()

	# Show the form
	$popupForm.ShowDialog()
}

Function Open-Folder ($form) {
	If (-not $form) {
		#try and get the last folder if it exists
		$savedFolderPath = Get-SectionLines 'CapConfig' $Script:capCfgPath | Where-Object {$_ -Match '^FolderPath=.+'}
		If ($savedFolderPath) {
			$savedFolderPath = $savedFolderPath.Split('=')[-1]
			If (Test-Path -LiteralPath $savedFolderPath) {
				$imageFiles = Get-ImageFiles $savedFolderPath
				If ($imageFiles) {
					Init-Form $savedFolderPath $form $ImageFiles
					return
				}
				Else {
					Write-CapCfg 'FolderPath' ''
					Open-Folder $form
					return
				}
			}
		}
	}
	
	$folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
	$folderBrowserDialog.ShowNewFolderButton = $false
	
	$response = $folderBrowserDialog.ShowDialog()
	
	If ($response -eq [System.Windows.Forms.DialogResult]::OK) {
		$imageFiles = Get-ImageFiles $folderBrowserDialog.SelectedPath
		If ($imageFiles) {
			Init-Form $folderBrowserDialog.SelectedPath $form $ImageFiles
		}
		Else {
			Open-Folder $form
		}
	}
	Else {
		If (-not $form) {exit}
	}
	
	End {
		If ($folderBrowserDialog) {$folderBrowserDialog.Dispose()}
	}
}

Function Get-ImageFiles ($folderPath) {
	#always get all the files
	$imageFiles = Get-ChildItem -LiteralPath $folderPath -Include $Script:imageFormats
		
	If ([string]::IsNullOrEmpty($searchBox.Text)) {
		Return $imageFiles
	}
	
	#if we are in the filename filtering mode
	If (-not $Script:filterMode) {
		$imageFiles = $imageFiles | Where-Object {$_.Name -match $searchBox.Text}
		Return $imageFiles
	}
	
	#otherwise we are in the content filtering mode
	$imageFiles = $imageFiles | Where-Object {(Get-Content -LiteralPath $([System.IO.Path]::ChangeExtension([string]$_.FullName, '.txt')) -ErrorAction SilentlyContinue) -match $searchBox.Text}
	Return $imageFiles
}

Function Init-Form ($InputDir, $form, $ImageFiles) {
	$InputDir = (Resolve-Path -LiteralPath $InputDir).Path
	$Script:InputDir = $InputDir
	$Script:InputDirName = Split-Path $InputDir -Leaf
	$Script:cfgPath = '{0}\{1}{2}' -f $InputDir, $InputDirName, '.cfg'
	
	#create a config file if one dne yet
	Create-Cfg
	
	$oldInputDir = $Script:InputDir
	$oldInputDirName = $Script:InputDirName
	$oldcfgPath = $Script:cfgPath
	
	Prompt-Save $form
	
	#only write config if we have changed directory
	If ($InputDir -ne $oldInputDir) {
		Write-Config $form
	}
	
	#if there is no form yet
	If (-not $form)	{
		$Script:presetControls = [ordered]@{}
		$Script:txtBox = $null
		(Get-SectionLines 'EditedFiles' $oldcfgPath) | ForEach-Object {$Script:editedList.Add($_)}
	}
	
	#if we have changed directory
	If ($InputDir -ne $oldInputDir) {
		If (-not $oldcfgPath) {$oldcfgPath = $Script:cfgPath}
		$Script:editedList = New-Object System.Collections.Generic.List[string]
		(Get-SectionLines 'EditedFiles' $oldcfgPath) | ForEach-Object {$Script:editedList.Add($_)}
	}
	
	#clear the search filter if needed
	If (($InputDir -ne $oldInputDir) -and (-not [string]::IsNullOrEmpty($searchBox.Text))) {
		$searchBox.Text = ''
	}
	
	$Script:imageFiles = $ImageFiles
	$Script:selectedImageFile = $null
	$Script:currentImagePath = $null
	$Script:currentTextPath = $null
	
	Populate-Form $form
}

Function Populate-Form ($form) {
	If (-not $form) {Return}
	
	#set and reset the tool combo box selection, this populates the tool box
	$selectedIndex = $toolCombo.SelectedIndex
	$toolCombo.SelectedIndex = ($selectedIndex + 1) % ($toolCombo.Items.Count - 1)
	$toolCombo.SelectedIndex = $selectedIndex
	
	#populate the flow layout panel
	Populate-FlowLayoutPanel
	
	#set the filtering mode
	Set-FilterMode
	
	#populate combo box
	$searchCombo.Items.Clear()
	$searchCombo.Items.AddRange($Script:imageFiles.Name)
	$searchCombo.SelectedIndex = 0

	#add all of the preset controls
	Populate-PresetControls
	
	# select the first control in the flow layout panel
	$control = $flowLayoutPanel.Controls[0]
	$control.Select()
	$control.PerformClick()
}

Function Set-FilterMode() {
	If ($Script:filterMode) {
		$radioButton1.Checked = $False
		$radioButton2.Checked = $True
		Return
	}
	
	$radioButton2.Checked = $False
	$radioButton1.Checked = $True
}

Function Write-Config ($form) {
	if (-not $form) {return}
		
	$cfgContent = New-Object System.Collections.Generic.List[string]
	
	#set the edited files section of the config file
	$cfgContent.Add('[EditedFiles]')
	ForEach ($image in ($Script:editedList | Where-Object {$_})) {
		#add all images in the edited list
		$cfgContent.Add($image)
	} 
	
	#set the preset texts section of the config file
	$cfgContent.Add('[PresetTexts]')
	ForEach($entry in ($Script:presetControls).GetEnumerator()) {
		$cfgContent.Add($entry.Value.Text)
	}
	
	Set-Content -LiteralPath $Script:cfgPath -Value $cfgContent
}

Function Write-CapCfg ($key, $value) {
	If (-not $key) {return}
	
	$entry = '{0}={1}' -f $key, $value
	
	$cfgContent = New-Object System.Collections.Generic.List[string]
	
	#set the config section of the config file
	$cfgContent.Add('[CapConfig]')
	
	$sectionLines = @()
	$sectionLines += Get-SectionLines 'CapConfig' $Script:capCfgPath
	
	$i = 0
	$matchFlag = $false
	ForEach ($line in $sectionLines) {
		If ($line -match "^$key=(.+|$)") {
			$matchFlag = $true
			$sectionLines[$i] = $entry
			break
		}
		
		$i++
	}
	
	If (-not $matchFlag) {
		$sectionLines += $entry
	}
	
	ForEach ($line in $sectionLines) {
		$cfgContent.Add($line)
	}
	
	#set the preset section of the config file
	$cfgContent.Add('[PresetTexts]')
	
	$sectionLines = @()
	$sectionLines += Get-SectionLines 'PresetTexts' $Script:capCfgPath
	ForEach ($line in $sectionLines) {
		$cfgContent.Add($line)
	}
	
	Set-Content -LiteralPath $Script:capCfgPath -Value $cfgContent	
}

Function Handle-ToolBoxChange ($selection) {
		Switch ($selection) {
		'Append' {
			$toolTextBox1.Width = $toolCombo.Left - $showButton.Left - 5
			$toolTextBox1.Enabled = $true
			$toolTextBox2.Visible = $false
		}
		'Prepend' {
			$toolTextBox1.Width = $toolCombo.Left - $showButton.Left - 5
			$toolTextBox1.Enabled = $true
			$toolTextBox2.Visible = $false
		}
		'Replace' {
			$toolTextBox1.Width = (($toolCombo.Left - $showButton.Left) / 2) - 10
			$toolTextBox1.Enabled = $true
			$toolTextBox2.Visible = $true
		}
		'Remove' {
			$toolTextBox1.Width = $toolCombo.Left - $showButton.Left - 5
			$toolTextBox1.Enabled = $true
			$toolTextBox2.Visible = $false
		}
		'Regex Replace' {
			$toolTextBox1.Width = (($toolCombo.Left - $showButton.Left) / 2) - 10
			$toolTextBox1.Enabled = $true
			$toolTextBox2.Visible = $true
		}
		'Set All Text' {
			$toolTextBox1.Width = $toolCombo.Left - $showButton.Left - 5
			$toolTextBox1.Enabled = $true
			$toolTextBox2.Visible = $false
		}
		'Clear All Text' {
			$toolTextBox1.Width = $toolCombo.Left - $showButton.Left - 5
			$toolTextBox1.Enabled = $false
			$toolTextBox2.Visible = $false	
		}
		Default {
			Throw(Write-Error ('Unknown Combo box selection: {0}' -f $_))
		}
	}
}

Function Handle-ToolBoxButton ($selection, $allFiles) {
	If ($allFiles) {
		$result = [System.Windows.Forms.MessageBox]::Show(
			('This will apply the ''{0}'' action to all current files. This cannot be undone. Are you sure you wish to proceed?' -f $selection.ToLower()),
			'Warning',
			[System.Windows.Forms.MessageBoxButtons]::YesNo,
			[System.Windows.Forms.MessageBoxIcon]::Warning
		)
		
		If ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
			Return
		}
		
		ForEach ($file in $Script:imageFiles) {
			$textFile = [System.IO.Path]::ChangeExtension($file.FullName, '.txt')
			If (-not (Test-Path -LiteralPath $textFile)) {
				New-Item -Path $textFile -ItemType File | Out-Null
			}
			
			$content = Get-Content -LiteralPath $textFile
			
			Switch ($selection) {
				'Append' {
					$content += $toolTextBox1.Text
				}
				'Prepend' {
					$content = $toolTextBox1.Text + $content
				}
				'Replace' {
					$content = $content.Replace($toolTextBox1.Text, $toolTextBox2.Text)
				}
				'Remove' {
					$content = $content.Replace($toolTextBox1.Text, '')
				}
				'Regex Replace' {
					$content = $content -replace $toolTextBox1.Text, $toolTextBox2.Text
				}
				'Set All Text' {
					$content = $toolTextBox1.Text
				}
				'Clear All Text' {
					$content  = ''
				}
				Default {
					Throw(Write-Error ('Unknown Combo box selection: {0}' -f $_))
				}
			}
			
			Set-Content -LiteralPath $textFile -Value $content
			$txtBox.Text = Get-Content -LiteralPath $Script:currentTextPath
		}
	}
	Else {
		Switch ($selection) {
			'Append' {
				$txtBox.Text += $toolTextBox1.Text
			}
			'Prepend' {
				$txtBox.Text = $toolTextBox1.Text + $txtBox.Text
			}
			'Replace' {
				$txtBox.Text = $txtBox.Text.Replace($toolTextBox1.Text, $toolTextBox2.Text)
			}
			'Remove' {
				$txtBox.Text = $txtBox.Text.Replace($toolTextBox1.Text, '')
			}
			'Regex Replace' {
				$txtBox.Text = $txtBox.Text -replace $toolTextBox1.Text, $toolTextBox2.Text
			}
			'Set All Text' {
				$txtBox.Text = $toolTextBox1.Text
			}
			'Clear All Text' {
				$txtBox.Text = ''
			}
			Default {
				Throw(Write-Error ('Unknown Combo box selection: {0}' -f $_))
			}
		}
	}
}