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

Function Add-PresetControl ($top, $left, $width, $height) {
	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Width = $width - 45
	$textBox.Height = $height
	$textBox.Location = New-Object System.Drawing.Point($left, $top)
	$form.Controls.Add($textBox)
	
	$button = New-Object System.Windows.Forms.Button
	$button.Text = 'Add'
	$button.Width = 40
	$button.Height = $height - 5
	$buttonLeft = $textBox.Left + $textBox.Width + 5
	$button.Location = New-Object System.Drawing.Point($buttonLeft, $top)
	$form.Controls.Add($button)
	
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
	
	#calculate the amount of controls that can be fit into the form's height
	$numOfControlsHeight = (($form.Height - $top - 40) / $controlHeight)
	$numOfControlsHeight = [System.Math]::Floor($numOfControlsHeight)
	
	#set the number of controls for the given width
	$numOfControlsWidth = 2
	
	#calculate the width that each button and textbox pair has to work with
	$controlWidth = [System.Math]::Floor($picBox.Width / $numOfControlsWidth)
	
	$topOffset = 0
	$leftOffset = 0
	$controlCount = 0
	For ($i = 0; $i -lt $numOfControlsWidth; $i++) {
		$controlLeft = $left + $leftOffset
		For ($j = 0; $j -lt $numOfControlsHeight; $j++) {
			$controlTop = $top + $topOffset
			Add-PresetControl $controlTop $controlLeft $controlWidth $controlHeight
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

Function Populate-FlowLayoutPanel ($inputDir, $flowLayoutPanel) {
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
		
		Set-ButtonBorder $button
		
		$button.Add_Click({
			Prompt-Save $form
			
			# Update the picture box and text box with the selected image and text file
			$Script:selectedImageFile = $this.Tag
			$selectedImageFile = $this.Tag
			$Script:currentImagePath = '{0}\{1}' -f $InputDir, $selectedImageFile
			$picBox.Image = [System.Drawing.Image]::FromFile($Script:currentImagePath)
			$selectedTextFile = '{0}\{1}' -f $InputDir, [System.IO.Path]::ChangeExtension($selectedImageFile, '.txt')
			
			# create the text file if it DNE
			If (-not (Test-Path -LiteralPath $selectedTextFile)) {
				Set-Content -LiteralPath $selectedTextFile -Value ''
			}
			
			$Script:currentTextPath = $selectedTextFile
			$Script:txtBox.Text = Get-Content -LiteralPath $selectedTextFile
			$Script:txtBox.SelectionStart = $txtBox.Text.Length
			$fileNameTxtBox.Text = $Script:selectedImageFile
			
			If ($this.FlatAppearance.BorderColor -eq [System.Drawing.Color]::Green) {
				$MarkAsEditedButton.Text = "Unmark as Edited"
			}
			Else {
				$MarkAsEditedButton.Text = "Mark as Edited"
			}
		})
		
		$button.Add_KeyPress({
			param($sender, $eventArgs)
			if ($eventArgs.KeyChar -eq 'M') {
				$MarkAsEditedButton.PerformClick()
			}
		})
	}
}

Function Set-ButtonBorder ($button) {
	$imageFile = $button.tag
	# Set the border colour of the button
	If ((Get-SectionLines 'EditedFiles' $Script:cfgPath -ErrorAction SilentlyContinue) -contains $imageFile) {
		$button.FlatAppearance.BorderSize = 5
		$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Green
	}
	Else {
		$button.FlatAppearance.BorderSize = 1
		$button.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
	}
}

Function Open-Folder ($form) {
	If (-not $form) {
		#try and get the last folder if it exists
		$savedFolderPath = Get-SectionLines 'CapConfig' $Script:capCfgPath | Where-Object {$_ -Match '^FolderPath=.+'}
		If ($savedFolderPath) {
			$savedFolderPath = $savedFolderPath.Split('=')[-1]
			If (Test-Path -LiteralPath $savedFolderPath) {
				$imageFiles = Get-ChildItem -LiteralPath $savedFolderPath -Include $Script:imageFormats
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
		$imageFiles = Get-ChildItem -LiteralPath $folderBrowserDialog.SelectedPath -Include $Script:imageFormats
		If ($imageFiles) {
			Init-Form $folderBrowserDialog.SelectedPath $form $ImageFiles
			$folderBrowserDialog.Dispose()
		}
		Else {
			Open-Folder $form
		}
	}
	Else {
		$folderBrowserDialog.Dispose()
		If (-not $form) {exit}
	}
}

Function Init-Form ($InputDir, $form, $ImageFiles) {
	Prompt-Save $form
	Write-Config $form
	
	If (-not $form)	{
		$Script:presetControls = [ordered]@{}
		$Script:txtBox = $null
	}
	
	$Script:InputDir = Resolve-Path -LiteralPath $InputDir
	$Script:InputDirName = Split-Path $InputDir -Leaf
	$Script:cfgPath = '{0}\{1}{2}' -f $InputDir, $InputDirName, '.cfg'
	
	$Script:imageFiles = $ImageFiles
	$Script:selectedImageFile = $null
	$Script:currentImagePath = $null
	$Script:currentTextPath = $null
	
	# make a cfg file if it dne
	Create-Cfg
	
	Populate-Form $form
}

Function Populate-Form ($form) {
	if (-not $form) {return}
	
	#populate the flow layout panel
	Populate-FlowLayoutPanel $InputDir $flowLayoutPanel

	#add all of the preset controls
	Populate-PresetControls
	
	# select the first control in the flow layout panel
	$control = $flowLayoutPanel.Controls[0]
	$control.Select()
	$control.PerformClick()
}

Function Write-Config ($form) {
	if (-not $form) {return}
	
	$cfgContent = New-Object System.Collections.Generic.List[string]
	
	#set the edited files section of the config file
	$cfgContent.Add('[EditedFiles]')
	ForEach ($button in $flowLayoutPanel.Controls) {
		#if the button is green, we add the file to the edited list
		If ($button.FlatAppearance.BorderColor -eq [System.Drawing.Color]::Green) {
			$cfgContent.Add($button.tag)
		}
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
