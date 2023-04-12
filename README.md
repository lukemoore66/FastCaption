# Fast Caption
An image captioning GUI written in PowerShell

This is useful for captioning images for text-to-image projects like Stable Diffusion.

PowerShell v7 required

This script will generate a GUI that will allow the fast editing of captions via a preset system.
It will display thumbnails of each image, so they can be browsed and edited very easily.
The captions are stored alongside the image in a text file that shares the same name as the image.

Basic Usage:
1. Run the batch file, or the script from a terminal
2. Select a folder with images in it
3. Choose an image to edit from the thumbnail panel
4. Edit it's caption in the text box
5. Add the text presets to your existing caption
6. Save and mark the caption as edited
7. Repeat as needed

Extra features:
- File filtering: Underneath the thumbnail panel, there is a file filtering text box. You can filter the displayed image files by either file name, or caption content.

- Tool box: Underneath the text box containing the current caption's content, there is a tool box that can perform various functions. Select the function you want from the drop down, and the appropriate controls will appear. Click the 'All Files' check box if you want to perform the action on all current files. The following functions are supported:
  - Append: Appends the content of the text box to the current caption.
  - Prepend: Prepends the content of the text box to the current caption.
  - Replace: Replaces all instances of the content in the left text box with the content in the right text box in the current caption.
  - Remove: Removes all instances of the content of the text box in the current caption.
  - Regex Replace: Performs a regex match of the contents of the left text box, and replaces it with the contents of the right text box with the current caption.
  - Set All Text: Sets the current caption's text to the contents of the text box.
  - Clear All Text: Clears the current caption.

- Shortcut keys:
   - 'm'            Pressing 'm' while the thumbnail panel is focussed will mark the current thumbnail as edited.
   - 'shift + m'    Pressing 'shift + m' while the thumbnail panel is focussed with mark all current thumbnails as edited.
   - 'u'            Pressing 'u' while the thumbnail panel is focussed will unmark the current thumbnail as edited.
   - 'shift + u'    Pressing 'shift + u' while the thumbnail panel is focussed with unmark all current thumbnails as edited.

- Configuration:
  - Presets are stored in a cfg file inside the directory you have chosen to edit your images in.
  - The default preset template is stored in the cap.cfg file, you can change this to suit you own needs.

![Screenshot](screenshot.png)
