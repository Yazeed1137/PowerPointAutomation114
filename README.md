# PowerPointAutomation
VBA code to help create ppt files, add slides, format text, etc.

This project contains VBA code (.bas files) and examples which let a user easily create a PowerPoint presentation.
Demo: https://youtu.be/7WWRDV0Cxbc

Steps to use it:

1. Create a txt file with paragraphs of text (e.g., song lyrics).
2. Go to the examples folder and download the latest pptm file.
3. Right click on the pptm file, open Properties, and Unblock it.
4. Open it, go to the Developer tab, then open the Visual Basic Editor.
5. Open the "CreatePPT" module and enter the name of the txt file as the value of "g_PPTContentFileName"
6. Update "g_ParagraphsPerSlide". If the song is of a single language, use "1". If the song has translation (telugu and english on the same slide), use "2".
7. Press F5 to run the Macro "MainCreatePPT"
8. Observe that new slides have been created, and text from the local text file is added to each slide.
9. If the slides look good, you can close the pptm file. Or you can change the text or font properties, etc., and generate slides again.
10. A new pptx file is created with the same name as the txt file. You can edit this for further manual modifications.

More info: 
1. Each paragraph (multiple lines) of text will be added to one Slide.
2. A blank line between paragraphs marks the transition to a new Slide.
3. Check each Module for Settings at the top of the file. Set values as per your preferences.
4. Run the procedure "MainCreatePPT".
5. This will read the .txt file specified in Settings, create slides, add text, and format text.
6. You can re-run it with different Settings or .txt files.

Update 4/13/2024:
For step 2 above: The number of paragraphs per Slide can be configured. 
Different font settings can be specified for each paragraph.

Update 5/3/2024:
Generated pptx will be saved as a new file in the same folder as the pptm.
pptx file name will be the same as that of the txt file.

Update 8/9/2024:
Added settings for Text Outline. 
By changing the color and weight of the outline, we can make the text more readable against backgrounds of different colors / shapes.

Update 10/23/2024:
Text box shape is filled with a transparent color, so text is more legible against the background image.


## Clipboard automation macro

A new macro `ClipboardAutomaton.PasteClipboardToSlides` is available for quick slide text population:

- Reads plain text directly from the clipboard.
- Applies text using slide style defaults:
  - Font: `DIN Next LT Arabic`
  - Size: `22`
  - Color: `#535353`
  - Alignment: `Justify Low`
- If the active slide already has text formatting, that font/size/color is reused.
- Limits text to **190 characters per slide**. Overflow text is continued on new duplicated slides with the same formatting and layout.

Run the macro from the VBA editor (F5) while your target presentation is open.

To make this click-based, run `ClipboardAutomaton.AddPasteFromClipboardButton` once to add a **Paste from Clipboard** button to the current slide. Clicking that button runs `PasteClipboardToSlides`.
