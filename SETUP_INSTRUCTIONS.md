# Setup Instructions for Interactive Review Mode

## Overview
The code has been updated with a new **Interactive Preview Mode** that shows each suggestion one-at-a-time with immediate accept/reject/skip options. The old tracked changes workflow is preserved and can be re-enabled by changing one constant.

## What You Need to Do

### 1. Import the New Form into Your VBA Project

1. Open your Word document with the VBA project
2. Press `Alt+F11` to open the VBA Editor
3. Go to **File > Import File...**
4. Navigate to `g:\My Drive\Programing\VBA\`
5. Select **frmSuggestionPreview.frm** and click **Open**
6. The form will be imported (you may see a warning about the .frx file - that's normal)

### 2. Design the Form UI

Since VBA forms don't have a visual designer in text files, you need to manually add controls to `frmSuggestionPreview`:

1. In the VBA Editor, double-click **frmSuggestionPreview** in the Project Explorer
2. If the form opens in code view, press `Shift+F7` to switch to design view
3. Add the following controls (from the Toolbox):

#### **Labels (for headings):**
- **lblProgress** - Top of form
  - Caption: "Suggestion 1 of 10"
  - Font: Bold, 10pt
  - Position: Top center

- **lblActionType** - Below progress
  - Caption: "Action: change"
  - Font: Bold, 9pt

#### **Labels (for field names):**
- **Label1** - Caption: "Context:"
- **Label2** - Caption: "Target:"
- **Label3** - Caption: "Replacement:"
- **Label4** - Caption: "Explanation:"

#### **TextBoxes (multiline, read-only):**
- **txtContext**
  - MultiLine: True
  - ScrollBars: 2 - fmScrollBarsVertical
  - Locked: True
  - Height: ~100 pixels
  - BackColor: Light gray (for read-only appearance)

- **txtTarget**
  - MultiLine: True
  - ScrollBars: 2 - fmScrollBarsVertical
  - Locked: True
  - Height: ~60 pixels
  - BackColor: Light gray

- **txtReplace**
  - MultiLine: True
  - ScrollBars: 2 - fmScrollBarsVertical
  - Locked: True
  - Height: ~80 pixels
  - BackColor: Light gray

- **txtExplanation**
  - MultiLine: True
  - ScrollBars: 2 - fmScrollBarsVertical
  - Locked: True
  - Height: ~80 pixels
  - BackColor: Light gray

#### **Command Buttons (bottom of form):**
- **cmdAccept**
  - Caption: "Accept (A)"
  - Width: ~80 pixels
  - Default: True (makes it respond to Enter key)

- **cmdReject**
  - Caption: "Reject (R)"
  - Width: ~80 pixels

- **cmdSkip**
  - Caption: "Skip (S)"
  - Width: ~80 pixels

- **cmdAcceptAll**
  - Caption: "Accept All Remaining"
  - Width: ~120 pixels

- **cmdStop**
  - Caption: "Stop (ESC)"
  - Width: ~80 pixels
  - Cancel: True (makes it respond to ESC key)

### 3. Suggested Layout

```
┌─────────────────────────────────────────────────────┐
│         Suggestion 1 of 10                          │
│         Action: change                              │
├─────────────────────────────────────────────────────┤
│ Context:                                            │
│ ┌─────────────────────────────────────────────────┐ │
│ │ [txtContext - shows surrounding text]           │ │
│ │                                                  │ │
│ └─────────────────────────────────────────────────┘ │
│                                                     │
│ Target:                                             │
│ ┌─────────────────────────────────────────────────┐ │
│ │ [txtTarget - shows text to change]              │ │
│ └─────────────────────────────────────────────────┘ │
│                                                     │
│ Replacement:                                        │
│ ┌─────────────────────────────────────────────────┐ │
│ │ [txtReplace - shows new text]                   │ │
│ └─────────────────────────────────────────────────┘ │
│                                                     │
│ Explanation:                                        │
│ ┌─────────────────────────────────────────────────┐ │
│ │ [txtExplanation - shows why]                    │ │
│ └─────────────────────────────────────────────────┘ │
│                                                     │
│  [Accept] [Reject] [Skip] [Accept All] [Stop]      │
└─────────────────────────────────────────────────────┘
```

### 4. Quick Setup Alternative (Copy/Paste Properties)

If you want to speed this up, you can:
1. Create one control of each type
2. Copy it (Ctrl+C)
3. Paste it (Ctrl+V) to create duplicates
4. Rename each control using the Properties window (F4)
5. Adjust positions and sizes

### 5. Test the Form

1. Save your VBA project
2. Close the VBA Editor
3. Run the macro: `ApplyLlmReview_V3`
4. Paste some test JSON
5. Click "Process"
6. The new interactive form should appear for each suggestion

## Switching Between Workflows

To switch back to the old tracked changes workflow:

1. Open `wordAIreviewer.bas`
2. Find the `RunReviewProcess` function (around line 49)
3. Change this line:
   ```vba
   Const USE_INTERACTIVE_MODE As Boolean = True
   ```
   To:
   ```vba
   Const USE_INTERACTIVE_MODE As Boolean = False
   ```

## Features of the New Workflow

### Interactive Mode (NEW - Default)
- ✅ Shows each suggestion one at a time
- ✅ Highlights the context in the document
- ✅ Displays explanation and details
- ✅ Accept/Reject/Skip individual suggestions
- ✅ "Accept All Remaining" for batch processing
- ✅ Keyboard shortcuts (A/R/S/ESC)
- ✅ No tracked changes or comments (direct application)
- ✅ Can stop at any time
- ✅ **Smart Skip**: Automatically skips suggestions where formatting is already correct
  - Example: If "LAmax" already has subscript "Amax", the suggestion to change it to "L<sub>Amax</sub>" is skipped
  - Checks text content AND formatting (bold, italic, subscript, superscript)
  - Saves time by not re-applying identical formatting

### Tracked Changes Mode (OLD - Preserved)
- Applies all suggestions as tracked changes
- Adds comments with explanations
- Opens separate review form
- Navigate through grouped changes
- Good for formal review processes

## Keyboard Shortcuts in Interactive Mode

- **A** - Accept current suggestion
- **R** - Reject current suggestion
- **S** - Skip current suggestion
- **ESC** - Stop processing

## Troubleshooting

### Form doesn't appear
- Make sure you imported the .frm file
- Check that `USE_INTERACTIVE_MODE = True` in the code
- Look for errors in the Immediate Window (Ctrl+G in VBA Editor)

### Controls are missing
- You need to manually add the controls to the form (see step 2 above)
- The .frm file contains the code but not the visual layout

### Context not highlighting
- This is normal if the context isn't found
- The form will still show the suggestion details
- You can skip or reject suggestions that can't be found

## Tips for Best Results

1. **Form Size**: Make the form large enough to see context (suggested: 460x400 twips)
2. **Font Sizes**: Use 9-10pt for readability
3. **Colors**: Use light gray backgrounds for read-only textboxes
4. **Tab Order**: Set tab order so Accept button is first (most common action)
5. **Test First**: Try with a small JSON file (2-3 suggestions) before processing large batches

## Need Help?

If you encounter issues:
1. Check the Immediate Window (Ctrl+G) for debug messages
2. Verify all control names match exactly (case-sensitive)
3. Make sure the form's KeyPreview property is set to True for keyboard shortcuts
4. Test with the old workflow first to isolate form issues
