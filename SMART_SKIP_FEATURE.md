# Smart Skip Feature - Automatic Formatting Detection

## Overview
The system now intelligently detects when formatting is already applied and **automatically skips** those suggestions, saving you time and avoiding redundant changes.

## How It Works

### Example Scenario:
**LLM Suggestion:**
```json
{
  "action": "change",
  "context": "The LAmax level was measured.",
  "target": "LAmax",
  "replace": "L<sub>Amax</sub>",
  "explanation": "Apply subscript to acoustic term"
}
```

### What Happens:

#### **Case 1: Formatting NOT Applied**
- Document has: `LAmax` (plain text)
- Suggestion wants: `L<sub>Amax</sub>` (with subscript)
- **Result**: ✅ Change is applied

#### **Case 2: Formatting ALREADY Applied**
- Document has: `LAmax` (with "Amax" already subscripted)
- Suggestion wants: `L<sub>Amax</sub>` (with subscript)
- **Result**: ⏭️ **SKIPPED** - No change made, no comment added

## What Gets Checked

The system compares:

1. **Text Content** (case-insensitive)
   - "LAmax" vs "LAmax" ✓

2. **Bold Formatting**
   - Checks if `<b>text</b>` matches actual bold

3. **Italic Formatting**
   - Checks if `<i>text</i>` matches actual italic

4. **Subscript Formatting**
   - Checks if `<sub>text</sub>` matches actual subscript

5. **Superscript Formatting**
   - Checks if `<sup>text</sup>` matches actual superscript

## Benefits

### ✅ **Saves Time**
- Don't waste time reviewing suggestions that are already correct
- Especially useful when re-running the same document

### ✅ **Prevents Redundancy**
- No duplicate comments saying "apply subscript" when it's already applied
- Cleaner document with fewer unnecessary changes

### ✅ **Handles Partial Formatting**
- If document has `L<sub>A</sub>max` but suggestion wants `L<sub>Amax</sub>`, it will still apply (different formatting)

### ✅ **Debug Visibility**
- Check the Immediate Window (Ctrl+G in VBA Editor) to see skip messages:
  ```
  Action 'change': SKIPPED - Formatting already matches for 'LAmax'
  ```

## Common Use Cases

### Acoustic Terms
- `LAeq` → `L<sub>Aeq</sub>`
- `LAmax` → `L<sub>Amax</sub>`
- `LA90` → `L<sub>A90</sub>`

### Scientific Notation
- `m2` → `m<sup>2</sup>`
- `CO2` → `CO<sub>2</sub>`
- `H2O` → `H<sub>2</sub>O`

### Emphasis
- `important` → `<b>important</b>`
- `note` → `<i>note</i>`

## When It Doesn't Skip

The change **will still be applied** if:

1. **Text content differs**
   - Document: `LAmax`, Suggestion: `LAeq` → Applied

2. **Formatting is different**
   - Document: `LAmax` (no formatting), Suggestion: `L<sub>Amax</sub>` → Applied
   - Document: `L<sub>A</sub>max`, Suggestion: `L<sub>Amax</sub>` → Applied (different subscript range)

3. **Formatting is partial**
   - Document: `LAmax` (only "A" subscripted), Suggestion: `L<sub>Amax</sub>` → Applied

4. **Additional text changes**
   - Document: `LAmax`, Suggestion: `L<sub>Aeq,T</sub>` → Applied (text differs)

## Technical Details

### Function: `IsFormattingAlreadyApplied()`
Located in `wordAIreviewer.bas`, this function:

1. Parses the replacement text to extract plain text and formatting tags
2. Compares plain text (case-insensitive)
3. If text matches, checks each formatting segment
4. Returns `True` if everything matches (skip), `False` otherwise

### Integration Point
Called in `ExecuteSingleAction()` before applying "change" actions:

```vba
Case "change"
    If IsFormattingAlreadyApplied(actionRange, replaceText) Then
        Debug.Print "SKIPPED - Formatting already matches"
        ' No change applied, no comment added
    Else
        ApplyFormattedReplacement actionRange, replaceText
        ' Add comment if explanation exists
    End If
```

## Debugging

To see what's happening:

1. Open VBA Editor (`Alt+F11`)
2. Open Immediate Window (`Ctrl+G`)
3. Run your review process
4. Watch for messages like:
   ```
   Action 'change': SKIPPED - Formatting already matches for 'LAmax'
   Action 'change': Replacing 'LAeq' with 'LAmax'
   ```

## Configuration

This feature is **always enabled** and cannot be disabled. If you need to force re-application of formatting:

1. Manually remove the formatting from the document first
2. Or modify the text slightly (add a space, then remove it)
3. Then run the review process

## Performance Impact

**Minimal** - The check is very fast:
- Parses replacement text (already done for application)
- Compares strings (instant)
- Checks font properties (instant)
- Only runs when text content matches

Typical overhead: **< 1ms per suggestion**
