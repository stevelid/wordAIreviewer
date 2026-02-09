# WordAI Reviewer First-Principles Refactor Plan

Complete architectural redesign of `wordAIreviewer.bas` to use deterministic structural anchoring instead of fragile text-based context searching.

---

## Executive Summary

**Current Problems:**
- Text-based context matching fails frequently with auto-numbered headings, tables, long contexts, and TOC duplicates
- Progressive fallback strategies are band-aids that increase false positives
- Table cell finding uses heuristics that often fail
- No structural awareness - LLM doesn't know Word's internal structure

**Proposed Solution:**
- Generate a **Document Structure Map (DSM)** with unique anchors before sending to LLM
- Use **index-based targeting** (paragraph index, table index, cell coordinates) instead of text search
- Integrate directly with **VA Addin styles** for formatting
- Simpler JSON schema with deterministic location references

---

## Phase 1: Document Structure Map (DSM) Generator

### 1.1 Create `GenerateDocumentStructureMap()` Function

This function scans the document and generates a JSON-like structure that the LLM will use for context.

```
Output Format (Markdown for LLM):
---
# DOCUMENT STRUCTURE MAP

## P1 [Report Level 1] "1. Introduction"
## P2 [Report Text] "This report presents the findings of..."
## P3 [Report Level 2] "1.1 Site Description"
## P4 [Report Text] "The site is located at..."

## T1 "Table 7.1: Noise Survey Results"
| Row | Col1 | Col2 | Col3 |
|-----|------|------|------|
| R1  | Location | Day dB | Night dB |
| R2  | Position 1 | 52 | 45 |
| R3  | Position 2 | 58 | 48 |

## P12 [Report Level 2] "1.2 Assessment Criteria"
...
---
```

**Key Design Decisions:**
- Each paragraph gets a unique ID: `P1`, `P2`, etc.
- Each table gets a unique ID: `T1`, `T2`, etc.
- Style names are shown in brackets to help LLM understand document hierarchy
- Table cells are addressable by `T1.R2.C3` format (Table 1, Row 2, Column 3)
- Headings show their auto-generated numbers as part of the text (what user sees)

### 1.2 Style Mapping Dictionary

Map detected Word styles to VA Addin style functions:

| Word Style Name | VA Addin Function | JSON Style Key |
|-----------------|-------------------|----------------|
| Report Level 1 | `RChapter` | `"heading_l1"` |
| Report Level 2 | `RSectionheading` | `"heading_l2"` |
| Report Text | `RSection` | `"body_text"` |
| Report Level 3 | `RSubsection` | `"heading_l3"` |
| Report Level 4 | `RHeadingL4` | `"heading_l4"` |
| Report Bullet | `RBullet` | `"bullet"` |
| Table Heading | `Tableheading` | `"table_heading"` |
| Table Text | `Tabletext` | `"table_text"` |
| Report Table Number | `RTabletitle` | `"table_title"` |
| Report Figure | `Rfigure` | `"figure"` |

### 1.3 Implementation Steps

1. **Create `Type DocumentElement`** - Stores paragraph/table info with:
   - `ElementID` (P1, T1, etc.)
   - `ElementType` (paragraph, table, figure)
   - `StyleName` 
   - `TextPreview` (first 100 chars)
   - `StartPos`, `EndPos` (character positions for direct Range access)
   - `ParentIndex` (for hierarchical structure)

2. **Create `BuildDocumentStructureMap()`**:
   ```vba
   ' Iterate through all StoryRanges
   ' For each paragraph:
   '   - Assign sequential ID (P1, P2...)
   '   - Record style, position, text preview
   '   - Detect if it's a table title (paragraph before/after table)
   ' For each table:
   '   - Assign sequential ID (T1, T2...)
   '   - Record all cell contents with R/C coordinates
   '   - Link to associated title paragraph
   ```

3. **Create `ExportStructureMapAsMarkdown()`**:
   - Generates the markdown format shown above
   - This is what gets sent to the LLM for context

---

## Phase 2: New JSON Schema for LLM Suggestions

### 2.1 Simplified Action Schema

```json
{
  "suggestions": [
    {
      "target": "P5",
      "action": "replace",
      "find": "exisitng text",
      "replace": "existing text",
      "explanation": "Spelling correction"
    },
    {
      "target": "P12",
      "action": "apply_style",
      "style": "heading_l2",
      "explanation": "This should be a section heading"
    },
    {
      "target": "T2.R3.C2",
      "action": "replace",
      "find": "52",
      "replace": "53",
      "explanation": "Corrected noise level"
    },
    {
      "target": "T1",
      "action": "replace_table",
      "replace": "| Col1 | Col2 |\n|---|---|\n| A | B |",
      "explanation": "Updated table data"
    },
    {
      "target": "P8",
      "action": "comment",
      "explanation": "Consider expanding this section"
    },
    {
      "target": "T3",
      "action": "insert_row",
      "after_row": 2,
      "data": ["Location 4", "55", "47"],
      "explanation": "Added missing measurement location"
    }
  ]
}
```

### 2.2 Target Reference Types

| Target Format | Description | Example |
|---------------|-------------|---------|
| `P{n}` | Paragraph by index | `P5` |
| `T{n}` | Entire table | `T2` |
| `T{n}.R{r}` | Table row | `T2.R3` |
| `T{n}.R{r}.C{c}` | Table cell | `T2.R3.C2` |
| `T{n}.H.C{c}` | Table header cell | `T2.H.C1` |

### 2.3 Action Types

| Action | Required Fields | Optional Fields |
|--------|-----------------|-----------------|
| `replace` | `find`, `replace` | `match_case` |
| `apply_style` | `style` | - |
| `comment` | `explanation` | - |
| `delete` | - | - |
| `replace_table` | `replace` (markdown) | - |
| `insert_row` | `data` (array) | `after_row`, `before_row` |
| `delete_row` | - | - |
| `insert_paragraph` | `text`, `style` | `after` (target ref) |

---

## Phase 3: Core Processing Engine Rewrite

### 3.1 Main Entry Point: `ApplyLlmReview_V4()`

```
1. Show input form (frmJsonInput)
2. Parse JSON suggestions
3. Load/refresh DocumentStructureMap
4. For each suggestion:
   a. Resolve target reference to Word Range
   b. Validate action is applicable
   c. Show preview (if interactive mode)
   d. Execute action
5. Report results
```

### 3.2 Target Resolution: `ResolveTargetToRange()`

This is the **key function** that replaces all the fragile text searching:

```vba
Private Function ResolveTargetToRange(ByVal targetRef As String) As Range
    ' Parse target reference format
    ' P5 -> Get paragraph 5 from structure map -> Return its Range
    ' T2 -> Get table 2 from structure map -> Return its Range
    ' T2.R3.C2 -> Get table 2, row 3, col 2 -> Return cell Range
    
    ' Uses pre-built structure map - NO TEXT SEARCHING
    ' Direct position-based Range creation
End Function
```

### 3.3 Action Executors

Create separate, focused functions for each action type:

- `ExecuteReplaceAction()` - Find/replace within resolved range
- `ExecuteApplyStyleAction()` - Apply VA Addin style function
- `ExecuteCommentAction()` - Add comment to range
- `ExecuteDeleteAction()` - Delete range content
- `ExecuteReplaceTableAction()` - Replace entire table
- `ExecuteInsertRowAction()` - Insert table row with data
- `ExecuteDeleteRowAction()` - Delete table row
- `ExecuteInsertParagraphAction()` - Insert new paragraph with style

### 3.4 Style Application Integration

```vba
Private Sub ApplyVAStyle(ByVal rng As Range, ByVal styleKey As String)
    rng.Select
    Select Case LCase$(styleKey)
        Case "heading_l1": Call RChapter
        Case "heading_l2": Call RSectionheading
        Case "body_text", "text": Call RSection
        Case "heading_l3": Call RSubsection
        Case "heading_l4": Call RHeadingL4
        Case "bullet": Call RBullet
        Case "table_heading": Call Tableheading
        Case "table_text": Call Tabletext
        Case "table_title": Call RTabletitle
        Case "figure": Call Rfigure
        Case Else
            ' Fallback: try to apply as Word built-in style
            On Error Resume Next
            rng.Style = styleKey
            On Error GoTo 0
    End Select
End Sub
```

---

## Phase 4: User Interface Updates

### 4.1 Input Form Updates (frmJsonInput)

Add new features:
- **"Generate Structure Map"** button - Creates and displays the DSM
- **"Copy to Clipboard"** button - Copies DSM for pasting to LLM
- **Structure Map Preview** text box - Shows generated map
- Keep existing JSON input area

### 4.2 Preview Form Updates (frmSuggestionPreview)

Simplify since targeting is now deterministic:
- Show target reference (P5, T2.R3.C2)
- Show resolved location (highlight in document)
- Show proposed change
- Accept/Skip/Stop buttons

### 4.3 New: Structure Map Viewer (Optional)

Simple form showing:
- Tree view of document structure
- Click to navigate to element
- Shows element ID for reference

---

## Phase 5: LLM Prompt Template

### 5.1 System Prompt for Document Review

```markdown
You are a technical document reviewer. You will receive a Document Structure Map (DSM) 
showing the structure of a Word document.

## DSM Format:
- P{n} = Paragraph number n
- T{n} = Table number n  
- [Style Name] = Current paragraph style
- T{n}.R{r}.C{c} = Table n, Row r, Column c

## Your Response Format:
Return a JSON array of suggestions. Each suggestion must have:
- "target": The element reference (P5, T2.R3.C2, etc.)
- "action": One of: replace, apply_style, comment, delete, replace_table, insert_row, delete_row
- "explanation": Why this change is needed

For "replace" actions, include "find" and "replace" fields.
For "apply_style" actions, include "style" field with one of:
  heading_l1, heading_l2, heading_l3, heading_l4, body_text, bullet, 
  table_heading, table_text, table_title, figure

## Style Hierarchy (use for apply_style):
1. heading_l1 - Chapter titles (e.g., "1. Introduction")
2. heading_l2 - Section headings (e.g., "1.1 Site Description")  
3. heading_l3 - Subsection headings (italic)
4. heading_l4 - Sub-subsection headings
5. body_text - Normal paragraph text
6. bullet - Bullet points
7. table_heading - Table header row cells
8. table_text - Table body cells
9. table_title - Table captions (e.g., "Table 7.1: Results")
10. figure - Figure captions

## Example Response:
```json
[
  {"target": "P5", "action": "replace", "find": "recieved", "replace": "received", "explanation": "Spelling"},
  {"target": "P12", "action": "apply_style", "style": "heading_l2", "explanation": "Should be section heading"},
  {"target": "T2.R3.C2", "action": "replace", "find": "52", "replace": "53", "explanation": "Incorrect value"},
  {"target": "P8", "action": "comment", "explanation": "Consider adding methodology details"}
]
```

Now review the following document:
```

---

## Phase 6: Migration & Compatibility

### 6.1 File Structure

Create new module or significantly refactor existing:

```
wordAIreviewer_v4.bas (new main module)
├── Document Structure Map functions
├── Target resolution functions  
├── Action executor functions
├── Style integration functions
├── JSON parsing (keep existing)
├── UI form handlers

frmJsonInput.frm (updated)
├── Structure map generation UI
├── JSON input area

frmSuggestionPreview.frm (simplified)
├── Target-based preview
```

### 6.2 Functions to Keep from Current Implementation

- `LLM_ParseJson()` - JSON parser works well
- `PreProcessJson()` - JSON cleanup
- `NormalizeForDocument()` - Text normalization
- `ParseFormattingTags()` - HTML tag parsing for `<b>`, `<i>`, etc.
- `ApplyFormattedReplacement()` - But remove Font.Reset call
- `ConvertMarkdownToTable()` - Table creation from markdown

### 6.3 Functions to Remove/Replace

- `FindWithProgressiveFallback()` - Replaced by `ResolveTargetToRange()`
- `FindLongString()` - No longer needed
- `FuzzyFindString()` - No longer needed
- `FindTableCell()` - Replaced by coordinate-based lookup
- `FindTableByTitle()` - Replaced by index-based lookup
- All `TextMatchesHeuristic()` family - No longer needed
- `BuildTableIndex()` - Replaced by DSM

---

## Phase 7: Implementation Order

### Step 1: Core Infrastructure (Day 1)
1. Create `DocumentElement` Type
2. Create `g_DocumentMap()` global array
3. Implement `BuildDocumentStructureMap()`
4. Implement `ExportStructureMapAsMarkdown()`
5. Test: Generate map for sample document

### Step 2: Target Resolution (Day 1-2)
1. Implement `ParseTargetReference()` - Parse "T2.R3.C2" format
2. Implement `ResolveTargetToRange()` - Return Word Range
3. Test: Verify correct range selection for various targets

### Step 3: Style Integration (Day 2)
1. Implement `ApplyVAStyle()` with all style mappings
2. Test: Apply each style type correctly

### Step 4: Action Executors (Day 2-3)
1. Implement `ExecuteReplaceAction()`
2. Implement `ExecuteApplyStyleAction()`
3. Implement `ExecuteCommentAction()`
4. Implement `ExecuteDeleteAction()`
5. Implement table actions (replace, insert_row, delete_row)
6. Test: Each action type individually

### Step 5: Main Processing Loop (Day 3)
1. Implement `ApplyLlmReview_V4()` main entry
2. Implement `ProcessSuggestionV4()` 
3. Implement preflight validation
4. Test: End-to-end with sample JSON

### Step 6: UI Updates (Day 4)
1. Update frmJsonInput with DSM generation
2. Simplify frmSuggestionPreview for target-based preview
3. Test: Full workflow

### Step 7: Testing & Polish (Day 4-5)
1. Test with real documents containing:
   - Auto-numbered headings
   - Multiple similar tables
   - Long paragraphs
   - TOC
2. Error handling refinement
3. Debug logging improvements

---

## Phase 8: Testing Checklist

### Unit Tests
- [ ] DSM generates correct paragraph indices
- [ ] DSM generates correct table indices  
- [ ] DSM captures correct styles
- [ ] Target "P5" resolves to correct paragraph
- [ ] Target "T2.R3.C2" resolves to correct cell
- [ ] Each action type executes correctly
- [ ] VA Addin styles apply correctly

### Integration Tests
- [ ] Full workflow: Generate DSM → Get LLM response → Apply changes
- [ ] Auto-numbered headings handled correctly
- [ ] Multiple similar tables distinguished correctly
- [ ] TOC not affected by changes
- [ ] Track Changes works with new system
- [ ] Undo works correctly

### Edge Cases
- [ ] Empty document
- [ ] Document with only tables
- [ ] Nested tables
- [ ] Merged cells
- [ ] Very long documents (100+ pages)
- [ ] Documents with fields/bookmarks

---

## Appendix A: Sample Document Structure Map Output

```markdown
# DOCUMENT STRUCTURE MAP
Generated: 2024-01-15 14:30:00

## Paragraphs

P1 [Report Level 1] "1. Introduction"
P2 [Report Text] "This report presents the findings of the noise impact assessment for the proposed residential development at 123 Example Street, London."
P3 [Report Level 2] "1.1 Site Description"
P4 [Report Text] "The site is located in a predominantly residential area..."
P5 [Report Text] "The surrounding area comprises..."
P6 [Report Level 2] "1.2 Proposed Development"
P7 [Report Text] "The proposed development consists of..."
P8 [Report Level 1] "2. Assessment Criteria"
P9 [Report Level 2] "2.1 National Planning Policy Framework"
P10 [Report Text] "The NPPF states that..."
P11 [Report Table Number] "Table 2.1: BS8233:2014 Internal Noise Criteria"

## Tables

T1 [after P11] "BS8233:2014 Internal Noise Criteria"
| R | C1 | C2 | C3 |
|---|----|----|-----|
| 1 | Room Type | Day (07:00-23:00) | Night (23:00-07:00) |
| 2 | Living Room | 35 dB LAeq,16hr | - |
| 3 | Bedroom | 35 dB LAeq,16hr | 30 dB LAeq,8hr |

P12 [Report Level 2] "2.2 Local Planning Policy"
P13 [Report Text] "The local development plan requires..."
...
```

---

## Appendix B: Error Handling Strategy

### Target Resolution Errors
```vba
' If target not found in structure map:
' 1. Log warning with target reference
' 2. Attempt to rebuild structure map (document may have changed)
' 3. If still not found, skip suggestion with user notification
' 4. Continue processing remaining suggestions
```

### Action Execution Errors
```vba
' If action fails:
' 1. Log detailed error (target, action, error message)
' 2. Wrap in UndoRecord so partial changes can be reverted
' 3. Notify user of specific failure
' 4. Continue with next suggestion
```

### Style Application Errors
```vba
' If VA Addin style function fails:
' 1. Fall back to direct style assignment
' 2. If that fails, log and skip
' 3. Never leave document in inconsistent state
```

---

## Appendix C: Why Not Markdown Anchoring?

The user asked about generating markdown with anchoring tags. This approach was considered but rejected because:

1. **Lossy conversion** - Converting Word to markdown loses formatting, styles, and structure
2. **Round-trip problems** - Changes in markdown would need to be mapped back to Word positions
3. **Table complexity** - Word tables with merged cells, formatting don't map cleanly to markdown
4. **Style information lost** - Can't preserve VA Addin styles in markdown
5. **Position drift** - If document is edited, anchors become invalid

The DSM approach provides structural anchoring **within Word's native model**, avoiding these issues.

---

## Summary

This refactor replaces fragile text-based searching with deterministic index-based targeting. The LLM receives a structured map of the document and returns changes using stable references (P5, T2.R3.C2) that map directly to Word Ranges. This eliminates the root cause of most failures: ambiguous text matching.

**Estimated effort**: 4-5 days for a skilled VBA developer
**Risk level**: Medium - significant architectural change but well-defined scope
**Backward compatibility**: Not required per user specification
