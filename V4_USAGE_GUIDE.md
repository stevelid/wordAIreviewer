# WordAI Reviewer V4.1 Guide (Tool Calls Only)

## Overview
V4.1 uses a strict tool-call schema and deterministic target references.
No V3 `context/actions` schema is accepted.

## Entry Points
- `V4_GenerateDocumentMap` - Builds/copies document map JSON to clipboard.
- `V4_RunInteractiveReview` - Paste tool-call JSON and review per suggestion.
- `V4_ApplyToolCalls` - Paste tool-call JSON and apply without per-item review.
- `V4_ValidateToolCallsJson` - Validate JSON schema only.
- `V4_TestTargetResolution` - Target resolution diagnostics.

## Target Format
- `P{n}` paragraph
- `T{n}` table
- `T{n}.R{r}` table row
- `T{n}.R{r}.C{c}` table cell
- `T{n}.H.C{c}` table header cell

## LLM Output Schema
```json
{
  "version": "4.1",
  "tool_calls": [
    {
      "id": "S001",
      "tool": "replace_text",
      "target": "P42",
      "args": {
        "find": "LAeq",
        "replace": "L<sub>Aeq</sub>",
        "match_case": false
      },
      "explanation": "Apply standard acoustic notation",
      "confidence": 0.93
    }
  ]
}
```

## Allowed Tools
- `replace_text` (`args.find`, `args.replace`, optional `args.match_case`)
- `apply_style` (`args.style`)
- `add_comment` (`args.text` or top-level `explanation`)
- `delete_range` (no required args)
- `replace_table` (`args.markdown`)
- `insert_table_row` (`args.data`, optional `args.after_row`)
- `delete_table_row` (target row reference)

## Notes
- Formatting tags are supported in replacements: `<b>`, `<i>`, `<sub>`, `<sup>`.
- V4 exports document map with `text_plain`, `text_tagged`, and formatting spans.
- Formatting-only suggestions are skipped when already applied.
