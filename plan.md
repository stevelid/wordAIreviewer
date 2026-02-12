# WordAI Reviewer Refactor Plan (V4 Only)

This plan is locked to your decisions:
- V4 schema only (no V3/context schema support in runtime path)
- Simplified tool-call format only
- All public entry points grouped at the end of `wordAIreviewer.bas`

---

## 1. Scope

### In scope
- Deterministic V4 review pipeline based on target IDs (`P#`, `T#.R#.C#`)
- Per-suggestion interactive review with robust preflight
- Document Map export that includes existing subscript/superscript state
- Unified schema validation and execution

### Out of scope
- Maintaining old `context/actions` V3 schema in active flow
- Hybrid V3+V4 runtime behavior

---

## 2. Canonical V4 Contracts

## 2.1 Document Map contract (input to LLM)

```json
{
  "version": "4.1",
  "document": {
    "name": "Report.docx",
    "generated_at": "2026-02-09T12:34:56"
  },
  "elements": [
    {
      "id": "P42",
      "kind": "paragraph",
      "style": "Report Text",
      "text_plain": "Assessment based on LAeq over 16hr.",
      "text_tagged": "Assessment based on L<sub>Aeq</sub> over 16hr.",
      "format_spans": [
        { "start": 23, "length": 3, "subscript": true, "superscript": false }
      ]
    },
    {
      "id": "T3",
      "kind": "table",
      "rows": 3,
      "cols": 2,
      "cells": [
        {
          "id": "T3.R2.C2",
          "text_plain": "m2",
          "text_tagged": "m<sup>2</sup>",
          "format_spans": [
            { "start": 2, "length": 1, "subscript": false, "superscript": true }
          ]
        }
      ]
    }
  ]
}
```

Notes:
- `text_plain` is unformatted content.
- `text_tagged` encodes real formatting from Word using `<sub>` and `<sup>` (plus `<b>/<i>` if present).
- `format_spans` is used for exact no-op detection and execution checks.

## 2.2 Simplified tool-call contract (output from LLM)

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

Allowed `tool` values:
- `replace_text`
- `apply_style`
- `add_comment`
- `delete_range`
- `replace_table`
- `insert_table_row`
- `delete_table_row`

Validation rules:
- Unknown tool names: reject suggestion.
- Missing required `args` fields: reject suggestion.
- Old keys (`context`, `actions`, `apply_heading_style`, etc.): reject as invalid V4 schema.
- Target must parse into valid V4 reference.

---

## 3. Execution Architecture

## 3.1 Pipeline
1. Parse and validate tool-call JSON.
2. Build current Document Map snapshot.
3. Preflight each suggestion:
   - target exists
   - action valid for target type
   - no-op check (including formatting parity)
4. Interactive review (accept/reject/skip/accept-all/stop).
5. Execute accepted suggestions with per-suggestion undo record.
6. Emit final report (applied/skipped/no-op/failed).

## 3.2 Per-suggestion review behavior
- Always show:
  - `tool`
  - `target`
  - before preview
  - proposed after preview
  - explanation/confidence
- `Accept All` must continue from current index only (never re-iterate from start).
- If target cannot resolve: show as `NOT RESOLVED` and default to skip.

---

## 4. Subscript/Superscript Intelligence

Goal: avoid unnecessary suggestions when sub/sup already correct.

Implementation:
- While building map, read character formatting from each paragraph/cell range.
- Emit `text_tagged` with `<sub>` / `<sup>` markers that reflect actual Word formatting.
- Emit `format_spans` for deterministic comparison.
- Preflight rule:
  - If `replace_text` changes only formatting tags and target formatting already matches, mark as `NO_OP`.
  - `NO_OP` suggestions are skipped by default and shown in review summary.

Acceptance checks:
- If document already has `L<sub>Aeq</sub>`, the same suggestion is skipped.
- If formatting differs (partial or wrong span), suggestion remains actionable.

---

## 5. Module Layout and Ordering Rule

`wordAIreviewer.bas` must be reorganized into ordered regions:

1. Constants, enums, types, module state
2. Private shared helpers (normalization, parsing, formatting utilities)
3. Document Map builders/exporters
4. Target parsing and resolution
5. Schema validation + preflight analyzers
6. Action executors (`replace_text`, `apply_style`, etc.)
7. Interactive review/controller logic
8. Reporting/logging helpers
9. `PUBLIC ENTRY POINTS` section (final section only)

Hard rule:
- Every `Public Sub` / `Public Function` appears only in section 9 at file end.
- Public entry points are thin wrappers that call private workers above.

Planned public entry points (at end):
- `Public Sub V4_GenerateDocumentMap()`
- `Public Sub V4_CopyDocumentMapToClipboard()`
- `Public Sub V4_ValidateToolCallsJson()`
- `Public Sub V4_ApplyToolCalls()`
- `Public Sub V4_RunInteractiveReview()`
- `Public Sub V4_TestTargetResolution()`

---

## 6. Implementation Phases

## Phase 1: Contract and validator
- Add strict V4 schema parser for `tool_calls`.
- Reject non-V4 keys and legacy action names.

## Phase 2: Document Map upgrade
- Build `text_plain`, `text_tagged`, and `format_spans`.
- Ensure map is in true document order.

## Phase 3: Resolver hardening
- Resolve all targets deterministically (`P#`, `T#`, `T#.R#`, `T#.R#.C#`, `T#.H.C#`).
- Validate target type compatibility per tool.

## Phase 4: Executor alignment
- Rename/align action executors to tool names.
- Apply formatted replacements to matched subrange, not whole target block.
- Add per-suggestion `UndoRecord`.

## Phase 5: Review UX
- Update preview form with before/after and no-op reason.
- Fix `Accept All` continuation indexing.

## Phase 6: File reordering
- Move all public entry points to final section.
- Keep all internal logic private above.

## Phase 7: Docs and prompt alignment
- Update `V4_USAGE_GUIDE.md` to only V4 schema.
- Replace `prompt.txt` with V4 tool-call instructions only.

---

## 7. File-level Change Plan

- `wordAIreviewer.bas`
  - Primary refactor target (schema, map, resolver, executors, ordering).
- `frmSuggestionPreview.frm`
  - Add V4-focused review display fields and result state.
- `frmJsonInput.frm`
  - V4 parse/validate/apply hooks only (no V3 routing).
- `V4_USAGE_GUIDE.md`
  - Match exact V4 schema and tool list.
- `prompt.txt`
  - Emit only V4 tool calls.

---

## 8. Acceptance Criteria

- V3 schema JSON is rejected with a clear validation message.
- V4 tool calls parse and execute deterministically.
- Existing correct sub/sup formatting is represented in map and skipped as no-op.
- Interactive review supports accept/reject/skip/accept-all/stop without duplicate processing.
- All public entry points are physically located at end of `wordAIreviewer.bas`.
- No mixed V3/V4 runtime path remains.

---

## 9. Rollout Sequence

1. Implement schema validator and map tagging.
2. Implement resolver/executor updates.
3. Update forms.
4. Reorder module and finalize public entry-point section.
5. Update docs (`V4_USAGE_GUIDE.md`, `prompt.txt`).
6. Run regression pass on representative reports (tables, TOC, existing sub/sup).
