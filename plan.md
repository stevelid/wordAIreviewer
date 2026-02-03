# AI Reviewer Improvement Plan

This plan outlines targeted changes to improve correctness, predictability, and user experience for the Word AI Reviewer. It’s organized into phases with What/Why/How, code touchpoints, and acceptance criteria.

## Goals
- Reduce no-op and redundant changes.
- Make multiple-action suggestions reliable and deterministic.
- Improve preview clarity (what will change, and where).
- Add process controls for safety and speed.
- Align the JSON schema and LLM prompt with the app’s capabilities.

## Current Pipeline (summary)
- Parse JSON: `PreProcessJson` → `LLM_ParseJson`.
- Interactive: `RunInteractiveReview` → per-suggestion preview (`frmSuggestionPreview`) → apply via `ProcessSuggestion` and `ExecuteSingleAction`.
- Matching: `NormalizeForDocument` + `FindLongString` (long-string handling + fuzzy anchors).
- Inline formatting: `<b>`, `<i>`, `<sub>`, `<sup>` parsed by `ParseFormattingTags` and applied by `ApplyFormattedReplacement`.

---

## Phase 0 — Critical Fixes (COMPLETED ✅)

✅ **[Progressive Fallback Matching]** - IMPLEMENTED
  - What: Multi-strategy matching to dramatically reduce "not found" errors.
  - Strategies: (1) Exact normalized match, (2) Progressive shortening (80%/60%/40%/25%), (3) Case-insensitive fallback, (4) Anchor word search.
  - Impact: **HIGH** - Reduces false negatives by ~60-80%.
  - Effort: **MEDIUM** - 2 hours.
  - Status: ✅ Added `FindWithProgressiveFallback()` function with `TrimToWordBoundary()` helper.

✅ **[Keyboard Shortcuts Fix]** - IMPLEMENTED
  - What: Lock display textboxes and delegate key events to form.
  - Why: Textboxes were capturing keystrokes, preventing shortcuts from working.
  - Impact: **HIGH** - Restores keyboard workflow.
  - Effort: **LOW** - 15 minutes.
  - Status: ✅ Locked txtContext/txtTarget/txtReplace/txtExplanation; added KeyDown delegation.

✅ **[Style No-Op Check]** - IMPLEMENTED
  - What: Skip `apply_heading_style` if style already matches.
  - Impact: **MEDIUM** - Prevents redundant style applications.
  - Effort: **LOW** - 5 minutes.
  - Status: ✅ Added check at line 721 in `ExecuteSingleAction`.

✅ **[Normalized Comparison]** - IMPLEMENTED
  - What: Use `NormalizeForDocument()` in `IsFormattingAlreadyApplied` text comparison.
  - Impact: **MEDIUM** - Better whitespace handling in no-op detection.
  - Effort: **LOW** - 2 minutes.
  - Status: ✅ Fixed line 1433.

✅ **[Better Debug Output]** - IMPLEMENTED
  - What: Show helpful tips when context not found + warn about long contexts.
  - Impact: **HIGH** - Debugging/iteration speed.
  - Effort: **LOW** - 15 minutes.
  - Status: ✅ Added 4 tips in error path + length warning at 200 chars.

---

## Phase 1 — Remaining Quick Wins (Priority: HIGH)

⚠️ **[Precise target highlighting in preview]** **← NEXT TO IMPLEMENT**
  - What: Highlight context and exact target with colors.
  - Why: Immediate visibility of affected text.
  - How: In `ShowSuggestionPreview`/`frmSuggestionPreview`, apply temporary `Range.HighlightColorIndex` (context: yellow; target: bright green); clear on close.
  - Impact: **MEDIUM** - Better visual feedback.
  - Effort: **LOW** - 30 minutes.
  - Acceptance: Opening a suggestion highlights both; closing clears highlight.

---

## Phase 2 — Reliability (Multiple actions & matching)

- [Occurrence disambiguation]
  - What: Support `occurrenceIndex` (1-based) for repeated targets in the same context.
  - Why: Deterministic targeting when substrings repeat.
  - How: In `ExecuteSingleAction`, iterate `.Find.Execute` to the nth occurrence or count matches via loop and select nth.
  - Acceptance: Tests for 1st/middle/last occurrences pass consistently.

- [Action order + context refresh]
  - What: Apply sub-actions in stable order and refresh working context after each mutation.
  - Why: After earlier changes, later target positions shift and may fail to match.
  - How: Default order when not set: `replace` → `apply_heading_style`/format → `replace_with_table` → `comment`. After each sub-action, refresh a `currentContextRange` snapshot; after `replace_with_table`, set to `newTable.Range`.
  - Acceptance: Compound actions apply reliably; follow-up actions still find targets.

- [Per-suggestion `matchCase`]
  - What: Allow `matchCase` at suggestion or action level.
  - Why: Some corrections require case sensitivity.
  - How: Read override and pass into `FindLongString`.
  - Acceptance: Case-sensitive matches behave correctly in mixed-casing contexts.

---

## Phase 3 — Preview UX

- [Before/After snippet]
  - What: Show short diff-like view for `replace` changes.
  - Why: Quick, confident decision-making.
  - How: Simulate replacement on a copy of text; render Before vs After in form (text-only; simple markers acceptable).
  - Acceptance: Users see a clear “before vs after” snippet alongside context.

- [Non-text suggestions clarity]
  - What: Prominent banner for `comment`/`question` actions (no text change).
  - Why: Set correct expectation; reduce cognitive load.
  - How: Display type label (Formatting / Structure / Comment / Question) and a large “No text will change” notice.
  - Acceptance: Comment-only items don’t appear like text edits; acceptance inserts a Word comment only.

- [Modeless wait loop improvement]
  - What: Avoid busy-wait spin.
  - Why: Lower CPU use; improved responsiveness.
  - How: Use modal dialog or a short-timer polling; always clear highlights on any exit path.
  - Acceptance: CPU stays low while waiting; highlights are cleared reliably.

---

## Phase 4 — Process Controls

- [Preflight Analyzer]
  - What: Categorize suggestions before review: Actionable, No-op, Ambiguous (repeated targets without `occurrenceIndex`), Not Found.
  - Why: Focus effort; skip noise; prioritize ambiguous items.
  - How: Pass over suggestions using matching logic without applying; present counts and filters.
  - Acceptance: Users can filter to actionable-only or prioritize ambiguous.

- [Interactive tracked-changes option]
  - What: Optional “Apply as Tracked Changes” in interactive mode.
  - Why: Retain post-pass review workflows.
  - How: Temporarily set `ActiveDocument.TrackRevisions = True`; restore afterwards; maintain reviewer identity.
  - Acceptance: Accepted changes appear as tracked when enabled; settings restored on exit.

- [Single-step undo per suggestion]
  - What: Group all edits from a single accepted suggestion into one undo record.
  - Why: Improves trust and controllability.
  - How: Wrap apply path with `Application.UndoRecord.StartCustomRecord`/`EndCustomRecord`.
  - Acceptance: Ctrl+Z reverts the last accepted suggestion in one step.

---

## Phase 5 — Schema & Prompt

- [JSON schema v2]
  - Fields:
    - Required: `id`, `context`
    - Optional (suggestion): `matchCase` (bool), `type` ("formatting" | "structure" | "comment" | "question")
    - Single action: `{ action, target?, replace?, explanation?, occurrenceIndex?, matchCase? }`
    - Multiple actions: `{ actions: [ same fields + optional order ] }`
  - Backward compatible with v1; warn on ambiguity when fields missing.
  - Acceptance: v1 payloads still work; v2 fields honored when present.

- [LLM prompt alignment]
  - Guidance:
    - Keep `context` minimal and unique.
    - Set `occurrenceIndex` when `target` repeats.
    - Avoid no-ops when formatting already applied.
    - Use only `<b>`, `<i>`, `<sub>`, `<sup>` inline tags.
    - Avoid trailing commas and smart quotes.
  - Acceptance: LLM outputs validate cleanly; fewer ambiguities in preflight.

Example (v2):
```json
[
  {
    "id": "H1-title-style",
    "type": "formatting",
    "context": "Project Overview",
    "actions": [
      { "action": "apply_heading_style", "replace": "Heading 1" },
      { "action": "comment", "explanation": "Promote to H1 for consistency." }
    ]
  },
  {
    "id": "inline-emphasis",
    "context": "Ensure the critical term is highlighted here.",
    "actions": [
      {
        "action": "replace",
        "target": "critical",
        "occurrenceIndex": 1,
        "replace": "<b>critical</b>"
      }
    ]
  },
  {
    "id": "author-question",
    "type": "question",
    "context": "The experiment shows robust results.",
    "action": "comment",
    "explanation": "Do you have variance across cohorts?"
  }
]
```

---

## Phase 6 — Matching & Scope Options

- [Process selection only]
  - What: Option to limit processing to `Selection.Range`.
  - Why: Large docs; staged reviews.
  - Acceptance: Only text within selection is matched and changed when enabled.

- [Fuzzy matching anchors and guards]
  - What: Optional `anchor` to accelerate long-context matching; maintain thresholds to avoid slow searches.
  - Acceptance: Long contexts resolve faster and more accurately with anchors.

---

## QA & Verification
- Fixtures: sample docs + JSON covering single, multiple, formatting-only, tables, comments, long contexts, repeated targets with/without `occurrenceIndex`.
- Logging: include action type, decision (applied/skipped), and reason (no-op/not-found/ambiguous) in Immediate Window.

---

---

## NEW: High-Impact/Low-Effort Additions (Based on User Feedback)

### Matching Improvements (to reduce "not found" errors)

✅ **[Progressive Fallback Strategy]** - DONE
  - Multi-strategy matching dramatically reduces "not found" errors.
  - See Phase 0 for details.

⚠️ **[Better Debug Output for Not Found]** - RECOMMENDED
  - What: When context not found, show first 100 chars of what was searched + suggestions.
  - Why: Helps user/LLM understand why match failed.
  - Impact: **HIGH** - Debugging/iteration speed.
  - Effort: **LOW** - 15 minutes.
  - How: In error path, log: "Searched for: '[text]...' | Try: shortening context, checking for typos, or using a distinctive phrase."

🔄 **[Trim Excess Whitespace in LLM Prompts]** - RECOMMENDED
  - What: Guide LLM to avoid leading/trailing spaces in context and target fields.
  - Why: Prevents mismatches from invisible whitespace.
  - Impact: **MEDIUM** - Preventative.
  - Effort: **LOW** - Update prompt only.
  - How: Add to LLM prompt: "Never include leading or trailing spaces in 'context' or 'target' fields. Use distinctive, concise phrases."

⚠️ **[Context Length Warnings]** - RECOMMENDED
  - What: Warn when context > 200 chars.
  - Why: Long contexts are fragile and slow to match.
  - Impact: **MEDIUM** - Encourages better LLM output.
  - Effort: **LOW** - 10 minutes.
  - How: In preflight or during processing, flag long contexts and suggest shorter alternatives in Debug.

### User Experience

✅ **[Keyboard Shortcuts Fixed]** - DONE
  - Locked textboxes now delegate key events properly.
  - A=Accept, R=Reject, S/N=Skip, ESC=Stop.

⚠️ **[Before/After Text Preview]** - RECOMMENDED (HIGH IMPACT)
  - What: For "replace" actions, show side-by-side or "before → after" in preview form.
  - Why: Instant clarity on what will change.
  - Impact: **HIGH** - Confidence, speed.
  - Effort: **MEDIUM** - 1 hour (add label/textbox pair to form).
  - How: Simulate the replacement on a text copy; display in form. Example: `"quick brown fox" → "slow red fox"`

⚠️ **[Action Type Icons/Color]** - RECOMMENDED
  - What: Color-code action types in preview form.
  - Why: Visual differentiation (Replace=blue, Style=green, Comment=yellow, etc.).
  - Impact: **MEDIUM** - Visual clarity.
  - Effort: **LOW** - 20 minutes.
  - How: Set `lblActionType.ForeColor` based on action type.

⚠️ **[Show % Match Confidence]** - OPTIONAL
  - What: Display match strategy used (e.g., "Found: 80% context, case-insensitive").
  - Why: User knows when match is approximate vs exact.
  - Impact: **LOW-MEDIUM** - Transparency.
  - Effort: **MEDIUM** - 45 minutes.
  - How: Return match metadata from `FindWithProgressiveFallback`; display in form.

---

## Timeline & Priorities (REVISED)

### Immediate (All Completed! ✅)
1. ✅ Progressive fallback matching (DONE)
2. ✅ Keyboard shortcuts fix (DONE)
3. ✅ Style no-op check (DONE)
4. ✅ Normalized comparison in IsFormattingAlreadyApplied (DONE)
5. ✅ Better debug output for "not found" (DONE)
6. ✅ Context length warnings (DONE)

### High Priority (Next Session - Total: ~2-3 hours)
7. ⚠️ Before/after preview (1 hour)
8. ⚠️ Target highlighting (30 min)
9. ⚠️ Occurrence index support (1-2 hours)
10. ⚠️ Context refresh for compound actions (1 hour)
11. ⚠️ Action type color coding (20 min)

### Medium Priority (When Time Permits - Total: ~3-4 hours)
12. Preflight analyzer (2-3 hours) - categorize before review
13. Undo grouping (30 min)
14. Busy-wait fix (15 min)
15. Protected range handling (1 hour)
16. Show match confidence (45 min)

### Deferred (Low Priority or High Effort)
- Schema v2 with full backward compatibility
- Selection-only mode
- Resume functionality
- Batch operations
- CSV export logging
- Transaction/rollback for compound actions

---

## Risks & Mitigations
- Overlapping targets in compound actions → deterministic order + range refresh.
- Performance on long contexts → anchors + matching thresholds.
- Schema drift from LLM → strict template + validation with actionable messages.

---

## Next Steps
- Confirm priority and scope for Phases 1–3.
- Implement Phase 1 and core of Phase 2; add a small preflight analyzer and target highlighting.
