# AI Report Review Pipeline — VBA Integration Notes

This folder contains the VBA/Word side of the Venta Acoustics AI report review
pipeline. The Python orchestration and AI prompt files have moved to the
canonical skill location. This note exists so that changes to the VBA side
don't inadvertently break the Python side without the developer knowing.

---

## Where the Python code lives

**Canonical location:**
`G:\My Drive\Venta AI\skills\report-checking\scripts\`

Key files:
- `reviewer_agent.py` — main pipeline orchestrator (see below)
- `extract_dsm.py` — DSM extractor (Python, uses python-docx; must stay in sync with VBA DSM format)
- `extract_th_data.py` — TH workbook extractor
- `extract_ebf_data.py` — EBF workbook extractor
- `generate_review_html.py` — HTML review output generator
- `vision_scout.py` — visual cross-check agent

Prompt files:
`G:\My Drive\Venta AI\skills\report-checking\prompts\`
- `prompt_plan.txt`, `prompt_execute.txt`, `prompt_manual.txt`
- `prompt_plan_tech.txt`, `prompt_plan_consist.txt`, `prompt_plan_spell.txt`
- `prompt_plan_lang.txt`, `prompt_plan_template.txt`, `prompt_plan_b11.txt`

---

## What reviewer_agent.py does

`reviewer_agent.py` is the top-level pipeline orchestrator. It:

1. Locates the exported DSM markdown file (`{stem}_dsm.md`) for a given document
2. Runs five specialist LLM review agents in parallel (tech, consistency, spelling,
   language, template) using the prompt files in `prompts\`
3. Optionally runs a B1.1 synthesis agent (`--deep` flag) to reconcile findings
   against supporting evidence (TH/EBF data, email context, visual cross-check)
4. Runs an execute pass to produce V5 SEARCH/REPLACE blocks
5. Converts those blocks into `{stem}_toolcalls.json` for import back into Word

Invocation (from terminal):
```powershell
python 'G:\My Drive\Venta AI\skills\report-checking\scripts\reviewer_agent.py' --file <doc_stem> --runner auto --multi-agent
```

The script supports `--runner auto|codex|claude`, `--multi-agent`, `--deep`,
`--manual-output-file`, `--manual-paste`, `--exchange-dir`, and other flags.
See `--help` or `V5_MARKDOWN_DIFF_README.md` in this folder for the full reference.

---

## VBA ↔ Python interface — what must stay in sync

The Python scripts depend on the DSM format produced by the VBA macros. If you
change the VBA export logic, check whether `extract_dsm.py` or `reviewer_agent.py`
needs updating too.

### DSM format contract (`extract_dsm.py`)

`extract_dsm.py` is a Python reimplementation of the VBA DSM export, used when
Word COM is unavailable. It produces V4.2 format JSON. The following VBA
behaviours are replicated and **must be kept in sync**:

| VBA behaviour | Python equivalent | Risk if changed |
|---|---|---|
| Body paragraphs only (`INCLUDE_TABLE_PARAGRAPHS_IN_DSM = False`) | `extract_dsm.py` iterates `body_children` only | P-IDs will diverge from VBA if this changes |
| Table cells use `T{n}.R{r}.C{c}` IDs (1-based) | Cell ID format in `extract_dsm.py` | Toolcall targets will not resolve in Word |
| Tracked changes shown as "Final" view (insertions in, deletions out) | `_collect_final_text()` in `extract_dsm.py` | DSM text will differ from what Word shows |
| `format_spans` for bold/italic/sub/sup | `get_paragraph_plain_and_spans()` | Minor — affects formatting fidelity only |
| Version field `"4.2"` | Top-level `"version"` key in JSON output | `reviewer_agent.py` may version-check this |

### Toolcalls format contract (`reviewer_agent.py`)

Word's `V4_ImportAndApplyToolCalls` macro reads `{stem}_toolcalls.json` and
expects this structure per tool call:

```json
{
  "tool": "replace_text",
  "target": "P12",
  "search": "original text",
  "replace": "revised text"
}
```

If the VBA import macro changes its expected JSON shape, `reviewer_agent.py`'s
output serialisation must be updated to match.

### Exchange folder contract

Both VBA and Python expect files in:
- Primary: `G:\My Drive\Venta AI\projects\{job number} {job name}\`
- Fallback: `%TEMP%\claude_review\`

VBA macros write `{stem}_dsm.md` and `{stem}_dsm.json` here.
Python reads from and writes to the same folder.
If the VBA export destination changes, update `reviewer_agent.py`'s path
resolution logic (search for `projects` and `claude_review` in the file).

---

## VBA files in this folder

The following files are the Word/VBA side of the pipeline and belong here:

- `wordAIreviewer.bas` — main VBA module (`V4_ExportDocumentMapSilent`,
  `V4_ImportAndApplyToolCalls`, etc.)
- `va_helpers.bas` — shared VBA helper functions
- `frmReviewer.frm`, `frmJsonInput.frm`, `frmSuggestionPreview.frm` — UI forms
- `V5_MARKDOWN_DIFF_README.md` — V5 pipeline end-to-end usage reference

The `.dotm` add-in is loaded at runtime from the user's Word startup folder.
Source for the add-in is maintained as separate `.bas` / `.frm` files here and
extracted/injected via `extract_vba.py` (in `G:\My Drive\Programing\`).
