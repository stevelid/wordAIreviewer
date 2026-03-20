# AI Report Review — Chat Web Interface Guide

Use this when reviewing a report via a chat LLM (e.g. ChatGPT, Claude web, Gemini) instead of the automated pipeline.

---

## Steps

### 1. Export the DSM from Word

In Word, run the macro:

```
V4_ExportDocumentMapToFile
```

This creates `{doc_stem}_dsm.md` in the project exchange folder.

The DSM now includes a metadata line declaring the assumed standard body and
table styles, plus inline style hints only where a paragraph or table cell uses
non-standard formatting. For example, `[P12]` means standard style, while
`[P4|Report Level 2]` means that paragraph uses a non-standard style.

Highlighted text may also appear as `<mark>...</mark>`, which reflects Word
highlighting in the source document. Treat this as formatting metadata rather
than literal angle-bracket text.

### 2. Send to the LLM

In your chat interface:

1. **First message:** Paste the full contents of `prompt_chat.txt` (found in `G:\My Drive\Venta AI\skills\report-checking\prompts\prompt_chat.txt`).
2. **Second message:** Paste the full contents of the `_dsm.md` file.

Wait for the LLM to respond with SEARCH/REPLACE blocks.

### 3. Save the LLM output

Copy the entire LLM response and save it as a `.txt` file, e.g.:

```
C:\temp\llm_output.txt
```

### 4. Convert to tool calls

Open a terminal and run:

```powershell
python "G:\My Drive\Venta AI\skills\report-checking\scripts\reviewer_agent.py" --file <doc_stem> --manual-output-file C:\temp\llm_output.txt
```

Or paste directly:

```powershell
python "G:\My Drive\Venta AI\skills\report-checking\scripts\reviewer_agent.py" --file <doc_stem> --manual-paste
```

Then paste the LLM response and press `Ctrl+Z` then `Enter` to end input.

Replace `<doc_stem>` with the document stem, e.g. `6237.260309.ADR`.

This generates `{doc_stem}_toolcalls.json` in the exchange folder.

If you are using the JSON prompt/tool path instead of SEARCH/REPLACE, the tool
schema now also supports `clear_highlight` for removing highlight without
changing text, and `delete_table` for removing an entire table.

### 5. Import into Word

In Word, run the macro:

```
V4_ImportAndApplyToolCalls
```

This reads the tool calls JSON and applies edits as tracked changes.

---

## Common mistakes

| Mistake | Fix |
|---|---|
| Gave the LLM a V4 JSON prompt instead of SEARCH/REPLACE | Use `prompt_chat.txt` — it asks for SEARCH/REPLACE blocks |
| Pasted JSON tool calls directly into `V4_RunInteractiveReview` | That path needs V4.1 JSON, not SEARCH/REPLACE. Use the pipeline above instead |
| LLM wrapped output in markdown fences | The parser handles this, but tell the LLM "no fences" if it happens |
| `reviewer_agent.py` says "Blocks parsed: 0" | The LLM output didn't contain valid `<<<<<<< SEARCH` / `>>>>>>> REPLACE` markers |
| Word import says file missing | Check that `_toolcalls.json` exists in the exchange folder and the stem matches the active document |

---

## File locations

| File | Location |
|---|---|
| Chat prompt | `G:\My Drive\Venta AI\skills\report-checking\prompts\prompt_chat.txt` |
| Converter script | `G:\My Drive\Venta AI\skills\report-checking\scripts\reviewer_agent.py` |
| Exported DSM | `G:\My Drive\Venta AI\projects\{job folder}\{doc_stem}_dsm.md` |
| Generated tool calls | Same folder as the DSM: `{doc_stem}_toolcalls.json` |
