"""
Extract Document Structure Map (DSM) from a Word document using python-docx.
Produces JSON output compatible with the V4.2 VBA LLMReviewTools format.

Tracked changes are handled as "Final" view:
  - Inserted text (w:ins) is INCLUDED
  - Deleted text (w:del) is EXCLUDED
This matches Word's "Final" display mode without actually accepting changes.

V4.2 vs V4.1 changes:
  - Paragraph IDs count only body paragraphs (table cell paragraphs excluded)
  - Table elements use a 'cells' array with T{n}.R{r}.C{c} IDs
  - format_spans added to paragraphs (1-based char indices, bold/italic/sub/sup)
  - heading_level, within_table, section_number, page_number added to paragraphs
  - Top-level structure: {version, document, tooling, elements}

Usage:
    python extract_dsm.py "path/to/report.docx"

Output:
    Writes {stem}_dsm.json to %TEMP%/claude_review/

Element IDs:
    P1, P2, P3... — body paragraphs only (table cell paragraphs not counted)
    T1, T2, T3... — tables
    T1.R1.C1 etc — table cells (1-based row/col indices)
"""

import sys
import os
import json
import re
from pathlib import Path
from datetime import datetime
from lxml import etree
from docx import Document

# Word XML namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


# ---------------------------------------------------------------------------
# Text extraction helpers
# ---------------------------------------------------------------------------

def _collect_final_text(element, parts: list):
    """Recursively collect text in Final view (include w:ins, skip w:del)."""
    tag = element.tag
    if tag == f"{W}del":
        return
    if tag == f"{W}t":
        if element.text:
            parts.append(element.text)
        return
    if tag == f"{W}tab":
        parts.append("\t")
        return
    if tag == f"{W}br":
        parts.append("\n")
        return
    for child in element:
        _collect_final_text(child, parts)


def get_final_text_from_element(element) -> str:
    parts = []
    _collect_final_text(element, parts)
    return "".join(parts)


def is_in_del(elem) -> bool:
    """Check if an element is nested inside a w:del (deleted text)."""
    parent = elem.getparent()
    while parent is not None:
        if parent.tag == f"{W}del":
            return True
        parent = parent.getparent()
    return False


def get_run_formatting(run_elem) -> tuple:
    """Return (bold, italic, subscript, superscript) for a run element."""
    rpr = run_elem.find(f"{W}rPr")
    bold = italic = sub = sup = False
    if rpr is not None:
        b = rpr.find(f"{W}b")
        if b is not None:
            val = b.get(f"{W}val", "true")
            bold = val.lower() not in ("false", "0")
        i_elem = rpr.find(f"{W}i")
        if i_elem is not None:
            val = i_elem.get(f"{W}val", "true")
            italic = val.lower() not in ("false", "0")
        vert = rpr.find(f"{W}vertAlign")
        if vert is not None:
            v = vert.get(f"{W}val", "")
            sub = v == "subscript"
            sup = v == "superscript"
    return bold, italic, sub, sup


def get_run_text_parts(run_elem) -> str:
    """Extract text from a single run element."""
    text = ""
    for child in run_elem:
        ctag = child.tag
        if ctag == f"{W}t":
            if child.text:
                text += child.text
        elif ctag == f"{W}tab":
            text += "\t"
        elif ctag == f"{W}br":
            text += "\n"
    return text


def get_paragraph_plain_and_spans(para_elem) -> tuple:
    """
    Extract text_plain (including trailing \\n for paragraph mark) and
    format_spans (1-based char indices) from a paragraph element.

    format_spans only records runs with bold, italic, subscript, or superscript.
    Consecutive characters with identical formatting are merged into one span.

    Returns: (text_plain: str, format_spans: list, text_tagged: str)
    """
    run_parts = []  # list of (text, bold, italic, sub, sup)

    for run_elem in para_elem.iter(f"{W}r"):
        if is_in_del(run_elem):
            continue
        text = get_run_text_parts(run_elem)
        if not text:
            continue
        bold, italic, sub, sup = get_run_formatting(run_elem)
        run_parts.append((text, bold, italic, sub, sup))

    # text_plain: join all run text + trailing \n (Word paragraph mark)
    text_plain = "".join(p[0] for p in run_parts) + "\n"

    # format_spans: 1-based, only where formatting is present
    spans = []
    pos = 1  # 1-based character position
    for text, bold, italic, sub, sup in run_parts:
        for _ch in text:
            if bold or italic or sub or sup:
                if (spans and
                        spans[-1]["start"] + spans[-1]["length"] == pos and
                        spans[-1]["bold"] == bold and
                        spans[-1]["italic"] == italic and
                        spans[-1]["subscript"] == sub and
                        spans[-1]["superscript"] == sup):
                    spans[-1]["length"] += 1
                else:
                    spans.append({
                        "start": pos,
                        "length": 1,
                        "subscript": sub,
                        "superscript": sup,
                        "bold": bold,
                        "italic": italic,
                    })
            pos += 1

    # text_tagged: inline tags for bold/italic/sub/sup
    tagged_parts = []
    for text, bold, italic, sub, sup in run_parts:
        t = text
        if sub:
            t = f"<sub>{t}</sub>"
        elif sup:
            t = f"<sup>{t}</sup>"
        if bold:
            t = f"<b>{t}</b>"
        if italic:
            t = f"<i>{t}</i>"
        tagged_parts.append(t)
    text_tagged = "".join(tagged_parts) + "\n"

    return text_plain, spans, text_tagged


def get_cell_text(cell_elem) -> str:
    """
    Extract trimmed text from a table cell element.
    Multiple paragraphs within the cell are joined with \\n.
    """
    para_texts = []
    for p_elem in cell_elem.findall(f"{W}p"):
        para_text = get_final_text_from_element(p_elem)
        para_texts.append(para_text)

    # Join paragraphs with \n (representing paragraph marks between them)
    # Then strip leading/trailing whitespace (matches VBA's Trim$)
    raw = "\n".join(para_texts)
    return raw.strip()


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def get_style_name_simple(para_elem, doc) -> str:
    """Get the display style name for a paragraph element."""
    ppr = para_elem.find(f"{W}pPr")
    if ppr is not None:
        pstyle = ppr.find(f"{W}pStyle")
        if pstyle is not None:
            style_id = pstyle.get(f"{W}val", "Normal")
            if not hasattr(doc, "_style_id_map"):
                doc._style_id_map = {s.style_id: s.name for s in doc.styles}
            return doc._style_id_map.get(style_id, style_id)
    return "Normal"


def get_heading_level(style_name: str) -> int:
    """
    Return the heading level for a style name, or 0 if not a heading.
    Handles standard Word styles ('Heading 1') and Venta custom styles
    ('Report Level 1').
    """
    normalized = style_name.lower().strip()
    for prefix in ("heading ", "report level "):
        if normalized.startswith(prefix):
            suffix = normalized[len(prefix):].strip()
            if suffix.isdigit():
                return int(suffix)
    return 0


# ---------------------------------------------------------------------------
# Table title / caption helpers
# ---------------------------------------------------------------------------

def get_table_title_and_caption(tbl_index: int, body_children: list) -> tuple:
    """
    Look for a caption paragraph above and title paragraph below the table
    in the body children list.

    Returns: (title_text: str, caption_text: str)
    """
    title_text = ""
    caption_text = ""

    # Find the table in body_children
    tbl_pos = None
    table_count = 0
    for i, child in enumerate(body_children):
        if child.tag == f"{W}tbl":
            table_count += 1
            if table_count == tbl_index:
                tbl_pos = i
                break

    if tbl_pos is None:
        return title_text, caption_text

    # Caption: look up to 3 non-empty body paragraphs above
    for i in range(tbl_pos - 1, max(tbl_pos - 4, -1), -1):
        child = body_children[i]
        if child.tag == f"{W}p":
            text = get_final_text_from_element(child).strip()
            if len(text) > 3:
                caption_text = text
                break
        elif child.tag == f"{W}tbl":
            break  # Stop at another table

    # Title: look up to 3 non-empty body paragraphs below
    for i in range(tbl_pos + 1, min(tbl_pos + 4, len(body_children))):
        child = body_children[i]
        if child.tag == f"{W}p":
            text = get_final_text_from_element(child).strip()
            if len(text) > 3:
                title_text = text
                break
        elif child.tag == f"{W}tbl":
            break  # Stop at another table

    return title_text, caption_text


# ---------------------------------------------------------------------------
# Tooling section
# ---------------------------------------------------------------------------

def build_tooling_section() -> dict:
    """Build the V4.2 tooling documentation block."""
    return {
        "overview": "Use these instructions to convert DSM v4.2 data into executable Word tool calls.",
        "ordering": "Elements are listed in document order (top-to-bottom).",
        "text_view": "Paragraph and table previews reflect Word's Final view (insertions included, deletions excluded).",
        "guidelines": [
            "Always anchor targets using DSM IDs from the elements array; never rely on searching raw text.",
            "Confine replacements to the resolved paragraph or cell to avoid accidental edits elsewhere.",
            "Paragraph IDs refer to body paragraphs; table cell text should be targeted with T#.R#.C# references.",
            "Favour table-specific tools (replace_table, insert_table_row, delete_table_row) for structured data changes.",
            "Add comments when recommending manual review steps or when data is missing.",
            "Keep tool calls atomic\u2014one logical change per entry.",
        ],
        "style_tokens": [
            "heading_l1", "heading_l2", "heading_l3", "heading_l4",
            "body_text", "bullet", "table_heading", "table_text",
            "table_title", "figure",
        ],
        "target_reference_format": [
            {
                "kind": "paragraph",
                "format": "P{n}",
                "description": "Use for any paragraph listed in the DSM elements array. The numeric suffix matches the paragraph order.",
                "example": "P5",
            },
            {
                "kind": "table",
                "format": "T{n}",
                "description": "References an entire table. Use this when replacing a whole table or inserting rows relative to it.",
                "example": "T2",
            },
            {
                "kind": "table_row",
                "format": "T{n}.R{r}",
                "description": "References a specific row within table n (e.g., R1 = first data row).",
                "example": "T2.R3",
            },
            {
                "kind": "table_header_row",
                "format": "T{n}.H",
                "description": "References the table header row (equivalent to row 1).",
                "example": "T2.H",
            },
            {
                "kind": "table_cell",
                "format": "T{n}.R{r}.C{c}",
                "description": "References a single cell using row and column numbers.",
                "example": "T2.R3.C2",
            },
            {
                "kind": "table_header_cell",
                "format": "T{n}.H.C{c}",
                "description": "Targets a header row cell for formatting or text edits.",
                "example": "T2.H.C1",
            },
        ],
    }


# ---------------------------------------------------------------------------
# Main DSM extraction
# ---------------------------------------------------------------------------

def extract_dsm(docx_path: str) -> dict:
    """
    Extract a V4.2 Document Structure Map from a Word document.

    Paragraph numbering:
      - Counts only body-level paragraphs (table cell paragraphs excluded).
      - Matches VBA's INCLUDE_TABLE_PARAGRAPHS_IN_DSM = False behaviour.

    Tables:
      - Each table gets a T{n} element with a 'cells' array.
      - Cells use T{n}.R{r}.C{c} IDs (1-based row and column indices).
    """
    doc = Document(docx_path)
    body = doc.element.body
    body_children = list(body)

    elements = []
    para_counter = 0
    table_counter = 0

    for child in body_children:

        if child.tag == f"{W}p":
            # Body paragraph — include in main elements list
            para_counter += 1
            text_plain, format_spans, text_tagged = get_paragraph_plain_and_spans(child)
            style_name = get_style_name_simple(child, doc)
            heading_level = get_heading_level(style_name)

            # Use text_tagged only when it adds information beyond plain text
            tagged_output = text_tagged if text_tagged != text_plain else text_plain

            elements.append({
                "id": f"P{para_counter}",
                "kind": "paragraph",
                "style": style_name,
                "text_plain": text_plain,
                "text_tagged": tagged_output,
                "format_spans": format_spans,
                "range": {"start": 0, "end": 0},  # Not computed (requires Word COM)
                "section_number": 0,
                "page_number": 0,
                "heading_level": heading_level,
                "within_table": False,
            })

        elif child.tag == f"{W}tbl":
            # Table — enumerate cells, build cell list
            table_counter += 1
            title_text, caption_text = get_table_title_and_caption(
                table_counter, body_children
            )

            rows = child.findall(f"{W}tr")
            num_rows = len(rows)
            num_cols = 0
            if rows:
                num_cols = len(rows[0].findall(f"{W}tc"))

            cells = []
            for row_idx, row in enumerate(rows, start=1):
                for col_idx, cell_elem in enumerate(row.findall(f"{W}tc"), start=1):
                    cell_text = get_cell_text(cell_elem)
                    cells.append({
                        "id": f"T{table_counter}.R{row_idx}.C{col_idx}",
                        "text_plain": cell_text,
                        "text_tagged": "",  # Matches VBA DSM_INCLUDE_TABLE_CELL_TAGGED_TEXT=False
                        "format_spans": [],  # Matches VBA DSM_INCLUDE_TABLE_CELL_FORMAT_SPANS=False
                    })

            elements.append({
                "id": f"T{table_counter}",
                "kind": "table",
                "rows": num_rows,
                "cols": num_cols,
                "range": {"start": 0, "end": 0},
                "section_number": 0,
                "page_number": 0,
                "title_text": title_text,
                "caption_text": caption_text,
                "within_table": False,
                "cells": cells,
            })

    doc_name = Path(docx_path).name
    generated_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

    dsm = {
        "version": "4.2",
        "document": {
            "name": doc_name,
            "generated_at": generated_at,
        },
        "tooling": build_tooling_section(),
        "elements": elements,
    }
    return dsm


def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_dsm.py <path_to_docx>")
        sys.exit(1)

    docx_path = sys.argv[1]
    if not os.path.exists(docx_path):
        print(f"ERROR: File not found: {docx_path}")
        sys.exit(1)

    output_dir = os.path.join(os.environ.get("TEMP", "/tmp"), "claude_review")
    os.makedirs(output_dir, exist_ok=True)

    print(f"Extracting DSM from: {docx_path}")
    dsm = extract_dsm(docx_path)

    stem = Path(docx_path).stem
    output_path = os.path.join(output_dir, f"{stem}_dsm.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(dsm, f, indent=2, ensure_ascii=False)

    para_count = sum(1 for e in dsm["elements"] if e["kind"] == "paragraph")
    table_count = sum(1 for e in dsm["elements"] if e["kind"] == "table")
    print(f"Elements: {len(dsm['elements'])} ({para_count} paragraphs, {table_count} tables)")
    print(f"DSM written to: {output_path}")
    print(f"Size: {os.path.getsize(output_path):,} bytes")


if __name__ == "__main__":
    main()
