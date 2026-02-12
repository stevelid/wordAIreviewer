"""
Extract Document Structure Map (DSM) from a Word document using python-docx.
Produces JSON output compatible with the V4.1 VBA LLMReviewTools format.

Tracked changes are handled as "Final" view:
  - Inserted text (w:ins) is INCLUDED
  - Deleted text (w:del) is EXCLUDED
This matches Word's "Final" display mode without actually accepting changes.

IMPORTANT: Paragraph numbering matches Word's doc.Paragraphs collection,
which includes ALL paragraphs — body-level AND inside table cells, including
empty ones. This ensures P-IDs are compatible with the VBA macro's numbering.

Usage:
    python extract_dsm.py "path/to/report.docx"

Output:
    Writes {stem}_dsm.json to %TEMP%/claude_review/

Element IDs match the VBA macro's numbering:
    P1, P2, P3... — all paragraphs (including empty and table-internal)
    T1, T2, T3... — tables (separate counter)
"""

import sys
import os
import json
from pathlib import Path
from lxml import etree
from docx import Document

# Word XML namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def get_final_text_from_element(element) -> str:
    """
    Extract text from an XML element as it would appear in Word's "Final" view.
    Includes inserted text (w:ins), excludes deleted text (w:del).
    """
    parts = []
    _collect_final_text(element, parts)
    return "".join(parts)


def _collect_final_text(element, parts: list):
    """Recursively collect text, skipping w:del elements entirely."""
    tag = element.tag

    # Skip deleted text entirely — not visible in "Final" view
    if tag == f"{W}del":
        return

    # w:t elements contain actual text
    if tag == f"{W}t":
        if element.text:
            parts.append(element.text)
        return

    # w:tab = tab character, w:br = line break
    if tag == f"{W}tab":
        parts.append("\t")
        return
    if tag == f"{W}br":
        parts.append("\n")
        return

    # Recurse into children (including w:ins — its content IS visible)
    for child in element:
        _collect_final_text(child, parts)


def get_final_tagged_text(para_elem) -> str:
    """
    Extract text with formatting tags from a paragraph element,
    respecting tracked changes (Final view).

    Tags: <b>, <i>, <sub>, <sup>
    """
    parts = []
    # Iterate over all runs in the paragraph, including those inside w:ins
    for run_elem in para_elem.iter(f"{W}r"):
        # Skip if this run is inside a w:del
        parent = run_elem.getparent()
        in_del = False
        while parent is not None:
            if parent.tag == f"{W}del":
                in_del = True
                break
            parent = parent.getparent()
        if in_del:
            continue

        # Get run text
        text = ""
        for t_elem in run_elem.findall(f"{W}t"):
            if t_elem.text:
                text += t_elem.text
        for tab_elem in run_elem.findall(f"{W}tab"):
            text += "\t"
        if not text:
            continue

        # Check formatting via run properties (w:rPr)
        rpr = run_elem.find(f"{W}rPr")
        is_bold = False
        is_italic = False
        is_sub = False
        is_sup = False

        if rpr is not None:
            # Bold
            b_elem = rpr.find(f"{W}b")
            if b_elem is not None:
                val = b_elem.get(f"{W}val", "true")
                is_bold = val.lower() not in ("false", "0")

            # Italic
            i_elem = rpr.find(f"{W}i")
            if i_elem is not None:
                val = i_elem.get(f"{W}val", "true")
                is_italic = val.lower() not in ("false", "0")

            # Subscript / Superscript
            vert_elem = rpr.find(f"{W}vertAlign")
            if vert_elem is not None:
                vert_val = vert_elem.get(f"{W}val", "")
                is_sub = vert_val == "subscript"
                is_sup = vert_val == "superscript"

        # Apply tags
        if is_sub:
            text = f"<sub>{text}</sub>"
        elif is_sup:
            text = f"<sup>{text}</sup>"
        if is_bold:
            text = f"<b>{text}</b>"
        if is_italic:
            text = f"<i>{text}</i>"

        parts.append(text)

    return "".join(parts)


def get_style_name(para_elem, doc) -> str:
    """Get the style name for a paragraph element."""
    ppr = para_elem.find(f"{W}pPr")
    if ppr is not None:
        pstyle = ppr.find(f"{W}pStyle")
        if pstyle is not None:
            style_id = pstyle.get(f"{W}val", "Normal")
            # Try to resolve style ID to display name via document styles
            try:
                style = doc.styles.get_by_id(style_id, doc.styles._element.get_type())
            except Exception:
                # Fallback: look up by style_id directly
                for s in doc.styles:
                    if s.style_id == style_id:
                        return s.name
                return style_id
            if style:
                return style.name
    return "Normal"


def get_style_name_simple(para_elem, doc) -> str:
    """Get the style name for a paragraph element using simple lookup."""
    ppr = para_elem.find(f"{W}pPr")
    if ppr is not None:
        pstyle = ppr.find(f"{W}pStyle")
        if pstyle is not None:
            style_id = pstyle.get(f"{W}val", "Normal")
            # Build style map on first call (cached on doc object)
            if not hasattr(doc, '_style_id_map'):
                doc._style_id_map = {}
                for s in doc.styles:
                    doc._style_id_map[s.style_id] = s.name
            return doc._style_id_map.get(style_id, style_id)
    return "Normal"


def extract_table_markdown_final(table_elem) -> str:
    """
    Convert a Word table XML element to markdown, using Final view for text.
    """
    rows = table_elem.findall(f"{W}tr")
    md_rows = []

    for i, row in enumerate(rows):
        cells = row.findall(f"{W}tc")
        cell_texts = []
        for cell in cells:
            # Get all paragraph text in the cell
            cell_paras = []
            for p in cell.findall(f"{W}p"):
                para_text = get_final_text_from_element(p).strip()
                if para_text:
                    cell_paras.append(para_text)
            cell_text = " ".join(cell_paras).replace("|", "\\|")
            cell_texts.append(cell_text)

        if i == 0:
            num_cols = len(cell_texts)

        md_rows.append("| " + " | ".join(cell_texts) + " |")

        # Add separator after header row
        if i == 0:
            md_rows.append("|" + "|".join(["---"] * len(cell_texts)) + "|")

    return "\n".join(md_rows)


def extract_dsm(docx_path: str) -> dict:
    """
    Extract a Document Structure Map from a Word document.
    Tracked changes shown as "Final" view.

    Paragraph numbering matches Word's doc.Paragraphs collection:
    ALL w:p elements in document order (body + table cells + empty).
    """
    doc = Document(docx_path)
    body = doc.element.body
    elements = []

    para_counter = 0
    table_counter = 0

    # Walk ALL w:p elements in document order (matches Word's Paragraphs collection)
    # Also track tables as we encounter them
    tables_seen = set()

    def process_body(body_elem):
        """Process body elements, yielding paragraphs and tables in document order."""
        nonlocal para_counter, table_counter

        for child in body_elem:
            if child.tag == f"{W}p":
                # Every paragraph gets counted, including empty ones
                para_counter += 1
                text_plain = get_final_text_from_element(child).strip()
                text_tagged = get_final_tagged_text(child) if text_plain else ""
                style_name = get_style_name_simple(child, doc)

                elements.append({
                    "id": f"P{para_counter}",
                    "type": "paragraph",
                    "style": style_name,
                    "text_plain": text_plain if text_plain else "\n",
                    "text_tagged": text_tagged if (text_tagged and text_tagged != text_plain) else (text_plain if text_plain else "\n"),
                })

            elif child.tag == f"{W}tbl":
                # Enumerate paragraphs inside the table row by row.
                # Word's doc.Paragraphs counts each cell's w:p elements PLUS
                # an end-of-row marker paragraph after each row.
                for tr in child.findall(f"{W}tr"):
                    # All w:p elements in this row's cells
                    for p_elem in tr.findall(f".//{W}p"):
                        para_counter += 1
                        text_plain = get_final_text_from_element(p_elem).strip()
                        text_tagged = get_final_tagged_text(p_elem) if text_plain else ""
                        style_name = get_style_name_simple(p_elem, doc)

                        elements.append({
                            "id": f"P{para_counter}",
                            "type": "paragraph",
                            "style": style_name,
                            "text_plain": text_plain if text_plain else "\n",
                            "text_tagged": text_tagged if (text_tagged and text_tagged != text_plain) else (text_plain if text_plain else "\n"),
                        })

                    # End-of-row marker — Word counts this as a paragraph
                    para_counter += 1
                    elements.append({
                        "id": f"P{para_counter}",
                        "type": "paragraph",
                        "style": "Normal",
                        "text_plain": "\n",
                        "text_tagged": "\n",
                    })

                # Then add the table itself
                table_counter += 1
                rows = child.findall(f"{W}tr")
                num_rows = len(rows)
                num_cols = 0
                header_cells = []

                if num_rows > 0:
                    first_row_cells = rows[0].findall(f"{W}tc")
                    num_cols = len(first_row_cells)
                    for cell in first_row_cells:
                        cell_text = get_final_text_from_element(cell).strip()
                        header_cells.append(cell_text)

                markdown = extract_table_markdown_final(child)

                elements.append({
                    "id": f"T{table_counter}",
                    "type": "table",
                    "rows": num_rows,
                    "cols": num_cols,
                    "header_row": header_cells,
                    "markdown": markdown,
                })

    process_body(body)

    doc_name = Path(docx_path).name
    dsm = {
        "version": "4.1",
        "source": doc_name,
        "generated_by": "python-docx extract_dsm.py (Final view)",
        "element_count": len(elements),
        "paragraph_count": para_counter,
        "table_count": table_counter,
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

    # Output folder
    output_dir = os.path.join(os.environ.get("TEMP", "/tmp"), "claude_review")
    os.makedirs(output_dir, exist_ok=True)

    # Extract DSM
    print(f"Extracting DSM from: {docx_path}")
    dsm = extract_dsm(docx_path)

    # Write JSON
    stem = Path(docx_path).stem
    output_path = os.path.join(output_dir, f"{stem}_dsm.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(dsm, f, indent=2, ensure_ascii=False)

    print(f"Elements: {dsm['element_count']} ({dsm['paragraph_count']} paragraphs, {dsm['table_count']} tables)")
    print(f"DSM written to: {output_path}")
    print(f"Size: {os.path.getsize(output_path):,} bytes")


if __name__ == "__main__":
    main()
