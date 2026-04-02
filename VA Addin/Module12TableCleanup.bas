Attribute VB_Name = "Module12TableCleanup"
Option Explicit

' ============================================================================
' Module12TableCleanup — Bulk table formatting cleanup
'
' Iterates every table in the active document and enforces consistent
' Venta report styles:
'   Table style  : "Report Table"
'   Header cells : "Table Heading" paragraph style
'   Body cells   : "Table Text" paragraph style
'   Caption para : "Report Table Number" (paragraph immediately after table)
'
' Heading-row detection:
'   1. Any row whose HeadingFormat is already True is treated as a header.
'   2. If no row has HeadingFormat, row 1 is assumed to be the header and
'      its HeadingFormat is set to True.
'
' Caption detection:
'   The first non-empty paragraph immediately after each table is checked.
'   If its text begins with "Table " (the standard Venta caption prefix)
'   it is styled as "Report Table Number".
'
' Usage:
'   - Run FormatAllTables manually from the Macros dialog, or
'   - Call from Python via COM:  doc.Application.Run("FormatAllTables")
' ============================================================================

Public Sub FormatAllTables()
    Dim doc As Document
    Dim tbl As Table
    Dim rw As Row
    Dim hasHeadingRow As Boolean
    Dim isHeading As Boolean
    Dim tblCount As Long
    Dim i As Long

    Set doc = ActiveDocument

    ' Ensure styles exist by copying from the addin template
    On Error Resume Next
    doc.CopyStylesFromTemplate AddinFolder & "\VA Addin.dotm"
    On Error GoTo 0

    Application.ScreenUpdating = False

    tblCount = doc.Tables.Count
    For i = 1 To tblCount
        Set tbl = doc.Tables(i)

        ' --- 1. Apply table style ---
        ApplyReportTableStyle tbl, doc

        ' --- 2. Identify heading rows ---
        '     HeadingFormat can error on rows involved in vertical merges.
        hasHeadingRow = TableHasHeadingRow(tbl)

        ' If no row is flagged, default row 1 to heading
        If Not hasHeadingRow Then
            On Error Resume Next
            tbl.Rows(1).HeadingFormat = True
            On Error GoTo 0
        End If

        ' --- 3. Apply cell paragraph styles ---
        '     Preserve existing paragraph alignment so numeric/right-aligned
        '     columns set by the builder survive the cleanup pass.
        On Error Resume Next
        For Each rw In tbl.Rows
            isHeading = False
            isHeading = rw.HeadingFormat
            StyleRowCells rw, doc, isHeading
        Next rw
        On Error GoTo 0

        ' --- 4. Style any caption/title row inside the table ---
        StyleCaptionRowInsideTable tbl, doc

        ' --- 5. Style the caption paragraph after the table ---
        StyleCaptionAfterTable tbl, doc
    Next i

    Application.ScreenUpdating = True
    Call LogDebug(tblCount & " table(s) formatted.", "FormatAllTables")
End Sub


' ---------------------------------------------------------------------------
' Apply the standard table style.
' ---------------------------------------------------------------------------
Private Sub ApplyReportTableStyle(tbl As Table, doc As Document)
    On Error Resume Next
    tbl.Style = doc.Styles("Report Table")
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' True if any row is already flagged as a heading row.
' ---------------------------------------------------------------------------
Private Function TableHasHeadingRow(tbl As Table) As Boolean
    Dim rw As Row
    Dim hasHeadingRow As Boolean
    Dim hf As Boolean

    hasHeadingRow = False
    On Error Resume Next
    For Each rw In tbl.Rows
        hf = False
        hf = rw.HeadingFormat
        If hf Then
            hasHeadingRow = True
            Exit For
        End If
    Next rw
    On Error GoTo 0

    TableHasHeadingRow = hasHeadingRow
End Function


' ---------------------------------------------------------------------------
' Apply the correct paragraph style to every cell in a row.
' ---------------------------------------------------------------------------
Private Sub StyleRowCells(rw As Row, doc As Document, isHeading As Boolean)
    Dim cel As Cell

    For Each cel In rw.Cells
        On Error Resume Next
        cel.VerticalAlignment = wdCellAlignVerticalTop
        On Error GoTo 0

        If isHeading Then
            ApplyCellStyle cel, "Table Heading", doc
        Else
            ApplyCellStyle cel, "Table Text", doc, KEEP_HEADING_STYLE:=True
        End If
    Next cel
End Sub


' ---------------------------------------------------------------------------
' Apply a paragraph style to every paragraph inside a cell, preserving
' existing alignment where no override is needed.
' KEEP_HEADING_STYLE avoids downgrading explicit Table Heading paragraphs that
' may appear inside a non-heading row.
' ---------------------------------------------------------------------------
Private Sub ApplyCellStyle(cel As Cell, styleName As String, doc As Document, _
                           Optional alignment As Long = -999, _
                           Optional KEEP_HEADING_STYLE As Boolean = False)
    Dim para As Paragraph
    Dim rng As Range
    Dim targetStyle As String
    Dim currentAlignment As Long

    On Error Resume Next
    For Each para In cel.Range.Paragraphs
        Set rng = para.Range
        ' Trim the end-of-cell marker from the range
        If rng.End > rng.Start + 1 Then
            Set rng = doc.Range(rng.Start, rng.End - 1)
        End If
        currentAlignment = rng.ParagraphFormat.Alignment
        targetStyle = styleName
        If KEEP_HEADING_STYLE Then
            If GetStyleNameSafe(rng) = "Table Heading" Then
                targetStyle = "Table Heading"
            End If
        End If
        rng.Style = doc.Styles(targetStyle)
        If alignment = -999 Then
            rng.ParagraphFormat.Alignment = currentAlignment
        Else
            rng.ParagraphFormat.Alignment = alignment
        End If
    Next para
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' If the last row is a merged caption/title row, style it as
' "Report Table Number", force top-left alignment, and leave only the top
' border visible.
' ---------------------------------------------------------------------------
Private Sub StyleCaptionRowInsideTable(tbl As Table, doc As Document)
    Dim capRow As Row
    Dim capCell As Cell
    Dim paraText As String
    Dim paraStyle As String

    On Error Resume Next
    Set capRow = tbl.Rows(tbl.Rows.Count)
    If capRow Is Nothing Then Exit Sub
    If capRow.Cells.Count <> 1 Then Exit Sub

    Set capCell = capRow.Cells(1)
    paraText = CleanRangeText(capCell.Range)
    paraStyle = GetStyleNameSafe(capCell.Range)
    If Not LooksLikeCaptionText(paraText, paraStyle) Then Exit Sub

    capCell.VerticalAlignment = wdCellAlignVerticalTop
    ApplyCellStyle capCell, "Report Table Number", doc, wdAlignParagraphLeft
    ApplyTopOnlyBordersToCell capCell
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' Look at the first non-empty paragraph immediately following a table.
' If it is a table caption/title, style it as "Report Table Number" and
' force left alignment with only the top border visible.
' ---------------------------------------------------------------------------
Private Sub StyleCaptionAfterTable(tbl As Table, doc As Document)
    Dim rng As Range
    Dim paraText As String
    Dim paraStyle As String
    Dim i As Long

    On Error Resume Next

    Set rng = doc.Range(tbl.Range.End, tbl.Range.End)
    For i = 1 To 3
        rng.Collapse wdCollapseEnd
        rng.Move Unit:=wdParagraph, Count:=1
        rng.Expand Unit:=wdParagraph

        paraText = CleanRangeText(rng)
        If Len(paraText) > 0 Then
            ' Check if already styled or text starts with "Table "
            paraStyle = GetStyleNameSafe(rng)

            If LooksLikeCaptionText(paraText, paraStyle) Then
                rng.Style = doc.Styles("Report Table Number")
                rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
                ApplyTopOnlyBordersToParagraphRange rng
            End If
            Exit For
        End If

        Set rng = doc.Range(rng.End, rng.End)
    Next i

    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' Return plain text for a range with control characters stripped.
' ---------------------------------------------------------------------------
Private Function CleanRangeText(rng As Range) As String
    Dim txt As String

    txt = rng.Text
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, Chr(7), "")
    CleanRangeText = Trim$(txt)
End Function


' ---------------------------------------------------------------------------
' Identify whether the supplied text/style pair represents a table caption.
' ---------------------------------------------------------------------------
Private Function LooksLikeCaptionText(ByVal paraText As String, ByVal paraStyle As String) As Boolean
    Dim probe As String

    probe = LCase$(Trim$(paraText))
    LooksLikeCaptionText = False
    If probe = "" Then Exit Function

    If Left$(probe, 6) = "table " Then
        LooksLikeCaptionText = True
        Exit Function
    End If

    If paraStyle = "Report Table Number" Then
        LooksLikeCaptionText = True
    End If
End Function


' ---------------------------------------------------------------------------
' Safe style-name lookup for a range.
' ---------------------------------------------------------------------------
Private Function GetStyleNameSafe(rng As Range) As String
    On Error Resume Next
    GetStyleNameSafe = rng.Style.NameLocal
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' Keep only the top border visible for caption/title rows.
' ---------------------------------------------------------------------------
Private Sub ApplyTopOnlyBordersToCell(cel As Cell)
    On Error Resume Next
    cel.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    cel.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    cel.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    cel.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' Keep only the top border visible for caption/title paragraphs.
' ---------------------------------------------------------------------------
Private Sub ApplyTopOnlyBordersToParagraphRange(rng As Range)
    On Error Resume Next
    rng.ParagraphFormat.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    rng.ParagraphFormat.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    rng.ParagraphFormat.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    rng.ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    On Error GoTo 0
End Sub
