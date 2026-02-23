    ' JSON parsing constants
    Private Const PARSE_SUCCESS As Long = 0
    Private Const PARSE_UNEXPECTED_END_OF_INPUT As Long = 1
    Private Const PARSE_INVALID_CHARACTER As Long = 2
    Private Const PARSE_INVALID_JSON_TYPE As Long = 3
    Private Const PARSE_INVALID_NUMBER As Long = 4
    Private Const PARSE_INVALID_KEY As Long = 5
    Private Const PARSE_INVALID_ESCAPE_CHARACTER As Long = 6

    ' Feature flags
    '
    ' FEATURE FLAGS GUIDE (what they do + tradeoffs when disabling)
    '
    ' 1) USE_GRANULAR_DIFF
    '    Purpose: Apply minimal text edits for plain replacements (better tracked-change readability).
    '    Disable (False) pros: Faster on large replacements; simpler behavior.
    '    Disable (False) cons: Coarser tracked changes (larger delete/insert blocks).
    '
    ' 2) INCLUDE_TABLE_PARAGRAPHS_IN_DSM
    '    Purpose: Include paragraph entries that are inside tables in the DSM paragraph list.
    '    Disable (False) pros: Smaller/faster DSM; avoids noisy duplicate coverage of table content.
    '    Disable (False) cons: Cannot target in-table paragraphs by P# IDs; table edits should use T#.R#.C#.
    '
    ' 3) INCLUDE_DSM_LOCATION_METADATA
    '    Purpose: Populate section/page metadata for DSM elements.
    '    Disable (False) pros: Noticeably faster map/export (page/section lookups can be expensive).
    '    Disable (False) cons: No section/page hints in exported DSM JSON.
    '
    ' 4) DSM_INCLUDE_TABLE_CELL_TAGGED_TEXT
    '    Purpose: Export rich tagged text (text_tagged) for each table cell.
    '    Disable (False) pros: Large speed-up for big/merged tables; smaller JSON payloads.
    '    Disable (False) cons: Loses per-cell inline formatting tags in DSM output.
    '
    ' 5) DSM_INCLUDE_TABLE_CELL_FORMAT_SPANS
    '    Purpose: Export per-cell format span arrays (format_spans) in DSM.
    '    Disable (False) pros: Large speed-up for big/merged tables; less memory use.
    '    Disable (False) cons: Loses precise formatting span diagnostics for table cells.
    '
    ' Recommended performance profile for large documents:
    ' - INCLUDE_DSM_LOCATION_METADATA = False
    ' - DSM_INCLUDE_TABLE_CELL_TAGGED_TEXT = False
    ' - DSM_INCLUDE_TABLE_CELL_FORMAT_SPANS = False
    ' Keep USE_GRANULAR_DIFF = True unless replacement speed is the top priority.
    Private Const USE_GRANULAR_DIFF As Boolean = True
    Private Const INCLUDE_TABLE_PARAGRAPHS_IN_DSM As Boolean = False
    Private Const INCLUDE_DSM_LOCATION_METADATA As Boolean = False ' Section/page lookups are expensive on large docs
    ' DSM_*_MAX_* and DSM_INITIAL_CAPACITY below are tuning knobs, not on/off flags.
    ' Lower values generally improve speed/memory usage but reduce preview/detail depth.
    Private Const DSM_INITIAL_CAPACITY As Long = 256
    Private Const DSM_PARAGRAPH_PREVIEW_MAX_CHARS As Long = 99999
    Private Const DSM_TABLE_PREVIEW_MAX_ROWS As Long = 200
    Private Const DSM_TABLE_PREVIEW_MAX_COLS As Long = 50
    Private Const DSM_TABLE_CELL_PREVIEW_MAX_CHARS As Long = 99999
    Private Const DSM_INCLUDE_TABLE_CELL_TAGGED_TEXT As Boolean = False
    Private Const DSM_INCLUDE_TABLE_CELL_FORMAT_SPANS As Boolean = False
    ' Performance tuning for long-running loops/logging.
    Private Const VERBOSE_TRACE_LOGGING As Boolean = False
    Private Const DSM_PARAGRAPH_PROGRESS_INTERVAL As Long = 100
    Private Const DSM_TABLE_PROGRESS_INTERVAL As Long = 10
    Private Const DSM_CELL_PROGRESS_INTERVAL As Long = 100
    Private Const APPLY_PREFLIGHT_PROGRESS_INTERVAL As Long = 50

    ' Diff operation type constants
Private Const DIFF_EQUAL As String = "equal"
Private Const DIFF_INSERT As String = "insert"
Private Const DIFF_DELETE As String = "delete"

    ' Safety limits for loop protection
    Private Const LOOP_SAFETY_LIMIT As Long = 5000       ' Max iterations for search loops
    Private Const WAIT_LOOP_SAFETY_LIMIT As Long = 600000 ' Max iterations for UI wait loops (~10 min at 1ms/iter)

    ' Windows API for Sleep (used in wait loops to reduce CPU usage)
    #If VBA7 Then
        Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    #Else
        Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    #End If

    Private p_Json As String
    Private p_Index As Long
    Private p_ParseError As Long

' Table index for deterministic table identification
Private Type TableIndexEntry
    TableIndex As Long              ' 1-based position in document
    TableRef As Table               ' Reference to the actual table
    TitleBelow As String            ' Title paragraph below the table (cleaned)
    CaptionAbove As String          ' Caption/text above the table (cleaned)
    HeaderRowText As String         ' Concatenated first row content
    StartPos As Long                ' Character position of table start
    EndPos As Long                  ' Character position of table end
End Type

Private g_TableIndex() As TableIndexEntry
Private g_TableIndexCount As Long
Private g_TableIndexBuilt As Boolean
Private g_TocFieldRanges As Collection
Private g_TocFieldCacheBuilt As Boolean

' =========================================================================================
' === DOCUMENT STRUCTURE MAP (DSM) - V4 Architecture =====================================
' =========================================================================================

' Document element type for structure map
Private Type DocumentElement
    ElementID As String             ' P1, P2, T1, T2, etc.
    ElementType As String           ' "paragraph", "table"
    StyleName As String             ' Word style name
    TextPreview As String           ' First 100 chars of content
    StartPos As Long                ' Character position start
    EndPos As Long                  ' Character position end
    TableRowCount As Long           ' For tables: number of rows
    TableColCount As Long           ' For tables: number of columns
    AssociatedTitleID As String     ' For tables: ID of title paragraph (if any)
    SectionNumber As Long           ' Section index (Word numbering)
    PageNumber As Long              ' Page number (Final view)
    HeadingLevel As Long            ' Heading level for paragraphs (0 if not a heading)
    TableTitle As String            ' Paragraph below table (if any)
    TableCaption As String          ' Paragraph above table (if any)
    WithinTable As Boolean          ' True when paragraph is inside a table
End Type

' Global structure map
Private g_DocumentMap() As DocumentElement
Private g_DocumentMapCount As Long
Private g_DocumentMapBuilt As Boolean

' Tool registry cache
Private g_ToolRegistry As Object
Private g_ToolRegistryOrder As Collection

' Last run metadata for COM/automation callers
Private g_LastRunResultJson As String
Private g_LastRunReportPath As String
Private g_LastRunId As String
Private g_LastActionErrorCode As String
Private g_LastActionErrorMessage As String

' Target reference type for V4 architecture
Private Type TargetReference
    IsValid As Boolean
    TargetType As String        ' "paragraph", "table", "table_row", "table_cell"
    ParagraphNum As Long        ' For P5
    TableNum As Long            ' For T2, T2.R3, T2.R3.C2
    RowNum As Long              ' For T2.R3, T2.R3.C2
    ColNum As Long              ' For T2.R3.C2
    IsHeaderRow As Boolean      ' For T2.H.C2
End Type

Option Explicit

' VBA Module to Apply LLM Review Suggestions (Version 3 - Production)
'
' REFINEMENTS:
' - Sets a distinct "AI Reviewer" identity for all changes.
' - Anchors comments to the full context for better visibility.
' - Pre-processes JSON to handle common errors (smart quotes, trailing commas).
' - UI now provides a progress indicator and validation.
' - Case sensitivity is now a user-configurable option.

' =========================================================================================
' === TABLE INDEX FUNCTIONS ===============================================================
' =========================================================================================

Private Sub BuildTableIndex(ByVal doc As Document)
    ' Builds a global index of all tables in the document for fast lookup
    ' Each table is indexed by its title (paragraph below) and caption (paragraph above)

    On Error GoTo ErrorHandler

    Dim tbl As Table
    Dim idx As Long
    Dim tableCount As Long

    tableCount = doc.Tables.Count
    If tableCount = 0 Then
        g_TableIndexCount = 0
        g_TableIndexBuilt = True
        Exit Sub
    End If

    ReDim g_TableIndex(1 To tableCount)
    idx = 0

    Debug.Print "Building table index for " & tableCount & " tables..."

    For Each tbl In doc.Tables
        idx = idx + 1

        With g_TableIndex(idx)
            .TableIndex = idx
            Set .TableRef = tbl
            .StartPos = tbl.Range.Start
            .EndPos = tbl.Range.End

            ' Get title below the table (the paragraph immediately after)
            .TitleBelow = GetTableTitleBelow(tbl)

            ' Get caption above the table (the paragraph immediately before)
            .CaptionAbove = GetTableCaptionAbove(tbl)

            ' Get first row content for additional disambiguation
            .HeaderRowText = GetTableHeaderRowText(tbl)

            Debug.Print "  Table " & idx & ":"
            If Len(.TitleBelow) > 0 Then Debug.Print "    Title below: " & Left$(.TitleBelow, 80)
            If Len(.CaptionAbove) > 0 Then Debug.Print "    Caption above: " & Left$(.CaptionAbove, 80)
        End With
    Next tbl

    g_TableIndexCount = idx
    g_TableIndexBuilt = True
    Debug.Print "Table index built: " & g_TableIndexCount & " tables indexed"
    Exit Sub

ErrorHandler:
    Debug.Print "Error building table index: " & Err.Description
    g_TableIndexCount = 0
    g_TableIndexBuilt = False
End Sub

Private Function GetTableTitleBelow(ByVal tbl As Table) As String
    ' Gets the title paragraph that appears BELOW the table
    ' Returns cleaned text or empty string if not found

    On Error GoTo Fail

    Dim p As Paragraph
    Dim t As String
    Dim i As Long

    ' Look at the paragraphs after the table (up to 3)
    Set p = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count).Next
    For i = 1 To 3
        If p Is Nothing Then Exit For

        ' Skip if this paragraph is inside another table
        If p.Range.Information(wdWithinTable) Then
            Set p = p.Next
            GoTo ContinueLoop
        End If

        t = Trim$(NormalizeForDocument(GetParagraphTextFinalView(p)))

        ' Skip empty paragraphs and very short ones
        If Len(t) > 3 Then
            ' Skip if it looks like an auto-generated table number (e.g., "Table 7.4")
            If Not IsAutoTableNumber(t) Then
                GetTableTitleBelow = t
                Exit Function
            End If
        End If

ContinueLoop:
        Set p = p.Next
    Next i

    GetTableTitleBelow = ""
    Exit Function

Fail:
    GetTableTitleBelow = ""
End Function

Private Function GetTableCaptionAbove(ByVal tbl As Table) As String
    ' Gets the caption paragraph that appears ABOVE the table
    ' Returns cleaned text or empty string if not found

    On Error GoTo Fail

    Dim p As Paragraph
    Dim t As String
    Dim i As Long

    ' Look at the paragraphs before the table (up to 3)
    Set p = tbl.Range.Paragraphs(1).Previous
    For i = 1 To 3
        If p Is Nothing Then Exit For

        ' Skip if this paragraph is inside another table
        If p.Range.Information(wdWithinTable) Then
            Set p = p.Previous
            GoTo ContinueLoop
        End If

        t = Trim$(NormalizeForDocument(GetParagraphTextFinalView(p)))

        ' Skip empty paragraphs
        If Len(t) > 3 Then
            GetTableCaptionAbove = t
            Exit Function
        End If

ContinueLoop:
        Set p = p.Previous
    Next i

    GetTableCaptionAbove = ""
    Exit Function

Fail:
    GetTableCaptionAbove = ""
End Function

Private Function GetTableHeaderRowText(ByVal tbl As Table) As String
    ' Gets concatenated text from the first row of the table
    ' Used for additional disambiguation when titles match

    On Error GoTo Fail

    Dim c As Cell
    Dim result As String
    Dim hasRow1Cell As Boolean

    result = ""

    For Each c In tbl.Range.Cells
        If c.RowIndex = 1 Then
            hasRow1Cell = True
            result = result & " " & Trim$(NormalizeForDocument(GetCellTextFinalView(c)))
        End If
    Next c

    If Not hasRow1Cell Then
        GetTableHeaderRowText = ""
        Exit Function
    End If

    GetTableHeaderRowText = Trim$(result)
    Exit Function

Fail:
    GetTableHeaderRowText = ""
End Function

Private Function IsAutoTableNumber(ByVal text As String) As Boolean
    ' Checks if text looks like an auto-generated table number (e.g., "Table 7.4")
    ' These are Word fields that may not be visible to the LLM

    Dim t As String
    t = LCase$(Trim$(text))

    ' Pattern: "table" followed by numbers/dots
    If Left$(t, 5) = "table" Then
        Dim rest As String
        rest = Trim$(Mid$(t, 6))
        ' Check if the rest is mostly numbers, dots, and dashes
        If Len(rest) <= 15 Then
            Dim i As Long
            Dim ch As String
            Dim hasDigit As Boolean
            For i = 1 To Len(rest)
                ch = Mid$(rest, i, 1)
                If ch Like "[0-9]" Then hasDigit = True
                If Not (ch Like "[0-9. -–:]") Then
                    IsAutoTableNumber = False
                    Exit Function
                End If
            Next i
            If hasDigit Then
                IsAutoTableNumber = True
                Exit Function
            End If
        End If
    End If

    IsAutoTableNumber = False
End Function

Private Function FindTableByTitle(ByVal tableTitle As String, ByRef matchCount As Long) As Table
    ' Finds a table by its title (paragraph below)
    ' Returns the table if exactly one match, Nothing if zero or ambiguous
    ' matchCount is set to the number of matching tables found

    On Error GoTo Fail

    If Not g_TableIndexBuilt Then
        BuildTableIndex ActiveDocument
    End If

    Dim i As Long
    Dim matches As New Collection
    Dim normalizedTitle As String

    normalizedTitle = NormalizeForDocument(tableTitle)

    ' First pass: exact match on title below
    For i = 1 To g_TableIndexCount
        If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).TitleBelow) Then
            matches.Add i
        End If
    Next i

    ' If no matches, try caption above
    If matches.Count = 0 Then
        For i = 1 To g_TableIndexCount
            If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).CaptionAbove) Then
                matches.Add i
            End If
        Next i
    End If

    matchCount = matches.Count

    If matches.Count = 1 Then
        Set FindTableByTitle = g_TableIndex(matches(1)).TableRef
    ElseIf matches.Count > 1 Then
        Debug.Print "  -> FindTableByTitle: " & matches.Count & " tables match title '" & Left$(tableTitle, 50) & "'"
        Set FindTableByTitle = Nothing
    Else
        Set FindTableByTitle = Nothing
    End If
    Exit Function

Fail:
    matchCount = 0
    Set FindTableByTitle = Nothing
End Function

Private Function GetMatchingTableCandidates(ByVal tableTitle As String) As Collection
    ' Returns a collection of TableIndexEntry indices that match the given title
    ' Used when multiple tables match and user selection is needed

    On Error GoTo Fail

    If Not g_TableIndexBuilt Then
        BuildTableIndex ActiveDocument
    End If

    Dim i As Long
    Dim matches As New Collection
    Dim normalizedTitle As String

    normalizedTitle = NormalizeForDocument(tableTitle)

    ' Check title below
    For i = 1 To g_TableIndexCount
        If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).TitleBelow) Then
            matches.Add i
        End If
    Next i

    ' If no matches from title below, try caption above
    If matches.Count = 0 Then
        For i = 1 To g_TableIndexCount
            If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).CaptionAbove) Then
                matches.Add i
            End If
        Next i
    End If

    Set GetMatchingTableCandidates = matches
    Exit Function

Fail:
    Set GetMatchingTableCandidates = New Collection
End Function

Private Sub ClearTableIndex()
    ' Clears the table index - call when document changes
    g_TableIndexCount = 0
    g_TableIndexBuilt = False
    Erase g_TableIndex
End Sub

Private Sub TraceLog(ByVal message As String)
    If VERBOSE_TRACE_LOGGING Then Debug.Print message
End Sub

Private Sub ClearTocFieldRangeCache()
    Set g_TocFieldRanges = Nothing
    g_TocFieldCacheBuilt = False
End Sub

Private Sub EnsureTocFieldRangeCache(ByVal doc As Document)
    On Error GoTo Fail

    If g_TocFieldCacheBuilt Then Exit Sub

    Dim fld As field
    Dim fldRange As Range

    Set g_TocFieldRanges = New Collection
    For Each fld In doc.Fields
        If fld.Type = wdFieldTOC Then
            Set fldRange = fld.Result
            If Not fldRange Is Nothing Then g_TocFieldRanges.Add fldRange.Duplicate
        End If
    Next fld

    g_TocFieldCacheBuilt = True
    Exit Sub

Fail:
    Set g_TocFieldRanges = New Collection
    g_TocFieldCacheBuilt = True
End Sub

' =========================================================================================
' === FINAL VIEW TEXT EXTRACTION (for tracked changes handling) =============================
' =========================================================================================

Private Function GetParagraphTextFinalView(ByVal para As Paragraph) As String
    ' FAST PATH: if the document has no revisions, return paragraph text immediately.
    ' SLOW PATH: only normalize tracked changes when revisions exist.

    On Error GoTo Fail

    If para.Range.Document.Revisions.Count = 0 Then
        GetParagraphTextFinalView = Trim$(para.Range.Text)
        Exit Function
    End If

    Dim rng As Range
    Set rng = para.Range
    GetParagraphTextFinalView = Trim$(GetRangeTextFinalView(rng))
    Exit Function

Fail:
    GetParagraphTextFinalView = Trim$(para.Range.Text)
End Function

Private Function GetRangeTextFinalView(ByVal rng As Range) As String
    ' Extracts range text as shown in "Final" view:
    ' includes insertions, excludes deletions.

    On Error GoTo Fail

    If rng Is Nothing Then
        GetRangeTextFinalView = ""
        Exit Function
    End If

    Dim workRange As Range
    Set workRange = rng.Duplicate

    If workRange.Revisions.Count = 0 Then
        GetRangeTextFinalView = workRange.Text
    Else
        GetRangeTextFinalView = GetTextExcludingDeletions(workRange)
    End If
    Exit Function

Fail:
    If rng Is Nothing Then
        GetRangeTextFinalView = ""
    Else
        GetRangeTextFinalView = rng.Text
    End If
End Function

Private Function GetTextExcludingDeletions(ByVal rng As Range) As String
    ' Helper to extract text from a range, excluding any deleted text.

    On Error GoTo Fail

    If rng Is Nothing Then
        GetTextExcludingDeletions = ""
        Exit Function
    End If

    Dim rev As Revision
    Dim delCount As Long
    Dim delStarts() As Long
    Dim delEnds() As Long
    Dim delStart As Long
    Dim delEnd As Long

    delCount = 0

    ' Collect deletion ranges overlapping this range and clamp to bounds.
    For Each rev In rng.Revisions
        If rev.Type = wdRevisionDelete Then
            delStart = rev.Range.Start
            delEnd = rev.Range.End

            If delStart < rng.Start Then delStart = rng.Start
            If delEnd > rng.End Then delEnd = rng.End

            If delEnd > delStart Then
                delCount = delCount + 1
                ReDim Preserve delStarts(1 To delCount)
                ReDim Preserve delEnds(1 To delCount)
                delStarts(delCount) = delStart
                delEnds(delCount) = delEnd
            End If
        End If
    Next rev

    If delCount = 0 Then
        GetTextExcludingDeletions = rng.Text
        Exit Function
    End If

    ' Sort deletion ranges by start position (insertion sort).
    Dim i As Long
    Dim j As Long
    Dim keyStart As Long
    Dim keyEnd As Long

    For i = 2 To delCount
        keyStart = delStarts(i)
        keyEnd = delEnds(i)
        j = i - 1
        Do While j >= 1
            If delStarts(j) > keyStart Then
                delStarts(j + 1) = delStarts(j)
                delEnds(j + 1) = delEnds(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        delStarts(j + 1) = keyStart
        delEnds(j + 1) = keyEnd
    Next i

    Dim doc As Document
    Set doc = rng.Document

    Dim result As String
    Dim cursorPos As Long
    Dim mergedStart As Long
    Dim mergedEnd As Long
    Dim subRng As Range

    result = ""
    cursorPos = rng.Start
    mergedStart = delStarts(1)
    mergedEnd = delEnds(1)

    For i = 2 To delCount + 1
        If i <= delCount Then
            If delStarts(i) <= mergedEnd Then
                If delEnds(i) > mergedEnd Then mergedEnd = delEnds(i)
            Else
                If mergedStart > cursorPos Then
                    Set subRng = doc.Range(cursorPos, mergedStart)
                    result = result & subRng.Text
                End If
                cursorPos = mergedEnd
                mergedStart = delStarts(i)
                mergedEnd = delEnds(i)
            End If
        Else
            If mergedStart > cursorPos Then
                Set subRng = doc.Range(cursorPos, mergedStart)
                result = result & subRng.Text
            End If
            cursorPos = mergedEnd
        End If
    Next i

    If cursorPos < rng.End Then
        Set subRng = doc.Range(cursorPos, rng.End)
        result = result & subRng.Text
    End If

    GetTextExcludingDeletions = result
    Exit Function

Fail:
    If rng Is Nothing Then
        GetTextExcludingDeletions = ""
    Else
        GetTextExcludingDeletions = rng.Text
    End If
End Function

Private Function StripCellEndMarkers(ByVal cellText As String) As String
    ' Removes trailing Word table-cell end markers (Chr(13), Chr(7)).

    Dim t As String
    t = cellText

    Do While Len(t) > 0
        If Right$(t, 1) = Chr$(7) Or Right$(t, 1) = Chr$(13) Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop

    StripCellEndMarkers = t
End Function

Private Function GetCellTextFinalView(ByVal cell As Cell) As String
    ' Extracts cell text as it appears in Word's "Final" view.
    ' - Insertions (w:ins) are INCLUDED
    ' - Deletions (w:del) are EXCLUDED
    ' This matches Word's "Final" display mode without actually accepting changes.
    
    On Error GoTo Fail

    Dim rng As Range
    Dim finalText As String

    Set rng = cell.Range
    finalText = GetRangeTextFinalView(rng)
    finalText = StripCellEndMarkers(finalText)
    GetCellTextFinalView = Trim$(finalText)
    Exit Function

Fail:
    GetCellTextFinalView = ""
End Function

Private Function GetSafeTableRowCount(ByVal tbl As Table) As Long
    ' Returns row count without failing on vertically merged tables.

    On Error GoTo Fallback
    GetSafeTableRowCount = tbl.Rows.Count
    Exit Function

Fallback:
    Err.Clear
    On Error GoTo Fail

    Dim c As Cell
    Dim maxRow As Long
    maxRow = 0

    For Each c In tbl.Range.Cells
        If c.RowIndex > maxRow Then maxRow = c.RowIndex
    Next c

    GetSafeTableRowCount = maxRow
    Exit Function

Fail:
    GetSafeTableRowCount = 0
End Function

Private Function GetSafeTableColCount(ByVal tbl As Table) As Long
    ' Returns column count without failing on merged-cell tables.

    On Error GoTo Fallback
    GetSafeTableColCount = tbl.Columns.Count
    Exit Function

Fallback:
    Err.Clear
    On Error GoTo Fail

    Dim c As Cell
    Dim maxCol As Long
    maxCol = 0

    For Each c In tbl.Range.Cells
        If c.ColumnIndex > maxCol Then maxCol = c.ColumnIndex
    Next c

    GetSafeTableColCount = maxCol
    Exit Function

Fail:
    GetSafeTableColCount = 0
End Function

Private Function TryGetTableCell(ByVal tbl As Table, ByVal rowNum As Long, ByVal colNum As Long, ByRef outCell As Cell) As Boolean
    ' Tries direct row/column addressing first, then falls back to merged-cell tolerant lookup.

    On Error GoTo Fail
    Set outCell = Nothing

    If rowNum <= 0 Or colNum <= 0 Then
        TryGetTableCell = False
        Exit Function
    End If

    On Error Resume Next
    Set outCell = tbl.Cell(rowNum, colNum)
    If Err.Number = 0 And Not outCell Is Nothing Then
        On Error GoTo Fail
        TryGetTableCell = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo Fail

    Dim c As Cell
    Dim rowFallback As Cell
    Dim rowFallbackCol As Long
    Dim colFallback As Cell
    Dim colFallbackRow As Long

    rowFallbackCol = 0
    colFallbackRow = 0

    For Each c In tbl.Range.Cells
        If c.RowIndex = rowNum Then
            If c.ColumnIndex = colNum Then
                Set outCell = c
                TryGetTableCell = True
                Exit Function
            End If

            ' Horizontal merge fallback: choose the closest anchor cell on the same row.
            If c.ColumnIndex <= colNum And c.ColumnIndex > rowFallbackCol Then
                Set rowFallback = c
                rowFallbackCol = c.ColumnIndex
            End If
        End If

        ' Vertical merge fallback: choose the closest anchor cell in the same column above.
        If c.ColumnIndex = colNum Then
            If c.RowIndex <= rowNum And c.RowIndex > colFallbackRow Then
                Set colFallback = c
                colFallbackRow = c.RowIndex
            End If
        End If
    Next c

    If Not rowFallback Is Nothing Then
        Set outCell = rowFallback
        TryGetTableCell = True
        Exit Function
    End If

    If Not colFallback Is Nothing Then
        Set outCell = colFallback
        TryGetTableCell = True
        Exit Function
    End If

    TryGetTableCell = False
    Exit Function

Fail:
    Set outCell = Nothing
    TryGetTableCell = False
End Function

Private Function TryGetTableRowRange(ByVal tbl As Table, ByVal rowNum As Long, ByRef outRange As Range) As Boolean
    ' Gets a row range with fallback for vertically merged tables where tbl.Rows(rowNum) may fail.

    On Error GoTo Fail
    Set outRange = Nothing

    If rowNum <= 0 Then
        TryGetTableRowRange = False
        Exit Function
    End If

    On Error Resume Next
    Set outRange = tbl.Rows(rowNum).Range
    If Err.Number = 0 And Not outRange Is Nothing Then
        On Error GoTo Fail
        TryGetTableRowRange = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo Fail

    Dim c As Cell
    Dim minStart As Long
    Dim maxEnd As Long
    minStart = 0
    maxEnd = 0

    For Each c In tbl.Range.Cells
        If c.RowIndex = rowNum Then
            If minStart = 0 Or c.Range.Start < minStart Then minStart = c.Range.Start
            If c.Range.End > maxEnd Then maxEnd = c.Range.End
        End If
    Next c

    If minStart > 0 And maxEnd > minStart Then
        Set outRange = tbl.Range.Document.Range(minStart, maxEnd)
        TryGetTableRowRange = True
    Else
        TryGetTableRowRange = False
    End If
    Exit Function

Fail:
    Set outRange = Nothing
    TryGetTableRowRange = False
End Function

Private Function TryGetFirstCellInRow(ByVal tbl As Table, ByVal rowNum As Long, ByRef outCell As Cell) As Boolean
    ' Returns the left-most anchor cell on a logical row.

    On Error GoTo Fail
    Set outCell = Nothing

    If rowNum <= 0 Then
        TryGetFirstCellInRow = False
        Exit Function
    End If

    Dim c As Cell
    Dim minCol As Long
    minCol = 0

    For Each c In tbl.Range.Cells
        If c.RowIndex = rowNum Then
            If outCell Is Nothing Then
                Set outCell = c
                minCol = c.ColumnIndex
            ElseIf c.ColumnIndex < minCol Then
                Set outCell = c
                minCol = c.ColumnIndex
            End If
        End If
    Next c

    TryGetFirstCellInRow = Not outCell Is Nothing
    Exit Function

Fail:
    Set outCell = Nothing
    TryGetFirstCellInRow = False
End Function

Private Function GetCellContentRange(ByVal c As Cell) As Range
    ' Returns a duplicate cell range without trailing cell/paragraph markers.

    On Error GoTo Fail

    Dim rng As Range
    Set rng = c.Range.Duplicate

    Do While rng.End > rng.Start
        Dim ch As String
        ch = Right$(rng.Text, 1)
        If ch = Chr$(7) Or ch = Chr$(13) Then
            rng.End = rng.End - 1
        Else
            Exit Do
        End If
    Loop

    Set GetCellContentRange = rng
    Exit Function

Fail:
    Set GetCellContentRange = Nothing
End Function

' =========================================================================================
' === DOCUMENT STRUCTURE MAP (DSM) FUNCTIONS =============================================
' =========================================================================================

Private Sub BuildDocumentStructureMap(ByVal doc As Document)
    ' Builds a comprehensive structure map of the document
    On Error GoTo ErrorHandler
    
    Dim idx As Long
    Dim para As Paragraph
    Dim paraRange As Range
    Dim tbl As Table
    Dim tblRange As Range
    Dim styleStr As String
    Dim textPrev As String
    Dim paraIdx As Long
    Dim tblIdx As Long
    Dim inTable As Boolean
    Dim totalParagraphs As Long
    Dim totalTables As Long
    Dim docHasRevisions As Boolean
    
    totalParagraphs = doc.Paragraphs.Count
    totalTables = doc.Tables.Count
    docHasRevisions = False
    On Error Resume Next
    docHasRevisions = (doc.Revisions.Count > 0)
    Err.Clear
    On Error GoTo ErrorHandler

    Debug.Print "Building Document Structure Map..."
    Debug.Print "Total Paragraphs to process: " & totalParagraphs

    ' Allocate with growth headroom
    ReDim g_DocumentMap(1 To DSM_INITIAL_CAPACITY)
    idx = 0
    
    ' Map paragraphs
    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        If paraIdx Mod DSM_PARAGRAPH_PROGRESS_INTERVAL = 0 Then
            TraceLog "  -> DSM paragraph progress: " & paraIdx & "/" & totalParagraphs
            DoEvents
        End If

        Set paraRange = para.Range
        inTable = paraRange.Information(wdWithinTable)
        If (Not INCLUDE_TABLE_PARAGRAPHS_IN_DSM) And inTable Then
            GoTo ContinueParagraphLoop
        End If

        idx = idx + 1
        EnsureDocumentMapCapacity idx
        
        With g_DocumentMap(idx)
            .ElementID = "P" & paraIdx
            .ElementType = "paragraph"
            
            ' Get style name
            On Error Resume Next
            styleStr = para.Style.NameLocal
            If Err.Number <> 0 Then styleStr = "Normal"
            Err.Clear
            On Error GoTo ErrorHandler
            .StyleName = styleStr
            
            ' Get full text content in 'Final' view
            If docHasRevisions Then
                textPrev = Trim$(GetRangeTextFinalView(paraRange))
            Else
                textPrev = Trim$(paraRange.Text)
            End If
            .TextPreview = LimitPreviewText(textPrev, DSM_PARAGRAPH_PREVIEW_MAX_CHARS)
            
            ' Store position
            .StartPos = paraRange.Start
            .EndPos = paraRange.End
            
            .TableRowCount = 0
            .TableColCount = 0
            .AssociatedTitleID = ""
            If INCLUDE_DSM_LOCATION_METADATA Then
                .SectionNumber = GetRangeSectionNumber(paraRange)
                .PageNumber = GetRangePageNumber(paraRange)
            Else
                .SectionNumber = 0
                .PageNumber = 0
            End If
            .HeadingLevel = GetHeadingLevelFromStyle(styleStr)
            .TableTitle = ""
            .TableCaption = ""
            .WithinTable = inTable
        End With
        
ContinueParagraphLoop:
    Next para
    
    Debug.Print "Paragraphs complete. Total Tables to process: " & totalTables
    
    ' Map tables
    tblIdx = 0
    For Each tbl In doc.Tables
        tblIdx = tblIdx + 1
        idx = idx + 1

        If tblIdx Mod DSM_TABLE_PROGRESS_INTERVAL = 0 Then
            TraceLog "  -> DSM table progress: " & tblIdx & "/" & totalTables
            DoEvents
        End If
        
        EnsureDocumentMapCapacity idx
        Set tblRange = tbl.Range
        
        With g_DocumentMap(idx)
            .ElementID = "T" & tblIdx
            .ElementType = "table"
            .StyleName = ""
            
            textPrev = GetTableTitleForDSM(tbl, doc)
            .TextPreview = LimitPreviewText(textPrev, DSM_PARAGRAPH_PREVIEW_MAX_CHARS)
            
            .StartPos = tblRange.Start
            .EndPos = tblRange.End
            
            .TableRowCount = GetSafeTableRowCount(tbl)
            .TableColCount = GetSafeTableColCount(tbl)
            
            .AssociatedTitleID = ""
            .SectionNumber = 0
            .PageNumber = 0
            .HeadingLevel = 0
            
            .TableTitle = GetTableTitleBelow(tbl)
            .TableCaption = GetTableCaptionAbove(tbl)
            .WithinTable = False
        End With
    Next tbl

    If idx = 0 Then
        g_DocumentMapCount = 0
        g_DocumentMapBuilt = True
        Exit Sub
    End If

    TraceLog "Sorting map by position..."
    ReDim Preserve g_DocumentMap(1 To idx)
    g_DocumentMapCount = idx
    SortDocumentMapByStartPos
    g_DocumentMapBuilt = True
    
    Debug.Print "  -> Structure map built: " & paraIdx & " paragraphs, " & tblIdx & " tables"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error building structure map: " & Err.Description & " (idx=" & idx & ", paraIdx=" & paraIdx & ", tblIdx=" & tblIdx & ")"
    g_DocumentMapCount = 0
    g_DocumentMapBuilt = False
End Sub

Private Sub EnsureDocumentMapCapacity(ByVal requiredIndex As Long)
    Dim currentUpper As Long
    Dim newUpper As Long

    If requiredIndex <= 0 Then Exit Sub

    currentUpper = UBound(g_DocumentMap)
    If requiredIndex <= currentUpper Then Exit Sub

    newUpper = currentUpper
    Do While newUpper < requiredIndex
        If newUpper < 1024 Then
            newUpper = newUpper * 2
        Else
            newUpper = newUpper + 1024
        End If
    Loop

    ReDim Preserve g_DocumentMap(1 To newUpper)
End Sub

Private Sub SortDocumentMapByStartPos()
    Dim i As Long
    Dim j As Long
    Dim tmp As DocumentElement

    If g_DocumentMapCount <= 1 Then Exit Sub

    For i = 2 To g_DocumentMapCount
        tmp = g_DocumentMap(i)
        j = i - 1
        Do While j >= 1
            If g_DocumentMap(j).StartPos > tmp.StartPos Then
                g_DocumentMap(j + 1) = g_DocumentMap(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        g_DocumentMap(j + 1) = tmp
    Next i
End Sub

Private Function GetRangeSectionNumber(ByVal rng As Range) As Long
    On Error GoTo Fail
    If rng Is Nothing Then Exit Function
    GetRangeSectionNumber = rng.Information(wdActiveEndSectionNumber)
    Exit Function
Fail:
    GetRangeSectionNumber = 0
End Function

Private Function GetRangePageNumber(ByVal rng As Range) As Long
    On Error GoTo Fail
    If rng Is Nothing Then Exit Function
    GetRangePageNumber = rng.Information(wdActiveEndAdjustedPageNumber)
    Exit Function
Fail:
    GetRangePageNumber = 0
End Function

Private Function GetHeadingLevelFromStyle(ByVal styleName As String) As Long
    Dim normalized As String
    Dim suffix As String
    normalized = LCase$(Trim$(styleName))

    If Left$(normalized, 8) = "heading " Then
        suffix = Trim$(Mid$(normalized, 9))
    ElseIf Left$(normalized, 13) = "report level " Then
        suffix = Trim$(Mid$(normalized, 14))
    Else
        GetHeadingLevelFromStyle = 0
        Exit Function
    End If

    If Len(suffix) > 0 And IsNumeric(suffix) Then
        GetHeadingLevelFromStyle = CLng(Val(suffix))
    Else
        GetHeadingLevelFromStyle = 0
    End If
End Function

Private Function LimitPreviewText(ByVal text As String, ByVal maxChars As Long) As String
    Dim normalized As String
    normalized = Replace(text, vbCr, " ")
    normalized = Replace(normalized, vbLf, " ")
    normalized = Trim$(normalized)

    If maxChars <= 0 Then
        LimitPreviewText = normalized
        Exit Function
    End If

    If Len(normalized) > maxChars Then
        LimitPreviewText = Left$(normalized, maxChars) & "..."
    Else
        LimitPreviewText = normalized
    End If
End Function

Private Function ResolveBodyParagraphRange(ByVal paragraphNum As Long) As Range
    On Error GoTo Fail

    Dim para As Paragraph

    If paragraphNum <= 0 Then GoTo Fail
    If paragraphNum > ActiveDocument.Paragraphs.Count Then GoTo Fail

    Set para = ActiveDocument.Paragraphs(paragraphNum)
    If para Is Nothing Then GoTo Fail

    If (Not INCLUDE_TABLE_PARAGRAPHS_IN_DSM) And para.Range.Information(wdWithinTable) Then GoTo Fail

    Set ResolveBodyParagraphRange = para.Range.Duplicate
    Exit Function

Fail:
    Set ResolveBodyParagraphRange = Nothing
End Function

Private Function GetTableTitleForDSM(ByVal tbl As Table, ByVal doc As Document) As String
    ' Gets a descriptive title for the table (paragraph before or after)
    On Error Resume Next
    
    Dim p As Paragraph
    Dim txt As String
    
    ' Try paragraph after table first
    Set p = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count).Next
    If Not p Is Nothing Then
        txt = GetParagraphTextFinalView(p)
        If Len(txt) > 3 And Len(txt) < 200 Then
            GetTableTitleForDSM = txt
            Exit Function
        End If
    End If
    
    ' Try paragraph before table
    Set p = tbl.Range.Paragraphs(1).Previous
    If Not p Is Nothing Then
        txt = GetParagraphTextFinalView(p)
        If Len(txt) > 3 And Len(txt) < 200 Then
            GetTableTitleForDSM = txt
            Exit Function
        End If
    End If
    
    ' Fallback: use first logical cell content in Final view.
    Dim firstCell As Cell
    If TryGetTableCell(tbl, 1, 1, firstCell) Then
        txt = GetCellTextFinalView(firstCell)
        If Len(txt) > 100 Then txt = Left$(txt, 100) & "..."
        GetTableTitleForDSM = txt
        Exit Function
    End If
    
    GetTableTitleForDSM = "(Table)"
End Function

Private Sub ClearDocumentStructureMap()
    ' Clears the structure map - call when document changes
    g_DocumentMapCount = 0
    g_DocumentMapBuilt = False
    Erase g_DocumentMap
    ClearTocFieldRangeCache
End Sub

Private Function ExportStructureMapAsMarkdown() As String
    ' Exports the structure map as annotated markdown for the V5 search/replace workflow.
    
    On Error GoTo ErrorHandler
    
    If Not g_DocumentMapBuilt Then
        BuildDocumentStructureMap ActiveDocument
    End If
    
    If g_DocumentMapCount = 0 Then
        ExportStructureMapAsMarkdown = "# DOCUMENT STRUCTURE MAP" & vbCrLf & vbCrLf & "(Empty document)"
        Exit Function
    End If
    
    Dim output As String
    Dim i As Long
    Dim elem As DocumentElement
    Dim paraRange As Range
    Dim paraText As String
    
    output = output & "# DOCUMENT STRUCTURE MAP" & vbCrLf
    output = output & "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    output = output & "## Annotated Content" & vbCrLf & vbCrLf
    
    For i = 1 To g_DocumentMapCount
        elem = g_DocumentMap(i)
        
        If elem.ElementType = "paragraph" Then
            Set paraRange = ActiveDocument.Range(elem.StartPos, elem.EndPos)
            paraText = Trim$(NormalizeForDocument(GetRangeTextFinalView(paraRange)))
            output = output & "[" & elem.ElementID & "] " & paraText & vbCrLf & vbCrLf
            
        ElseIf elem.ElementType = "table" Then
            output = output & "## Table " & elem.ElementID
            If Len(elem.TextPreview) > 0 Then
                output = output & ": " & NormalizeForDocument(elem.TextPreview)
            End If
            output = output & vbCrLf
            output = output & ExportTableContentPreview(elem.ElementID) & vbCrLf & vbCrLf
        End If
    Next i
    
    ExportStructureMapAsMarkdown = output
    Exit Function
    
ErrorHandler:
    ExportStructureMapAsMarkdown = "# ERROR" & vbCrLf & "Failed to export structure map: " & Err.Description
End Function

Private Sub ExportStructureMapToFile()
    ' Exports the structure map to a text file in the document's parent folder
    ' File is named: [DocumentName]_StructureMap.txt
    
    On Error GoTo ErrorHandler
    
    Dim markdown As String
    Dim filePath As String
    Dim docPath As String
    Dim docName As String
    Dim parentFolder As String
    Dim fso As Object
    Dim fileNum As Integer
    
    ' Check if document is saved
    If ActiveDocument.Path = "" Then
        MsgBox "Please save the document first before exporting the structure map.", vbExclamation, "Document Not Saved"
        Exit Sub
    End If
    
    ' Get document path and name
    docPath = ActiveDocument.FullName
    docName = ActiveDocument.Name
    
    ' Remove extension from document name
    If InStrRev(docName, ".") > 0 Then
        docName = Left$(docName, InStrRev(docName, ".") - 1)
    End If
    
    ' Build file path in parent folder
    parentFolder = ActiveDocument.Path
    filePath = parentFolder & "\" & docName & "_StructureMap.txt"
    
    ' Generate the markdown content
    markdown = ExportStructureMapAsMarkdown()
    
    ' Write to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, markdown
    Close #fileNum
    
    MsgBox "Structure map exported to:" & vbCrLf & filePath, vbInformation, "Export Complete"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting structure map: " & Err.Description, vbCritical, "Export Error"
End Sub

Private Function EscapeMarkdownCellPreview(ByVal cellValue As String) As String
    Dim normalized As String

    normalized = NormalizeForDocument(cellValue)
    normalized = Replace(normalized, vbCrLf, " ")
    normalized = Replace(normalized, vbCr, " ")
    normalized = Replace(normalized, vbLf, " ")
    normalized = Replace(normalized, "|", "\|")
    normalized = Trim$(normalized)

    EscapeMarkdownCellPreview = normalized
End Function

Private Function ExportTableContentPreview(ByVal tableID As String) As String
    ' Exports table content in markdown format for preview
    ' Uses bounded previews to keep DSM generation responsive on large tables.
    
    On Error GoTo Fail
    
    Dim tbl As Table
    Dim r As Long
    Dim c As Long
    Dim cellText As String
    Dim output As String
    Dim maxRows As Long
    Dim maxCols As Long
    Dim actualRows As Long
    Dim actualCols As Long
    Dim cellLookup As Object
    Dim tblCell As Cell
    Dim lookupKey As String
    Dim cellCounter As Long
    
    ' Find the table
    Set tbl = ResolveTableByID(tableID)
    If tbl Is Nothing Then
        ExportTableContentPreview = "(Table not found)"
        Exit Function
    End If
    
    Set cellLookup = NewDictionary()
    cellCounter = 0
    For Each tblCell In tbl.Range.Cells
        lookupKey = CStr(tblCell.RowIndex) & ":" & CStr(tblCell.ColumnIndex)
        If Not cellLookup.Exists(lookupKey) Then
            cellLookup.Add lookupKey, EscapeMarkdownCellPreview(LimitPreviewText(GetCellTextFinalView(tblCell), DSM_TABLE_CELL_PREVIEW_MAX_CHARS))
        End If
        cellCounter = cellCounter + 1
        If cellCounter Mod 100 = 0 Then DoEvents
    Next tblCell

    actualRows = GetSafeTableRowCount(tbl)
    actualCols = GetSafeTableColCount(tbl)

    maxRows = actualRows
    If maxRows > DSM_TABLE_PREVIEW_MAX_ROWS Then maxRows = DSM_TABLE_PREVIEW_MAX_ROWS
    
    maxCols = actualCols
    If maxCols > DSM_TABLE_PREVIEW_MAX_COLS Then maxCols = DSM_TABLE_PREVIEW_MAX_COLS
    
    If maxRows = 0 Or maxCols = 0 Then
        ExportTableContentPreview = "(Table has no readable cells)"
        Exit Function
    End If

    output = "|"
    For c = 1 To maxCols
        lookupKey = "1:" & CStr(c)
        If cellLookup.Exists(lookupKey) Then
            cellText = CStr(cellLookup(lookupKey))
        Else
            cellText = ""
        End If
        output = output & " [" & tableID & ".H.C" & c & "] " & cellText & " |"
    Next c
    output = output & vbCrLf & "|"
    For c = 1 To maxCols
        output = output & "---|"
    Next c
    output = output & vbCrLf
    
    For r = 2 To maxRows
        output = output & "|"
        For c = 1 To maxCols
            lookupKey = CStr(r) & ":" & CStr(c)
            If cellLookup.Exists(lookupKey) Then
                cellText = CStr(cellLookup(lookupKey))
            Else
                cellText = ""
            End If
            output = output & " [" & tableID & ".R" & r & ".C" & c & "] " & cellText & " |"
        Next c
        output = output & vbCrLf
        If r Mod 10 = 0 Then DoEvents
    Next r
    
    If actualRows > maxRows Or actualCols > maxCols Then
        output = output & "(Preview truncated to " & maxRows & " rows x " & maxCols & " columns)" & vbCrLf
    End If
    
    ExportTableContentPreview = output
    Exit Function

Fail:
    ExportTableContentPreview = "(Table preview unavailable: " & Err.Description & ")"
End Function

Private Function ResolveTableByID(ByVal tableID As String) As Table
    ' Helper to find table by ID (e.g., "T2")
    On Error Resume Next
    
    Dim tblNum As Long
    tblNum = CLng(Mid$(tableID, 2))  ' Extract number from "T2"
    
    If tblNum > 0 And tblNum <= ActiveDocument.Tables.Count Then
        Set ResolveTableByID = ActiveDocument.Tables(tblNum)
    Else
        Set ResolveTableByID = Nothing
    End If
End Function

' =========================================================================================
' === TARGET RESOLUTION FUNCTIONS (V4 Architecture) ======================================
' =========================================================================================

Private Function ParseTargetReference(ByVal targetRef As String) As TargetReference
    ' Parses target reference string into structured format
    ' Formats: P5, T2, T2.R3, T2.H, T2.R3.C2, T2.H.C2
    
    On Error GoTo ParseError
    
    Dim result As TargetReference
    Dim parts() As String
    Dim firstChar As String
    
    targetRef = Trim$(UCase$(targetRef))
    
    If Len(targetRef) = 0 Then GoTo ParseError
    
    firstChar = Left$(targetRef, 1)
    
    If firstChar = "P" Then
        ' Paragraph reference: P5
        result.TargetType = "paragraph"
        result.ParagraphNum = CLng(Mid$(targetRef, 2))
        result.IsValid = True
        
    ElseIf firstChar = "T" Then
        ' Table reference: T2, T2.R3, T2.H, T2.R3.C2, T2.H.C2
        parts = Split(targetRef, ".")
        
        ' Extract table number
        result.TableNum = CLng(Mid$(parts(0), 2))
        
        If UBound(parts) = 0 Then
            ' Just T2 - entire table
            result.TargetType = "table"
            result.IsValid = True
            
        ElseIf UBound(parts) = 1 Then
            ' T2.R3 or T2.H - table row
            If Left$(parts(1), 1) = "R" Then
                result.TargetType = "table_row"
                result.RowNum = CLng(Mid$(parts(1), 2))
                result.IsHeaderRow = False
                result.IsValid = True
            ElseIf parts(1) = "H" Then
                result.TargetType = "table_row"
                result.RowNum = 1
                result.IsHeaderRow = True
                result.IsValid = True
            Else
                GoTo ParseError
            End If
            
        ElseIf UBound(parts) = 2 Then
            ' T2.R3.C2 or T2.H.C2 - table cell
            result.TargetType = "table_cell"
            
            If Left$(parts(1), 1) = "H" Then
                ' Header row
                result.IsHeaderRow = True
                result.RowNum = 1
            ElseIf Left$(parts(1), 1) = "R" Then
                result.IsHeaderRow = False
                result.RowNum = CLng(Mid$(parts(1), 2))
            Else
                GoTo ParseError
            End If
            
            If Left$(parts(2), 1) = "C" Then
                result.ColNum = CLng(Mid$(parts(2), 2))
                result.IsValid = True
            Else
                GoTo ParseError
            End If
        Else
            GoTo ParseError
        End If
    Else
        GoTo ParseError
    End If
    
    With ParseTargetReference
        .IsValid = result.IsValid
        .TargetType = result.TargetType
        .ParagraphNum = result.ParagraphNum
        .TableNum = result.TableNum
        .RowNum = result.RowNum
        .ColNum = result.ColNum
        .IsHeaderRow = result.IsHeaderRow
    End With
    Exit Function
    
ParseError:
    With ParseTargetReference
        .IsValid = False
        .TargetType = ""
        .ParagraphNum = 0
        .TableNum = 0
        .RowNum = 0
        .ColNum = 0
        .IsHeaderRow = False
    End With
    TraceLog "Failed to parse target reference: " & targetRef
End Function

Private Function ResolveTargetToRange(ByVal targetRef As String) As Range
    ' Resolves a target reference to a Word Range
    ' This is the KEY function that replaces all text-based searching
    
    On Error GoTo ErrorHandler
    
    Dim ref As TargetReference
    Dim tbl As Table
    Dim rowCount As Long
    
    ' Ensure structure map is built
    ' If Not g_DocumentMapBuilt Then
    '     BuildDocumentStructureMap ActiveDocument
    ' End If
    
    ' Parse the reference
    ref = ParseTargetReference(targetRef)
    
    If Not ref.IsValid Then
        TraceLog "Invalid target reference: " & targetRef
        Set ResolveTargetToRange = Nothing
        Exit Function
    End If
    
    ' Resolve based on type
    Select Case ref.TargetType
        Case "paragraph"
            ' Resolve paragraph against the current document to avoid stale coordinates after earlier edits.
            Set ResolveTargetToRange = ResolveBodyParagraphRange(ref.ParagraphNum)
            If Not ResolveTargetToRange Is Nothing Then
                TraceLog "Resolved " & targetRef & " to paragraph at position " & ResolveTargetToRange.Start
                Exit Function
            End If
            
        Case "table"
            ' Return entire table range
            Set tbl = ResolveTableByID("T" & ref.TableNum)
            If Not tbl Is Nothing Then
                Set ResolveTargetToRange = tbl.Range
                TraceLog "Resolved " & targetRef & " to table"
                Exit Function
            End If
            
        Case "table_row"
            ' Return row range
            Set tbl = ResolveTableByID("T" & ref.TableNum)
            If Not tbl Is Nothing Then
                Dim rowRange As Range
                rowCount = GetSafeTableRowCount(tbl)
                If ref.RowNum > 0 And ref.RowNum <= rowCount Then
                    If TryGetTableRowRange(tbl, ref.RowNum, rowRange) Then
                        Set ResolveTargetToRange = rowRange
                        TraceLog "Resolved " & targetRef & " to table row"
                        Exit Function
                    End If
                End If
            End If
            
        Case "table_cell"
            ' Return cell range
            Set tbl = ResolveTableByID("T" & ref.TableNum)
            If Not tbl Is Nothing Then
                Dim colCount As Long
                rowCount = GetSafeTableRowCount(tbl)
                colCount = GetSafeTableColCount(tbl)
                If ref.RowNum > 0 And ref.RowNum <= rowCount Then
                    If ref.ColNum > 0 And ref.ColNum <= colCount Then
                        Dim targetCell As Cell
                        If TryGetTableCell(tbl, ref.RowNum, ref.ColNum, targetCell) Then
                            Dim cellRange As Range
                            Set cellRange = GetCellContentRange(targetCell)
                            If Not cellRange Is Nothing Then
                                Set ResolveTargetToRange = cellRange
                                TraceLog "Resolved " & targetRef & " to table cell"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    ' If we get here, resolution failed
    TraceLog "Failed to resolve target: " & targetRef
    Set ResolveTargetToRange = Nothing
    Exit Function
    
ErrorHandler:
    TraceLog "Error resolving target " & targetRef & ": " & Err.Description
    Set ResolveTargetToRange = Nothing
End Function

' =========================================================================================
' === HOUSE STYLE INTEGRATION =============================================================
' =========================================================================================
' Customize this function to call your organization's style functions.
' Return True if the style was applied successfully, False to use fallback.

Private Function ApplyHouseStyle(ByVal rng As Range, ByVal styleName As String) As Boolean
    ' Applies VA Addin styles by calling the appropriate style function
    ' Returns True if style was applied successfully, False to use fallback

    On Error GoTo Fail

    Dim normalizedStyle As String
    Dim macroName As String
    normalizedStyle = LCase$(Trim$(styleName))
    macroName = ""

    ' Select the range (required by VA Addin style functions)
    rng.Select

    Select Case normalizedStyle
        ' VA Addin Report Styles
        Case "heading_l1", "report level 1"
            macroName = "RChapter"

        Case "heading_l2", "report level 2"
            macroName = "RSectionheading"

        Case "heading_l3", "report level 3"
            macroName = "RSubsection"

        Case "heading_l4", "report level 4"
            macroName = "RHeadingL4"

        Case "body_text", "text", "report text"
            macroName = "RSection"

        Case "bullet", "report bullet"
            macroName = "RBullet"

        Case "table_heading", "table heading"
            macroName = "Tableheading"

        Case "table_text", "table text"
            macroName = "Tabletext"

        Case "table_title", "report table number"
            macroName = "RTabletitle"

        Case "figure", "report figure"
            macroName = "Rfigure"
    End Select

    If Len(macroName) > 0 Then
        If CallVaStyleMacro(macroName) Then
            ApplyHouseStyle = True
            Exit Function
        End If
        Debug.Print "  -> VA style macro unavailable, falling back: " & macroName
    End If

    ' No matching macro or call failed - use fallback
    ApplyHouseStyle = False
    Exit Function

Fail:
    ' Error occurred - use fallback
    ApplyHouseStyle = False
End Function

Private Function CallVaStyleMacro(ByVal macroName As String) As Boolean
    ' Invokes a VA Addin style macro by name without compile-time dependency

    On Error GoTo Fail

    Dim candidates As Variant
    Dim candidate As Variant

    ' Try common macro name formats to handle Normal, add-in, or module-qualified macros
    candidates = Array( _
        macroName, _
        "Module03Formatting." & macroName, _
        "VA_Addin.Module03Formatting." & macroName, _
        "VAAddin.Module03Formatting." & macroName)

    For Each candidate In candidates
        If Len(candidate) = 0 Then GoTo NextCandidate

        On Error Resume Next
        Application.Run candidate
        If Err.Number = 0 Then
            On Error GoTo Fail
            CallVaStyleMacro = True
            Exit Function
        End If
        On Error GoTo Fail
NextCandidate:
    Next candidate

Fail:
    If Err.Number <> 0 Then
        Debug.Print "  -> CallVaStyleMacro failed: " & Err.Description
        Err.Clear
    End If
    CallVaStyleMacro = False
End Function

' =========================================================================================
' === V4 ACTION EXECUTORS (DSM-Based) ====================================================
' =========================================================================================

Private Sub ClearLastActionError()
    g_LastActionErrorCode = ""
    g_LastActionErrorMessage = ""
End Sub

Private Sub SetLastActionError(ByVal errorCode As String, ByVal message As String)
    g_LastActionErrorCode = errorCode
    g_LastActionErrorMessage = message
End Sub

Private Function ExecuteReplaceActionV4(ByVal targetRange As Range, ByVal findText As String, ByVal replaceText As String, Optional ByVal matchCase As Boolean = False) As Boolean
    ' V4 Replace action - find and replace within the resolved target range
    
    On Error GoTo ErrorHandler
    
    Dim actionRange As Range
    Set actionRange = FindLongString(NormalizeForDocument(findText), targetRange, matchCase)
    If actionRange Is Nothing Then
        TraceLog "  -> Replace action: text not found in target range"
        SetLastActionError "TEXT_NOT_FOUND", "Find text was not found inside the resolved target."
        ExecuteReplaceActionV4 = False
        Exit Function
    End If

    ' Skip formatting-only changes if target already matches formatting and text.
    If IsFormattingAlreadyApplied(actionRange, replaceText) Then
        TraceLog "  -> Replace action skipped (already matches): '" & findText & "'"
        ExecuteReplaceActionV4 = True
        Exit Function
    End If

    If InStr(1, replaceText, "<", vbBinaryCompare) > 0 Then
        ' Formatted replacement - parse tags and apply with granular diff
        ApplyFormattedReplacement actionRange, replaceText
    ElseIf USE_GRANULAR_DIFF And Len(actionRange.Text) <= 1000 And Len(replaceText) <= 1000 Then
        ' Plain text - use granular diff for minimal tracked changes
        Dim diffOps As Collection
        Set diffOps = ComputeDiff(actionRange.Text, replaceText)
        ApplyDiffOperations actionRange, diffOps, Nothing
    Else
        ' Fallback for very long text - wholesale replacement
        actionRange.Text = replaceText
    End If

    TraceLog "  -> Replace action executed: '" & findText & "' -> '" & replaceText & "'"
    ClearLastActionError
    ExecuteReplaceActionV4 = True
    Exit Function
    
ErrorHandler:
    TraceLog "  -> Error in ExecuteReplaceActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteReplaceActionV4 = False
End Function

Private Function ExecuteApplyStyleActionV4(ByVal targetRange As Range, ByVal styleKey As String) As Boolean
    ' V4 Apply Style action - applies VA Addin style or Word style
    
    On Error GoTo ErrorHandler
    
    ' Try VA Addin style first
    If ApplyHouseStyle(targetRange, styleKey) Then
        Debug.Print "  -> Style applied via VA Addin: " & styleKey
        ClearLastActionError
        ExecuteApplyStyleActionV4 = True
        Exit Function
    End If
    
    ' Fallback to direct Word style application
    On Error Resume Next
    targetRange.Style = styleKey
    If Err.Number = 0 Then
        Debug.Print "  -> Style applied directly: " & styleKey
        ClearLastActionError
        ExecuteApplyStyleActionV4 = True
    Else
        Debug.Print "  -> Style not found: " & styleKey
        SetLastActionError "STYLE_NOT_FOUND", "Style token not found: " & styleKey
        ExecuteApplyStyleActionV4 = False
    End If
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    Debug.Print "  -> Error in ExecuteApplyStyleActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteApplyStyleActionV4 = False
End Function

Private Function ExecuteCommentActionV4(ByVal targetRange As Range, ByVal commentText As String) As Boolean
    ' V4 Comment action - adds a comment to the target range
    
    On Error GoTo ErrorHandler
    
    ActiveDocument.Comments.Add Range:=targetRange, Text:=commentText
    Debug.Print "  -> Comment added: " & Left$(commentText, 50)
    ClearLastActionError
    ExecuteCommentActionV4 = True
    Exit Function
    
ErrorHandler:
    Debug.Print "  -> Error in ExecuteCommentActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteCommentActionV4 = False
End Function

Private Function ExecuteDeleteActionV4(ByVal targetRange As Range) As Boolean
    ' V4 Delete action - deletes the target range content
    
    On Error GoTo ErrorHandler
    
    targetRange.Delete
    Debug.Print "  -> Content deleted"
    ClearLastActionError
    ExecuteDeleteActionV4 = True
    Exit Function
    
ErrorHandler:
    Debug.Print "  -> Error in ExecuteDeleteActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteDeleteActionV4 = False
End Function

Private Function ExecuteReplaceTableActionV4(ByVal targetRange As Range, ByVal markdownTable As String) As Boolean
    ' V4 Replace Table action - replaces entire table with new markdown table
    
    On Error GoTo ErrorHandler
    
    Dim newTable As Table
    Set newTable = ConvertMarkdownToTable(targetRange, markdownTable)
    
    If Not newTable Is Nothing Then
        Debug.Print "  -> Table replaced successfully"
        ClearLastActionError
        ExecuteReplaceTableActionV4 = True
    Else
        Debug.Print "  -> Table replacement failed"
        SetLastActionError "TABLE_REPLACE_FAILED", "Markdown conversion did not return a table."
        ExecuteReplaceTableActionV4 = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "  -> Error in ExecuteReplaceTableActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteReplaceTableActionV4 = False
End Function

Private Function ExecuteInsertRowActionV4(ByVal tableRef As String, ByVal afterRow As Long, ByVal rowData As Variant) As Boolean
    ' V4 Insert Row action - inserts a new row in a table
    
    On Error GoTo ErrorHandler
    
    Dim tbl As Table
    Dim ref As TargetReference
    Dim newRow As Row
    Dim colIndex As Long
    Dim cellRange As Range
    Dim rowCount As Long
    
    ' Parse table reference
    ref = ParseTargetReference(tableRef)
    If Not ref.IsValid Or ref.TargetType <> "table" Then
        Debug.Print "  -> Invalid table reference: " & tableRef
        SetLastActionError "TARGET_NOT_FOUND", "Invalid table reference: " & tableRef
        ExecuteInsertRowActionV4 = False
        Exit Function
    End If
    
    ' Get table
    Set tbl = ResolveTableByID("T" & ref.TableNum)
    If tbl Is Nothing Then
        Debug.Print "  -> Table not found: " & tableRef
        SetLastActionError "TARGET_NOT_FOUND", "Table not found: " & tableRef
        ExecuteInsertRowActionV4 = False
        Exit Function
    End If
    
    ' Insert row (fallback to append when merged structure blocks indexed insertion).
    rowCount = GetSafeTableRowCount(tbl)
    If rowCount = 0 Then
        SetLastActionError "TABLE_STRUCTURE_UNSUPPORTED", "Unable to determine row count for this table (possibly complex merges)."
        ExecuteInsertRowActionV4 = False
        Exit Function
    End If

    On Error Resume Next
    If afterRow >= rowCount Then
        Set newRow = tbl.Rows.Add  ' Add at end
    Else
        Set newRow = tbl.Rows.Add(tbl.Rows(afterRow + 1))  ' Add after specified row
        If Err.Number <> 0 Then
            Err.Clear
            Set newRow = tbl.Rows.Add
        End If
    End If
    On Error GoTo ErrorHandler

    If newRow Is Nothing Then
        SetLastActionError "TABLE_STRUCTURE_UNSUPPORTED", "Unable to insert row in this merged table structure."
        ExecuteInsertRowActionV4 = False
        Exit Function
    End If
    
    ' Populate cells if data provided
    If IsArray(rowData) Then
        For colIndex = LBound(rowData) To UBound(rowData)
            If colIndex + 1 <= newRow.Cells.Count Then
                Set cellRange = newRow.Cells(colIndex + 1).Range
                If cellRange.End > cellRange.Start Then cellRange.End = cellRange.End - 1
                cellRange.Text = CStr(rowData(colIndex))
            End If
        Next colIndex
    End If
    
    Debug.Print "  -> Row inserted after row " & afterRow
    ClearLastActionError
    ExecuteInsertRowActionV4 = True
    Exit Function
    
ErrorHandler:
    Debug.Print "  -> Error in ExecuteInsertRowActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteInsertRowActionV4 = False
End Function

Private Function ExecuteDeleteRowActionV4(ByVal targetRange As Range) As Boolean
    ' V4 Delete Row action - deletes a table row
    
    On Error GoTo ErrorHandler
    
    If Not targetRange.Information(wdWithinTable) Then
        Debug.Print "  -> Target is not within a table"
        SetLastActionError "TARGET_NOT_FOUND", "Target does not resolve to a table row."
        ExecuteDeleteRowActionV4 = False
        Exit Function
    End If
    
    Dim tbl As Table
    Dim rowNum As Long
    
    Set tbl = targetRange.Tables(1)
    rowNum = targetRange.Cells(1).RowIndex
    
    tbl.Rows(rowNum).Delete
    Debug.Print "  -> Row " & rowNum & " deleted"
    ClearLastActionError
    ExecuteDeleteRowActionV4 = True
    Exit Function
    
ErrorHandler:
    Debug.Print "  -> Error in ExecuteDeleteRowActionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ExecuteDeleteRowActionV4 = False
End Function

' =========================================================================================
' === V4 MAIN PROCESSING FUNCTIONS =======================================================
' =========================================================================================

Private Function ProcessSuggestionV4(ByVal suggestion As Object) As Boolean
    ' V4 suggestion processor - uses target references instead of text searching
    
    On Error GoTo ErrorHandler
    
    Dim targetRef As String
    Dim action As String
    Dim targetRange As Range
    Dim findText As String
    Dim replaceText As String
    Dim styleKey As String
    Dim explanation As String
    Dim matchCase As Boolean
    Dim success As Boolean

    ClearLastActionError
    
    ' Extract required fields
    If Not HasDictionaryKey(suggestion, "target") Then
        TraceLog "  -> Missing 'target' field"
        SetLastActionError "INVALID_TOOL_CALL", "Missing 'target' field."
        ProcessSuggestionV4 = False
        Exit Function
    End If
    
    If Not HasDictionaryKey(suggestion, "action") Then
        TraceLog "  -> Missing 'action' field"
        SetLastActionError "INVALID_TOOL_CALL", "Missing 'action' field."
        ProcessSuggestionV4 = False
        Exit Function
    End If
    
    targetRef = GetSuggestionText(suggestion, "target", "")
    action = LCase$(Trim$(GetSuggestionText(suggestion, "action", "")))
    explanation = GetSuggestionText(suggestion, "explanation", "")
    
    TraceLog "  Processing: " & targetRef & " | Action: " & action
    
    ' NEW: Use the Pre-Resolved Range if available, otherwise fallback to lookup
    If HasDictionaryKey(suggestion, "pre_resolved_range") Then
        Set targetRange = suggestion("pre_resolved_range")
        TraceLog "  -> Using pre-resolved Range for: " & targetRef
    Else
        Set targetRange = ResolveTargetToRange(targetRef)
    End If
    
    If targetRange Is Nothing Then
        TraceLog "  -> Failed to resolve target: " & targetRef
        SetLastActionError "TARGET_NOT_FOUND", "Target could not be resolved: " & targetRef
        ProcessSuggestionV4 = False
        Exit Function
    End If
    
    ' Execute action based on type
    Select Case action
        Case "replace"
            findText = GetSuggestionText(suggestion, "find", "")
            replaceText = GetSuggestionText(suggestion, "replace", "")
            matchCase = False
            If HasDictionaryKey(suggestion, "match_case") Then
                On Error Resume Next
                matchCase = CBool(suggestion("match_case"))
                On Error GoTo ErrorHandler
            End If
            success = ExecuteReplaceActionV4(targetRange, findText, replaceText, matchCase)
            
        Case "apply_style"
            styleKey = GetSuggestionText(suggestion, "style", "")
            success = ExecuteApplyStyleActionV4(targetRange, styleKey)
            
        Case "comment"
            success = ExecuteCommentActionV4(targetRange, explanation)
            
        Case "delete"
            success = ExecuteDeleteActionV4(targetRange)
            
        Case "replace_table"
            replaceText = GetSuggestionText(suggestion, "replace", "")
            success = ExecuteReplaceTableActionV4(targetRange, replaceText)
            
        Case "insert_row"
            Dim afterRow As Long
            Dim rowData As Variant
            afterRow = 0
            If HasDictionaryKey(suggestion, "after_row") Then
                afterRow = CLng(Val(CStr(suggestion("after_row"))))
            End If
            If HasDictionaryKey(suggestion, "data") Then
                rowData = suggestion("data")
            End If
            success = ExecuteInsertRowActionV4(targetRef, afterRow, rowData)
            
        Case "delete_row"
            success = ExecuteDeleteRowActionV4(targetRange)
            
        Case Else
            TraceLog "  -> Unknown action type: " & action
            SetLastActionError "INVALID_TOOL_CALL", "Unknown action type: " & action
            success = False
    End Select
    
    ProcessSuggestionV4 = success
    Exit Function
    
ErrorHandler:
    TraceLog "  -> Error in ProcessSuggestionV4: " & Err.Description
    SetLastActionError "EXECUTION_ERROR", Err.Description
    ProcessSuggestionV4 = False
End Function

Private Sub ApplyLlmReview_V4()
    ' V4 Main entry point - uses Document Structure Map for deterministic targeting
    
    On Error GoTo ErrorHandler
    
    Dim inputForm As New frmJsonInput
    Dim jsonString As String
    Dim suggestions As Object
    Dim suggestion As Object
    Dim successCount As Long
    Dim failCount As Long
    Dim totalCount As Long
    Dim startTime As Single
    
    Debug.Print vbCrLf & "================ APPLY LLM REVIEW V4 (DSM) ================"
    startTime = Timer
    
    ' Show input form
    inputForm.Show vbModal
    
    ' Get JSON from form
    jsonString = PreProcessJson(inputForm.txtJson.value)
    
    ' Parse JSON
    Set suggestions = LLM_ParseJson(jsonString)
    If suggestions Is Nothing Or TypeName(suggestions) <> "Collection" Then
        MsgBox "Failed to parse JSON. Please check the format.", vbCritical, "JSON Error"
        Exit Sub
    End If
    
    totalCount = suggestions.Count
    Debug.Print "Found " & totalCount & " suggestions"
    
    ' Build/refresh structure map
    ClearDocumentStructureMap
    BuildDocumentStructureMap ActiveDocument
    
    If g_DocumentMapCount = 0 Then
        MsgBox "Document structure map could not be built.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Debug.Print "Structure map ready: " & g_DocumentMapCount & " elements"
    
    ' Process each suggestion
    successCount = 0
    failCount = 0
    
    For Each suggestion In suggestions
        If ProcessSuggestionV4(suggestion) Then
            successCount = successCount + 1
        Else
            failCount = failCount + 1
        End If
    Next suggestion
    
    ' Report results
    Dim duration As Single
    duration = Timer - startTime
    
    Dim report As String
    report = "LLM Review V4 Complete!" & vbCrLf & vbCrLf
    report = report & "Total Suggestions: " & totalCount & vbCrLf
    report = report & "Successfully Applied: " & successCount & vbCrLf
    report = report & "Failed: " & failCount & vbCrLf
    report = report & "Duration: " & Format(duration, "0.0") & " seconds"
    
    Debug.Print report
    MsgBox report, vbInformation, "Review Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ApplyLlmReview_V4: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub RunInteractiveReview_V4()
    ' V4 Interactive review - shows preview for each suggestion with accept/reject
    
    On Error GoTo ErrorHandler
    
    Dim inputForm As New frmJsonInput
    Dim jsonString As String
    Dim suggestions As Object
    Dim suggestion As Object
    Dim index As Long
    Dim totalCount As Long
    Dim acceptedCount As Long
    Dim rejectedCount As Long
    Dim skippedCount As Long
    Dim userAction As String
    Dim startTime As Single
    
    Debug.Print vbCrLf & "================ INTERACTIVE REVIEW V4 (DSM) ================"
    startTime = Timer
    
    ' Show input form
    inputForm.Show vbModal
    
    ' Get JSON from form
    jsonString = PreProcessJson(inputForm.txtJson.value)
    
    ' Parse JSON
    Set suggestions = LLM_ParseJson(jsonString)
    If suggestions Is Nothing Or TypeName(suggestions) <> "Collection" Then
        MsgBox "Failed to parse JSON. Please check the format.", vbCritical, "JSON Error"
        Exit Sub
    End If
    
    totalCount = suggestions.Count
    Debug.Print "Found " & totalCount & " suggestions for interactive review"
    
    ' Build/refresh structure map
    ClearDocumentStructureMap
    BuildDocumentStructureMap ActiveDocument
    
    If g_DocumentMapCount = 0 Then
        MsgBox "Document structure map could not be built.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Debug.Print "Structure map ready: " & g_DocumentMapCount & " elements"
    
    ' Process each suggestion interactively
    acceptedCount = 0
    rejectedCount = 0
    skippedCount = 0
    index = 0
    
    For Each suggestion In suggestions
        index = index + 1
        
        ' Show preview and get user action
        userAction = ShowSuggestionPreviewV4(suggestion, index, totalCount)
        
        Select Case UCase$(userAction)
            Case "ACCEPT"
                ' Apply the suggestion
                If ProcessSuggestionV4(suggestion) Then
                    acceptedCount = acceptedCount + 1
                    Debug.Print "  [" & index & "/" & totalCount & "] ACCEPTED"
                Else
                    Debug.Print "  [" & index & "/" & totalCount & "] ACCEPT FAILED"
                    rejectedCount = rejectedCount + 1
                End If
                
            Case "REJECT"
                rejectedCount = rejectedCount + 1
                Debug.Print "  [" & index & "/" & totalCount & "] REJECTED"
                
            Case "SKIP"
                skippedCount = skippedCount + 1
                Debug.Print "  [" & index & "/" & totalCount & "] SKIPPED"
                
            Case "ACCEPT_ALL"
                ' Accept this one and all remaining
                If ProcessSuggestionV4(suggestion) Then
                    acceptedCount = acceptedCount + 1
                End If
                
                ' Process remaining suggestions without preview
                Dim remainingSuggestion As Object
                Dim remainingIndex As Long
                remainingIndex = index
                
                For Each remainingSuggestion In suggestions
                    remainingIndex = remainingIndex + 1
                    If remainingIndex > index Then
                        If ProcessSuggestionV4(remainingSuggestion) Then
                            acceptedCount = acceptedCount + 1
                        Else
                            rejectedCount = rejectedCount + 1
                        End If
                    End If
                Next remainingSuggestion
                
                Exit For
                
            Case "STOP"
                skippedCount = skippedCount + (totalCount - index + 1)
                Debug.Print "  [" & index & "/" & totalCount & "] STOPPED by user"
                Exit For
                
            Case Else
                ' Treat unknown as skip
                skippedCount = skippedCount + 1
        End Select
    Next suggestion
    
    ' Report results
    Dim duration As Single
    duration = Timer - startTime
    
    Dim report As String
    report = "Interactive Review V4 Complete!" & vbCrLf & vbCrLf
    report = report & "Total Suggestions: " & totalCount & vbCrLf
    report = report & "Accepted: " & acceptedCount & vbCrLf
    report = report & "Rejected: " & rejectedCount & vbCrLf
    report = report & "Skipped: " & skippedCount & vbCrLf
    report = report & "Duration: " & Format(duration, "0.0") & " seconds"
    
    Debug.Print report
    MsgBox report, vbInformation, "Review Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in RunInteractiveReview_V4: " & Err.Description, vbCritical, "Error"
End Sub

Private Function ShowSuggestionPreviewV4(ByVal suggestion As Object, ByVal index As Long, ByVal total As Long) As String
    ' V4 Preview function - shows target reference and resolved range
    
    On Error GoTo ErrorHandler
    
    Dim targetRef As String
    Dim action As String
    Dim explanation As String
    Dim targetRange As Range
    Dim previewForm As frmSuggestionPreview
    Dim findText As String
    Dim replaceText As String
    Dim styleKey As String
    
    ' Extract fields
    targetRef = GetSuggestionText(suggestion, "target", "")
    action = GetSuggestionText(suggestion, "action", "")
    explanation = GetSuggestionText(suggestion, "explanation", "")
    
    ' Resolve target
    Set targetRange = ResolveTargetToRange(targetRef)
    
    ' Create preview form
    Set previewForm = New frmSuggestionPreview
    
    If targetRange Is Nothing Then
        ' Target not found - show error in preview
        previewForm.LoadSuggestionV4 suggestion, index, total, Nothing, Nothing, "Target not found: " & targetRef
    Else
        ' Build preview text based on action type
        Select Case LCase$(action)
            Case "replace"
                findText = GetSuggestionText(suggestion, "find", "")
                replaceText = GetSuggestionText(suggestion, "replace", "")
                previewForm.LoadSuggestionV4 suggestion, index, total, targetRange, targetRange, ""
                
            Case "apply_style"
                styleKey = GetSuggestionText(suggestion, "style", "")
                previewForm.LoadSuggestionV4 suggestion, index, total, targetRange, targetRange, ""
                
            Case Else
                previewForm.LoadSuggestionV4 suggestion, index, total, targetRange, targetRange, ""
        End Select
        
        ' Highlight the target range
        targetRange.HighlightColorIndex = wdYellow
    End If
    
    ' Show form and wait for user action
    previewForm.Show vbModal
    
    ' Get user action
    ShowSuggestionPreviewV4 = previewForm.UserAction
    
    ' Clear highlighting
    If Not targetRange Is Nothing Then
        targetRange.HighlightColorIndex = wdNoHighlight
    End If
    
    ' Cleanup
    Unload previewForm
    Set previewForm = Nothing
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ShowSuggestionPreviewV4: " & Err.Description
    ShowSuggestionPreviewV4 = "SKIP"
End Function

Private Function PromptUserToSelectTable(ByVal searchTitle As String, ByVal rowHeader As String, ByVal columnHeader As String) As Table
    ' Prompts the user to select from multiple matching tables
    ' Returns the selected table or Nothing if cancelled

    On Error GoTo Fail

    Dim candidates As Collection
    Set candidates = GetMatchingTableCandidates(searchTitle)

    If candidates.Count = 0 Then
        Set PromptUserToSelectTable = Nothing
        Exit Function
    End If

    If candidates.Count = 1 Then
        Set PromptUserToSelectTable = g_TableIndex(candidates(1)).TableRef
        Exit Function
    End If

    ' Build selection dialog message
    Dim msg As String
    Dim i As Long
    Dim idx As Long
    Dim previewText As String

    msg = "Multiple tables match: """ & Left$(searchTitle, 50) & """" & vbCrLf & vbCrLf
    msg = msg & "Looking for cell with:" & vbCrLf
    If Len(rowHeader) > 0 Then msg = msg & "  Row: " & rowHeader & vbCrLf
    If Len(columnHeader) > 0 Then msg = msg & "  Column: " & columnHeader & vbCrLf
    msg = msg & vbCrLf & "Please select the correct table:" & vbCrLf & vbCrLf

    For i = 1 To candidates.Count
        idx = candidates(i)
        msg = msg & i & ". Table at position " & g_TableIndex(idx).TableIndex & vbCrLf

        ' Show title/caption info
        If Len(g_TableIndex(idx).TitleBelow) > 0 Then
            msg = msg & "   Title: " & Left$(g_TableIndex(idx).TitleBelow, 60)
            If Len(g_TableIndex(idx).TitleBelow) > 60 Then msg = msg & "..."
            msg = msg & vbCrLf
        End If

        ' Show first row preview
        If Len(g_TableIndex(idx).HeaderRowText) > 0 Then
            previewText = Left$(g_TableIndex(idx).HeaderRowText, 50)
            If Len(g_TableIndex(idx).HeaderRowText) > 50 Then previewText = previewText & "..."
            msg = msg & "   First row: " & previewText & vbCrLf
        End If
        msg = msg & vbCrLf
    Next i

    msg = msg & "Enter number (1-" & candidates.Count & "), or 0 to skip:"

    ' Get user input
    Dim userInput As String
    userInput = InputBox(msg, "Select Table", "1")

    If Len(userInput) = 0 Then
        ' User cancelled
        Set PromptUserToSelectTable = Nothing
        Exit Function
    End If

    Dim selection As Long
    On Error Resume Next
    selection = CLng(userInput)
    On Error GoTo Fail

    If selection < 1 Or selection > candidates.Count Then
        ' Invalid selection or 0 to skip
        Set PromptUserToSelectTable = Nothing
        Exit Function
    End If

    ' Return the selected table
    idx = candidates(selection)
    Set PromptUserToSelectTable = g_TableIndex(idx).TableRef

    ' Highlight the selected table briefly for visual confirmation
    On Error Resume Next
    g_TableIndex(idx).TableRef.Range.Select
    Application.ScreenRefresh
    On Error GoTo Fail

    Debug.Print "    -> User selected table " & selection & " (index " & idx & ")"
    Exit Function

Fail:
    Set PromptUserToSelectTable = Nothing
End Function

Private Function IsRangeInTOC(ByVal testRange As Range) As Boolean
    ' Checks if a given range is within a Table of Contents field
    ' Returns True if the range is inside a TOC, False otherwise

    On Error GoTo Fail

    If testRange Is Nothing Then
        IsRangeInTOC = False
        Exit Function
    End If

    EnsureTocFieldRangeCache testRange.Document
    If g_TocFieldRanges Is Nothing Then
        IsRangeInTOC = False
        Exit Function
    End If

    Dim fieldRange As Range
    For Each fieldRange In g_TocFieldRanges
        If testRange.Start >= fieldRange.Start And testRange.Start < fieldRange.End Then
            IsRangeInTOC = True
            TraceLog "    -> Skipping match: Range is within TOC at position " & fieldRange.Start
            Exit Function
        End If
    Next fieldRange

    IsRangeInTOC = False
    Exit Function

Fail:
    IsRangeInTOC = False
End Function

Private Function IsSkippablePipeSegment(ByVal segment As String) As Boolean
    Dim t As String
    t = Trim$(segment)
    If Len(t) = 0 Then
        IsSkippablePipeSegment = True
        Exit Function
    End If
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch Like "[A-Za-z]") Then
            IsSkippablePipeSegment = False
            Exit Function
        End If
    Next i
    IsSkippablePipeSegment = True
End Function

Private Function FindSegmentWithFallback(ByVal segment As String, ByVal searchRange As Range, ByVal matchCase As Boolean) As Range
    Dim r As Range
    Set r = FindLongString(segment, searchRange, matchCase)
    If Not r Is Nothing Then
        Set FindSegmentWithFallback = r
        Exit Function
    End If

    Set r = FindSegmentByStableTokens(segment, searchRange, matchCase)
    Set FindSegmentWithFallback = r
End Function

Private Function FindSegmentByStableTokens(ByVal segment As String, ByVal searchRange As Range, ByVal matchCase As Boolean) As Range
    Dim tokens As Collection
    Set tokens = ExtractStableTokens(segment)
    If tokens Is Nothing Then
        Set FindSegmentByStableTokens = Nothing
        Exit Function
    End If
    If tokens.Count = 0 Then
        Set FindSegmentByStableTokens = Nothing
        Exit Function
    End If

    Dim cursor As Range
    Set cursor = searchRange.Duplicate

    Dim firstFound As Range
    Dim lastFound As Range
    Dim i As Long
    For i = 1 To tokens.Count
        Dim look As Range
        Set look = cursor.Duplicate
        If look.End - look.Start > 200 Then look.End = look.Start + 200

        Dim found As Range
        Set found = FindLongString(CStr(tokens(i)), look, matchCase)
        If found Is Nothing Then
            Set FindSegmentByStableTokens = Nothing
            Exit Function
        End If

        If firstFound Is Nothing Then Set firstFound = found
        Set lastFound = found

        cursor.Start = found.End
        If cursor.Start >= cursor.End Then Exit For
    Next i

    If firstFound Is Nothing Or lastFound Is Nothing Then
        Set FindSegmentByStableTokens = Nothing
        Exit Function
    End If

    Dim resultRange As Range
    Set resultRange = ActiveDocument.Range(firstFound.Start, lastFound.End)
    Set FindSegmentByStableTokens = resultRange
End Function

Private Function ExtractStableTokens(ByVal segment As String) As Collection
    Dim parts() As String
    Dim tokens As New Collection
    Dim p As Variant
    Dim t As String

    parts = Split(Trim$(segment), " ")
    For Each p In parts
        t = Trim$(CStr(p))
        If Len(t) > 0 Then
            Dim hasLetter As Boolean
            Dim hasDigit As Boolean
            Dim i As Long
            Dim ch As String
            For i = 1 To Len(t)
                ch = Mid$(t, i, 1)
                If ch Like "[A-Za-z]" Then hasLetter = True
                If ch Like "[0-9]" Then hasDigit = True
            Next i

            If hasLetter Then
                If hasDigit Then
                    If Left$(t, 1) Like "[A-Za-z]" Then
                        tokens.Add t
                    Else
                        Dim lettersOnly As String
                        lettersOnly = ""
                        For i = 1 To Len(t)
                            ch = Mid$(t, i, 1)
                            If ch Like "[A-Za-z]" Then lettersOnly = lettersOnly & ch
                        Next i
                        If Len(lettersOnly) >= 2 Then tokens.Add lettersOnly
                    End If
                Else
                    tokens.Add t
                End If
            End If
        End If
    Next p

    Set ExtractStableTokens = tokens
End Function

Private Function TextMatchesHeuristic(ByVal expected As String, ByVal actual As String) As Boolean
    Dim e As String
    Dim a As String
    e = NormalizeForDocument(Trim$(expected))
    a = NormalizeForDocument(Trim$(actual))

    If Len(e) = 0 Then
        TextMatchesHeuristic = True
        Exit Function
    End If

    If InStr(1, a, e, vbTextCompare) > 0 Then
        TextMatchesHeuristic = True
        Exit Function
    End If

    Dim tokens As Collection
    Set tokens = ExtractStableTokens(e)
    If tokens Is Nothing Then
        TextMatchesHeuristic = False
        Exit Function
    End If
    If tokens.Count = 0 Then
        TextMatchesHeuristic = False
        Exit Function
    End If

    Dim i As Long
    For i = 1 To tokens.Count
        If InStr(1, a, CStr(tokens(i)), vbTextCompare) = 0 Then
            TextMatchesHeuristic = False
            Exit Function
        End If
    Next i

    TextMatchesHeuristic = True
End Function

' --- Preflight Analyzer (non-mutating) ---
Private Function PreflightAnalyze(ByVal suggestions As Object, ByVal docRange As Range, ByVal baseMatchCase As Boolean) As Object
    On Error GoTo ErrorHandler
    Dim result As Object
    Set result = NewDictionary()
    Dim actionable As New Collection
    Dim notFound As New Collection
    Dim found As New Collection
    Dim actionableCount As Long, ambiguousCount As Long, notFoundCount As Long, noopCount As Long, total As Long
    Dim tableAnchorCache As Object
    Set tableAnchorCache = NewDictionary()
    Dim usedCaptionKeys As Object
    Set usedCaptionKeys = NewDictionary()

    Dim suggestion As Object
    For Each suggestion In suggestions
        total = total + 1

        Dim hasAmbiguousTableTitle As Boolean
        hasAmbiguousTableTitle = False
        Dim usedTableAnchor As String
        usedTableAnchor = ""
        If HasDictionaryKey(suggestion, "actions") Then
            Dim ao As Object
            For Each ao In suggestion("actions")
                If HasDictionaryKey(ao, "tableCell") Then
                    usedTableAnchor = ""
                    If HasDictionaryKey(ao, "context") Then
                        usedTableAnchor = GetSuggestionText(ao, "context", "")
                    ElseIf HasDictionaryKey(suggestion, "context") Then
                        usedTableAnchor = GetSuggestionText(suggestion, "context", "")
                    End If
                    If Len(Trim$(usedTableAnchor)) > 0 Then
                        Dim cacheKey As String
                        cacheKey = NormalizeForDocument(usedTableAnchor)
                        Dim captionKey As String
                        captionKey = NormalizeForDocument(StripTableNumberPrefix(usedTableAnchor))
                        If Len(Trim$(captionKey)) > 0 Then
                            If Not usedCaptionKeys.Exists(captionKey) Then usedCaptionKeys.Add captionKey, True
                        End If
                        If Not tableAnchorCache.Exists(cacheKey) Then
                            tableAnchorCache.Add cacheKey, CLng(CountTablesByCaptionAnchor(usedTableAnchor, docRange))
                        End If
                        If CLng(tableAnchorCache(cacheKey)) > 1 Then
                            Debug.Print "  - WARNING: Ambiguous table caption/title used by tableCell action: '" & usedTableAnchor & "' (" & CLng(tableAnchorCache(cacheKey)) & " tables match)"
                            hasAmbiguousTableTitle = True
                            Exit For
                        End If
                    End If
                End If
            Next ao
        ElseIf HasDictionaryKey(suggestion, "tableCell") Then
            If HasDictionaryKey(suggestion, "context") Then
                usedTableAnchor = GetSuggestionText(suggestion, "context", "")
            End If
            If Len(Trim$(usedTableAnchor)) > 0 Then
                cacheKey = NormalizeForDocument(usedTableAnchor)
                captionKey = NormalizeForDocument(StripTableNumberPrefix(usedTableAnchor))
                If Len(Trim$(captionKey)) > 0 Then
                    If Not usedCaptionKeys.Exists(captionKey) Then usedCaptionKeys.Add captionKey, True
                End If
                If Not tableAnchorCache.Exists(cacheKey) Then
                    tableAnchorCache.Add cacheKey, CLng(CountTablesByCaptionAnchor(usedTableAnchor, docRange))
                End If
                If CLng(tableAnchorCache(cacheKey)) > 1 Then
                    Debug.Print "  - WARNING: Ambiguous table caption/title used by tableCell action: '" & usedTableAnchor & "' (" & CLng(tableAnchorCache(cacheKey)) & " tables match)"
                    hasAmbiguousTableTitle = True
                End If
            End If
        End If

        If hasAmbiguousTableTitle Then
            ambiguousCount = ambiguousCount + 1
            GoTo ContinueNext
        End If

        Dim context As String
        context = GetSuggestionContextText(suggestion)
        Dim effMatch As Boolean
        effMatch = baseMatchCase
        If HasDictionaryKey(suggestion, "matchCase") Then
            On Error Resume Next
            effMatch = CBool(suggestion("matchCase"))
            On Error GoTo ErrorHandler
        End If
        Dim ctxNorm As String
        ctxNorm = NormalizeForDocument(context)
        Dim ctxRange As Range
        Set ctxRange = FindWithProgressiveFallback(ctxNorm, docRange, effMatch, suggestion)
        If ctxRange Is Nothing Then
            notFoundCount = notFoundCount + 1
            notFound.Add suggestion
            GoTo ContinueNext
        End If

        found.Add suggestion

        Dim isAmbiguous As Boolean
        isAmbiguous = False
        If HasDictionaryKey(suggestion, "target") And Not HasDictionaryKey(suggestion, "occurrenceIndex") Then
            Dim tgt As String
            tgt = GetSuggestionText(suggestion, "target", "")
            If Len(tgt) > 0 Then
                Dim occ As Long
                occ = CountOccurrencesInRange(NormalizeForDocument(tgt), ctxRange, effMatch)
                If occ > 1 Then isAmbiguous = True
            End If
        End If
        If isAmbiguous Then
            ambiguousCount = ambiguousCount + 1
            GoTo ContinueNext
        End If

        If IsSuggestionNoOp(suggestion, ctxRange, effMatch) Then
            noopCount = noopCount + 1
            GoTo ContinueNext
        End If

        actionable.Add suggestion
        actionableCount = actionableCount + 1

ContinueNext:
    Next suggestion

    Call WarnUnusedAmbiguousTableTitles(docRange, usedCaptionKeys)

    result.Add "actionable", actionable
    result.Add "found", found
    result.Add "notFound", notFound
    result.Add "actionableCount", actionableCount
    result.Add "ambiguousCount", ambiguousCount
    result.Add "notFoundCount", notFoundCount
    result.Add "noopCount", noopCount
    result.Add "total", total
    Set PreflightAnalyze = result
    Exit Function

ErrorHandler:
    Set PreflightAnalyze = NewDictionary()
End Function

Private Function CountOccurrencesInRange(ByVal normalizedTarget As String, ByVal ctxRange As Range, ByVal matchCase As Boolean) As Long
    On Error Resume Next
    Dim count As Long
    Dim safetyCounter As Long
    Dim lastPos As Long
    count = 0
    safetyCounter = 0
    lastPos = -1
    Dim work As Range
    Set work = ctxRange.Duplicate
    Do
        safetyCounter = safetyCounter + 1
        If safetyCounter > LOOP_SAFETY_LIMIT Then
            Debug.Print "    -> SAFETY EXIT: CountOccurrencesInRange exceeded " & LOOP_SAFETY_LIMIT & " iterations"
            Exit Do
        End If
        Dim found As Range
        Set found = FindLongString(normalizedTarget, work, matchCase)
        If found Is Nothing Then Exit Do
        ' Prevent infinite loop if Find returns same position
        If found.Start = lastPos Then
            Debug.Print "    -> SAFETY EXIT: CountOccurrencesInRange stuck at position " & lastPos
            Exit Do
        End If
        lastPos = found.Start
        count = count + 1
        work.Start = found.End
        If work.Start >= work.End Then Exit Do
    Loop
    CountOccurrencesInRange = count
End Function

Private Function IsSuggestionNoOp(ByVal suggestion As Object, ByVal ctxRange As Range, ByVal matchCase As Boolean) As Boolean
    On Error GoTo CleanFail
    IsSuggestionNoOp = False
    If suggestion.Exists("actions") Then
        Dim subs As Object, ao As Object
        Set subs = suggestion("actions")
        For Each ao In subs
            If Not IsActionNoOp(ao, suggestion, ctxRange, matchCase) Then Exit Function
        Next ao
        IsSuggestionNoOp = True
    Else
        IsSuggestionNoOp = IsActionNoOp(suggestion, suggestion, ctxRange, matchCase)
    End If
    Exit Function
CleanFail:
    IsSuggestionNoOp = False
End Function

Private Function IsActionNoOp(ByVal actionObject As Object, ByVal topSuggestion As Object, ByVal ctxRange As Range, ByVal matchCase As Boolean) As Boolean
    On Error GoTo FailFalse
    Dim aName As String
    aName = LCase$(Trim$(GetSuggestionText(actionObject, "action", "")))
    Dim target As String
    target = GetSuggestionText(actionObject, "target", "")
    Dim repl As String
    repl = GetSuggestionText(actionObject, "replace", "")

    Dim actionRange As Range
    If Len(target) > 0 Then
        Dim tNorm As String
        tNorm = NormalizeForDocument(target)
        Dim occ As Long
        occ = 1
        If HasDictionaryKey(actionObject, "occurrenceIndex") Then occ = CLng(Val(CStr(actionObject("occurrenceIndex"))))
        Dim i As Long
        Dim searchSub As Range
        Set searchSub = ctxRange.Duplicate
        For i = 1 To occ
            Set actionRange = FindWithProgressiveFallback(tNorm, searchSub, matchCase, actionObject)
            If actionRange Is Nothing Then GoTo FailFalse
            If i < occ Then searchSub.Start = actionRange.End
        Next i
    Else
        Set actionRange = ctxRange.Duplicate
    End If

    Select Case aName
        Case "apply_heading_style"
            If Len(Trim$(repl)) = 0 Then GoTo FailFalse
            On Error Resume Next
            IsActionNoOp = (actionRange.Style = Trim$(repl))
            On Error GoTo 0
            Exit Function
        Case "change", "replace"
            IsActionNoOp = IsFormattingAlreadyApplied(actionRange, repl)
            Exit Function
        Case "comment", "replace_with_table"
            IsActionNoOp = False
            Exit Function
        Case Else
            IsActionNoOp = False
            Exit Function
    End Select

FailFalse:
    IsActionNoOp = False
End Function

' =========================================================================================
' === MAIN SUBROUTINE TO RUN ==============================================================
' =========================================================================================
Private Sub ApplyLlmReview_V3()
    ' This sub now only serves to launch the UserForm.
    ' The form's code-behind now controls the workflow.
    Dim inputForm As New frmJsonInput
    Debug.Print vbCrLf & "================ RUN START: " & Now() & " ================"
    inputForm.Show vbModal
    Unload inputForm
End Sub

' =========================================================================================

' === CORE WORKFLOW (Called by the Form) ==================================================

' =========================================================================================

Private Sub RunReviewProcess(ByVal TheForm As frmJsonInput)
    ' *** WORKFLOW SELECTOR ***
    ' Set to True for INTERACTIVE mode (new), False for TRACKED CHANGES mode (old)
    ' Disable (False) pros: Applies all suggestions in one pass with summary/fallback comments.
    ' Disable (False) cons: Less granular operator control than per-suggestion interactive approval.
    Const USE_INTERACTIVE_MODE As Boolean = True

    Dim jsonString As String
    Dim suggestions As Object
    Dim startTime As Single
    Debug.Print "Starting RunReviewProcess"

    startTime = Timer

    ' --- 1. Get and Pre-Process JSON from the form ---
    Debug.Print "Preprocessing JSON..."
    jsonString = PreProcessJson(TheForm.txtJson.value)

    ' --- 2. Parse the JSON ---
    Debug.Print "JSON head: " & Left$(jsonString, 200)
    Debug.Print "Parsing JSON now..."
    On Error Resume Next ' temporarily catch errors from the parser call
    Set suggestions = LLM_ParseJson(jsonString)
    If Err.Number <> 0 Then
        Debug.Print "Parser call failed. Err " & Err.Number & ": " & Err.Description
        MsgBox "Failed to parse JSON. Error: " & Err.Description, vbCritical, "JSON Parse Error"
        Exit Sub
    End If

    On Error GoTo 0
    If suggestions Is Nothing Or Not TypeName(suggestions) = "Collection" Then
        Debug.Print "INFO: Failed to parse JSON. Please validate the format first."
        MsgBox "Failed to parse JSON. Please validate the format first.", vbCritical, "JSON Error"
        Exit Sub
    End If

    Dim totalCount As Long
    totalCount = suggestions.Count
    Debug.Print "Found " & totalCount & " suggestions to process."

    ' --- 2b. Build table index for fast table lookup ---
    ClearTableIndex
    BuildTableIndex ActiveDocument

    ' --- 3. Route to appropriate workflow ---
    If USE_INTERACTIVE_MODE Then
        ' NEW: Interactive preview mode
        TheForm.Hide
        Dim baseMatchCase As Boolean
        baseMatchCase = CBool(TheForm.chkCaseSensitive.value)
        Dim analyzed As Object
        Set analyzed = PreflightAnalyze(suggestions, ActiveDocument.Content, baseMatchCase)
        Dim actionableCount As Long, ambiguousCount As Long, notFoundCount As Long, noopCount As Long, totalAnalyzed As Long
        actionableCount = analyzed("actionableCount")
        ambiguousCount = analyzed("ambiguousCount")
        notFoundCount = analyzed("notFoundCount")
        noopCount = analyzed("noopCount")
        totalAnalyzed = analyzed("total")
        Dim msg As String
        msg = "Preflight results:" & vbCrLf & _
              "  Actionable: " & actionableCount & vbCrLf & _
              "  Ambiguous: " & ambiguousCount & vbCrLf & _
              "  Not Found: " & notFoundCount & vbCrLf & _
              "  No-Op: " & noopCount & vbCrLf & vbCrLf & _
              "Proceed with actionable items?" & vbCrLf & _
              "(Missing contexts will be reviewed at the end)"
        Dim resp As VbMsgBoxResult
        resp = MsgBox(msg, vbQuestion + vbYesNoCancel, "Preflight Analyzer")
        If resp = vbCancel Then Exit Sub
        Dim toReview As New Collection
        Dim item As Object
        
        If resp = vbYes Then
            ' Actionable + Not Found
            For Each item In analyzed("actionable")
                toReview.Add item
            Next item
        Else
            ' Found + Not Found (Review All, but move Not Found to end)
            For Each item In analyzed("found")
                toReview.Add item
            Next item
        End If
        
        ' Append Not Found items at the end
        For Each item In analyzed("notFound")
            toReview.Add item
        Next item
        
        Call RunInteractiveReview(toReview, baseMatchCase, startTime)
    Else
        ' OLD: Tracked changes mode (preserved for fallback)
        Call RunTrackedChangesReview(TheForm, suggestions, startTime)
    End If
End Sub

' =========================================================================================
' === NEW: INTERACTIVE REVIEW WORKFLOW ===================================================
' =========================================================================================

Private Sub RunInteractiveReview(ByVal suggestions As Object, ByVal matchCase As Boolean, ByVal startTime As Single)
    ' This workflow shows each suggestion interactively and applies changes immediately
    ' based on user decisions. No tracked changes or comments are used.
    
    Dim searchRange As Range
    Set searchRange = ActiveDocument.Content
    
    Dim successCount As Long, skippedCount As Long, notFoundCount As Long, errorCount As Long
    Dim notFoundLog As String, errorLog As String
    Dim i As Long
    Dim userAction As String
    Dim acceptAll As Boolean
    
    acceptAll = False
    i = 1
    
    Dim suggestion As Object
    For Each suggestion In suggestions
        Debug.Print "  - Processing suggestion " & i & "/" & suggestions.Count & "..."
        
        ' Always call ShowSuggestionPreview. It will handle auto-skipping if acceptAll is True AND context is found.
        ' If context is NOT found, it will show the form even if acceptAll is True.
        userAction = ShowSuggestionPreview(suggestion, i, suggestions.Count, searchRange, matchCase, acceptAll)
        
        Select Case userAction
            Case "STOP"
                Debug.Print "User stopped processing at suggestion " & i
                Exit For
            Case "ACCEPT_ALL"
                acceptAll = True
                ' Fall through to accept this one
            Case "SKIP"
                skippedCount = skippedCount + 1
                i = i + 1
                GoTo NextSuggestion
            Case "REJECT"
                skippedCount = skippedCount + 1
                i = i + 1
                GoTo NextSuggestion
            Case "ACCEPT"
                ' Continue to apply
            Case Else
                ' Unknown action, skip
                skippedCount = skippedCount + 1
                i = i + 1
                GoTo NextSuggestion
        End Select
        
        ' Apply the suggestion
        Err.Clear
        On Error Resume Next
        Dim processedOk As Boolean
        processedOk = ProcessSuggestion(searchRange, suggestion, matchCase)
        
        If Err.Number <> 0 Then
            errorCount = errorCount + 1
            errorLog = errorLog & vbCrLf & "- Suggestion " & i & ": " & GetSuggestionContextText(suggestion) & " (Error: " & Err.Description & ")"
            Err.Clear
        ElseIf processedOk Then
            successCount = successCount + 1
        Else
            notFoundCount = notFoundCount + 1
            notFoundLog = notFoundLog & vbCrLf & "- Suggestion " & i & ": '" & GetSuggestionContextText(suggestion) & "'"
        End If
        
        On Error GoTo 0
        i = i + 1
NextSuggestion:
    Next suggestion
    
    ' --- Display final report ---
    Dim report As String
    report = "Interactive Review Complete!" & vbCrLf & vbCrLf
    report = report & "Summary:" & vbCrLf
    report = report & "  - Total Suggestions: " & suggestions.Count & vbCrLf
    report = report & "  - Applied: " & successCount & vbCrLf
    report = report & "  - Skipped/Rejected: " & skippedCount & vbCrLf
    report = report & "  - Not Found: " & notFoundCount & vbCrLf
    report = report & "  - Errors: " & errorCount & vbCrLf
    report = report & "  - Duration: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf
    
    If notFoundCount > 0 Then report = report & vbCrLf & "Contexts not found:" & notFoundLog
    If errorCount > 0 Then report = report & vbCrLf & "Errors encountered:" & errorLog
    
    Debug.Print "--- FINAL REPORT ---" & vbCrLf & report
    MsgBox report, vbInformation, "Review Complete"
End Sub

Private Function ShowSuggestionPreview(ByVal suggestion As Object, ByVal index As Long, _
                                       ByVal total As Long, ByVal searchRange As Range, _
                                       ByVal matchCase As Boolean, Optional ByVal autoAccept As Boolean = False) As String
    ' Shows the preview form and returns the user's action choice
    
    On Error GoTo ErrorHandler
    
    ' Declare form once at the top
    Dim previewForm As frmSuggestionPreview
    Dim context As String
    Dim contextForSearch As String
    Dim contextRange As Range
    Dim actionRange As Range
    Dim target As String
    Dim targetForSearch As String
    
    Dim effectiveMatchCase As Boolean
    effectiveMatchCase = matchCase
    If HasDictionaryKey(suggestion, "matchCase") Then
        On Error Resume Next
        effectiveMatchCase = CBool(suggestion("matchCase"))
        On Error GoTo ErrorHandler
    End If
    ' Find the context and action ranges
    context = GetSuggestionContextText(suggestion)
    contextForSearch = NormalizeForDocument(context)
    Set contextRange = FindWithProgressiveFallback(contextForSearch, searchRange, effectiveMatchCase, suggestion)

    If contextRange Is Nothing Then
        ' Context not found - show error in form
        ' even if autoAccept is True, because we can't apply it automatically.
        Set previewForm = New frmSuggestionPreview
        previewForm.LoadSuggestion suggestion, index, total, Nothing, Nothing
        previewForm.Show vbModeless

        ' Wait for user action (modeless form requires a wait loop)
        Dim waitCounter1 As Long
        waitCounter1 = 0
        Do While previewForm.UserAction = ""
            DoEvents
            waitCounter1 = waitCounter1 + 1
            If waitCounter1 Mod 100 = 0 Then Sleep 10  ' Yield CPU every 100 iterations
            If waitCounter1 > WAIT_LOOP_SAFETY_LIMIT Then
                Debug.Print "    -> SAFETY EXIT: Wait loop timeout in ShowSuggestionPreview (context not found)"
                ShowSuggestionPreview = "SKIP"
                On Error Resume Next
                Unload previewForm
                On Error GoTo ErrorHandler
                Exit Function
            End If
            If waitCounter1 Mod 10000 = 0 Then
                Debug.Print "    -> ShowSuggestionPreview: Still waiting for user action (" & waitCounter1 & " iterations)"
            End If
        Loop

        ShowSuggestionPreview = previewForm.UserAction
        Unload previewForm
        Exit Function
    Else
        ' Context Found
        If autoAccept Then
            ShowSuggestionPreview = "ACCEPT"
            Exit Function
        End If
    End If

    ' Determine action range (target within context, or whole context)
    If suggestion.Exists("target") Then
        target = GetSuggestionText(suggestion, "target", "")
        If Len(target) > 0 Then
            targetForSearch = NormalizeForDocument(target)
            Set actionRange = FindWithProgressiveFallback(targetForSearch, contextRange, effectiveMatchCase, suggestion)
        End If
    End If
    
    If actionRange Is Nothing Then
        Set actionRange = contextRange.Duplicate
    End If
    
    ' Show the form modeless (allows document interaction)
    Set previewForm = New frmSuggestionPreview
    previewForm.LoadSuggestion suggestion, index, total, contextRange, actionRange
    previewForm.Show vbModeless

    ' Wait for user action (modeless form requires a wait loop)
    Dim waitCounter2 As Long
    waitCounter2 = 0
    Do While previewForm.UserAction = ""
        DoEvents
        waitCounter2 = waitCounter2 + 1
        If waitCounter2 Mod 100 = 0 Then Sleep 10  ' Yield CPU every 100 iterations
        If waitCounter2 > WAIT_LOOP_SAFETY_LIMIT Then
            Debug.Print "    -> SAFETY EXIT: Wait loop timeout in ShowSuggestionPreview (main path)"
            ShowSuggestionPreview = "SKIP"
            On Error Resume Next
            Unload previewForm
            On Error GoTo ErrorHandler
            Exit Function
        End If
        If waitCounter2 Mod 10000 = 0 Then
            Debug.Print "    -> ShowSuggestionPreview: Still waiting for user action (" & waitCounter2 & " iterations)"
        End If
    Loop

    ShowSuggestionPreview = previewForm.UserAction
    Unload previewForm
    Exit Function

ErrorHandler:
    Debug.Print "Error in ShowSuggestionPreview: " & Err.Description
    ShowSuggestionPreview = "SKIP"
End Function

' =========================================================================================
' === OLD: TRACKED CHANGES REVIEW WORKFLOW (PRESERVED) ===================================
' =========================================================================================

Private Sub RunTrackedChangesReview(ByVal TheForm As frmJsonInput, ByVal suggestions As Object, ByVal startTime As Single)
    ' This is the ORIGINAL workflow that applies all changes as tracked changes,
    ' then opens the review form. Preserved for fallback/comparison.
    
    Dim userResponse As VbMsgBoxResult
    
    ' --- Store original user details ---
    Dim originalUserName As String, originalUserInitials As String
    originalUserName = Application.UserName
    originalUserInitials = Application.UserInitials

    ' --- Store original Track Revisions state ---
    Dim originalTrackRevisions As Boolean
    originalTrackRevisions = ActiveDocument.TrackRevisions

    ' --- Counters for the final report ---
    Dim successCount As Long, notFoundCount As Long, errorCount As Long, fallbackCount As Long
    Dim totalCount As Long
    Dim notFoundLog As String, errorLog As String, fallbackLog As String
    On Error GoTo FinalCleanup ' Ensure original settings are always restored

    totalCount = suggestions.Count

    ' --- Set AI Reviewer Identity ---
    ActiveDocument.TrackRevisions = True

    ' --- Process each suggestion sequentially ---
    Dim searchRange As Range
    Set searchRange = ActiveDocument.Content
    Dim suggestion As Object
    Dim i As Long
    i = 1
    For Each suggestion In suggestions
        Debug.Print "  - Processing item " & i & "/" & totalCount & "..."
        TheForm.lblProgress.Caption = "Processing " & i & " of " & totalCount & "..."
        DoEvents

        Err.Clear
        On Error Resume Next ' Isolate errors to a single item
        Dim processedOk As Boolean
        processedOk = ProcessSuggestion(searchRange, suggestion, CBool(TheForm.chkCaseSensitive.value))
        If Err.Number <> 0 Then
            errorCount = errorCount + 1
            errorLog = errorLog & vbCrLf & "- Context: '" & GetSuggestionContextText(suggestion) & "' (Error: " & Err.Description & ")"
            Err.Clear
        ElseIf processedOk Then
            successCount = successCount + 1
        Else
            Dim fallbackPlaced As Boolean
            fallbackPlaced = HandleNotFoundContext(searchRange, suggestion, CBool(TheForm.chkCaseSensitive.value))

            If fallbackPlaced Then
                fallbackCount = fallbackCount + 1
                fallbackLog = fallbackLog & vbCrLf & "- Context: '" & GetSuggestionContextText(suggestion) & "'"
            Else
                notFoundCount = notFoundCount + 1
                notFoundLog = notFoundLog & vbCrLf & "- '" & GetSuggestionContextText(suggestion) & "'"
            End If
        End If

        On Error GoTo FinalCleanup
        i = i + 1
    Next suggestion

    ' --- Display the final report ---
    Dim report As String
    report = "LLM Review Processing Complete!" & vbCrLf & vbCrLf
    report = report & "Summary:" & vbCrLf
    report = report & "  - Total Suggestions: " & totalCount & vbCrLf
    report = report & "  - Successfully Applied: " & successCount & vbCrLf
    report = report & "  - Fallback Comment Added: " & fallbackCount & vbCrLf
    report = report & "  - Context Not Found: " & notFoundCount & vbCrLf
    report = report & "  - Errors During Processing: " & errorCount & vbCrLf
    report = report & "  - Duration: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf
    If fallbackCount > 0 Then report = report & vbCrLf & "Fallback comments were added for the following contexts:" & fallbackLog
    If notFoundCount > 0 Then report = report & vbCrLf & "The following contexts could not be found:" & notFoundLog
    If errorCount > 0 Then report = report & vbCrLf & "The following items caused errors:" & errorLog
    Debug.Print "--- FINAL REPORT ---" & vbCrLf & report
    report = report & vbCrLf & vbCrLf & "Do you want to begin reviewing the changes now?"

    TheForm.Hide
    userResponse = MsgBox(report, vbQuestion + vbYesNo, "Processing Complete")
    If userResponse = vbYes Then
        StartAiReview
    End If

FinalCleanup:
    ' --- ALWAYS restore original user settings ---
    Application.UserName = originalUserName
    Application.UserInitials = originalUserInitials
    ActiveDocument.TrackRevisions = originalTrackRevisions
    If Err.Number <> 0 Then
        HandleError "RunTrackedChangesReview", Err
    End If
End Sub

' =========================================================================================

' === GENERIC ERROR HANDLER ===============================================================

' =========================================================================================

Private Sub HandleError(ByVal procedureName As String, ByVal errSource As ErrObject)
    ' Provides a detailed, consistent error message and logs to the Immediate Window.
    Dim msg As String
    msg = "An unexpected error occurred in: " & procedureName & vbCrLf & vbCrLf
    msg = msg & "Error Number: " & errSource.Number & vbCrLf
    msg = msg & "Description: " & errSource.Description & vbCrLf & vbCrLf
    msg = msg & "The macro will now stop."

    ' --- NEW: Output to Immediate Window for easy copy/paste ---
    Debug.Print "--- ERROR --- " & Now() & " ---"
    Debug.Print "Procedure: " & procedureName
    Debug.Print "Error #: " & errSource.Number
    Debug.Print "Description: " & errSource.Description
    Debug.Print "---------------------------------"
    MsgBox msg, vbCritical, "Runtime Error in " & procedureName

    ' Optional: Uncomment the line below to automatically break at the error location
    ' when debugging from the VBA editor.
    ' Stop

End Sub

' =========================================================================================

Private Function ProcessSuggestion(ByRef searchRange As Range, ByVal suggestion As Object, ByVal matchCase As Boolean) As Boolean
    On Error GoTo ErrorHandler
    If TypeName(suggestion) = "Collection" Then
        Err.Raise vbObjectError + 509, "ProcessSuggestion", "Each suggestion must be a JSON object, not an array."
    End If

    If Not HasDictionaryKey(suggestion, "context") Then
        Err.Raise vbObjectError + 510, "ProcessSuggestion", "Suggestion is missing the required 'context' field."
    End If

    Dim context As String
    context = suggestion("context")

    Dim contextForSearch As String
    contextForSearch = NormalizeForDocument(context)
    Dim effectiveMatchCase As Boolean
    effectiveMatchCase = matchCase
    If HasDictionaryKey(suggestion, "matchCase") Then
        On Error Resume Next
        effectiveMatchCase = CBool(suggestion("matchCase"))
        On Error GoTo ErrorHandler
    End If

    ' Find the main context range just once.
    Dim contextRange As Range
    Set contextRange = FindWithProgressiveFallback(contextForSearch, searchRange, effectiveMatchCase, suggestion)

    If contextRange Is Nothing Then
        ' Let the main loop handle the fallback logic.
        ProcessSuggestion = False
        Exit Function
    End If
    Debug.Print "  - RESULT: Context anchor found at position " & contextRange.Start

    ' Now, decide which kind of suggestion this is.
    Dim undoStarted As Boolean
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "AI Suggestion"
    undoStarted = (Err.Number = 0)
    Err.Clear
    On Error GoTo ErrorHandler
    If suggestion.Exists("actions") Then
        ' === COMPOUND ACTION PATH ===
        Debug.Print "  - Processing as Compound Action..."
        Dim subActions As Object
        Set subActions = suggestion("actions")

        If TypeName(subActions) <> "Collection" Then
            Err.Raise vbObjectError + 511, "ProcessSuggestion", "The 'actions' field must be an array of sub-actions."
        End If
        If subActions.Count = 0 Then
            Err.Raise vbObjectError + 515, "ProcessSuggestion", "The 'actions' array must contain at least one sub-action."
        End If

        Dim cReplace As Collection, cStyle As Collection, cTable As Collection, cComment As Collection, cOther As Collection
        Set cReplace = New Collection
        Set cStyle = New Collection
        Set cTable = New Collection
        Set cComment = New Collection
        Set cOther = New Collection

        Dim actionObject As Object
        For Each actionObject In subActions
            Dim aName As String
            aName = LCase$(Trim$(GetSuggestionText(actionObject, "action", "")))
            Select Case aName
                Case "change", "replace"
                    cReplace.Add actionObject
                Case "apply_heading_style"
                    cStyle.Add actionObject
                Case "replace_with_table"
                    cTable.Add actionObject
                Case "comment"
                    cComment.Add actionObject
                Case "insert_table_row", "delete_table_row"
                    cOther.Add actionObject
                Case Else
                    cOther.Add actionObject
            End Select
        Next actionObject

        Dim currentContextRange As Range
        Set currentContextRange = contextRange.Duplicate

        Dim ao As Object
        For Each ao In cReplace
            Call ExecuteSingleAction(currentContextRange, ao, suggestion, effectiveMatchCase)
        Next ao
        For Each ao In cStyle
            Call ExecuteSingleAction(currentContextRange, ao, suggestion, effectiveMatchCase)
        Next ao
        For Each ao In cTable
            Call ExecuteSingleAction(currentContextRange, ao, suggestion, effectiveMatchCase)
        Next ao
        For Each ao In cComment
            Call ExecuteSingleAction(currentContextRange, ao, suggestion, effectiveMatchCase)
        Next ao
        For Each ao In cOther
            Call ExecuteSingleAction(currentContextRange, ao, suggestion, effectiveMatchCase)
        Next ao

    Else
        ' === SINGLE ACTION PATH (BACKWARDS COMPATIBILITY) ===
        Debug.Print "  - Processing as Single Action..."
        Call ExecuteSingleAction(contextRange, suggestion, suggestion, effectiveMatchCase)
    End If

    ' If we get here without an error, it was successful.
    ' For compound actions, propagate any context changes back to the outer range
    If suggestion.Exists("actions") Then
        Set contextRange = currentContextRange
    End If
    ' Advance the main search range PAST the context we just processed.
    searchRange.Start = contextRange.End
    ProcessSuggestion = True
    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Debug.Print "-> SUCCESS: Suggestion block applied for context '" & Left(context, 50) & "...'"
    Exit Function

ErrorHandler:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.source

    Debug.Print "--- PROCESS SUGGESTION ERROR --- " & Now() & " ---"
    Debug.Print "Error #: " & errNum
    Debug.Print "Source: " & errSrc
    Debug.Print "Description: " & errDesc
    Debug.Print "Context head: " & Left$(GetSuggestionContextText(suggestion), 120)
    Debug.Print "-------------------------------------------"

    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Err.Raise errNum, "ProcessSuggestion: " & errSrc, errDesc
End Function

Private Function HandleNotFoundContext(ByVal searchRange As Range, ByVal suggestion As Object, ByVal matchCase As Boolean) As Boolean
    ' Fallback logic for when the primary context is not found.
    ' Tries to find the first significant word and adds a comment there.
    ' Returns TRUE if a fallback comment was successfully placed, FALSE otherwise.

    On Error GoTo ErrorHandler
    HandleNotFoundContext = False ' Default to failure

    Dim context As String, explanation As String, replaceText As String, target As String
    context = GetSuggestionText(suggestion, "context", "<no context provided>")

    ' 1. Find the first word in the context > 4 letters, or just the first word.
    Dim keyword As String
    keyword = GetFirstSignificantWord(context)

    If Len(keyword) = 0 Then
        Debug.Print "  - FALLBACK: Could not extract a keyword from context: '" & context & "'"
        Exit Function
    End If

    ' 2. Search for this keyword in the document.
    Dim keywordRange As Range
    Dim keywordForSearch As String
    keywordForSearch = NormalizeForDocument(keyword)

    ' IMPORTANT: We search from the START of the main searchRange, not from a previous find.
    ' This ensures we find the first available anchor point for our fallback comment.
    Set keywordRange = FindLongString(keywordForSearch, searchRange, matchCase)

    If keywordRange Is Nothing Then
        Debug.Print "  - FALLBACK: Keyword '" & keyword & "' not found in the remaining document."
        Exit Function
    End If

    ' 3. Construct the comment text.
    explanation = GetSuggestionText(suggestion, "explanation", "<no explanation provided>")
    target = GetSuggestionText(suggestion, "target", "<no target specified>")
    replaceText = GetSuggestionText(suggestion, "replace", "<no replacement specified>")

    Dim commentText As String
    commentText = "AI SUGGESTION (CONTEXT NOT FOUND):" & vbCrLf & _
                  "------------------" & vbCrLf & _
                  "Original Context (not found): " & context & vbCrLf & _
                  "Intended Target: " & target & vbCrLf & _
                  "Suggested Replacement: " & replaceText & vbCrLf & _
                  "Explanation: " & explanation

    ' 4. Add the comment at the keyword's location.
    ActiveDocument.Comments.Add Range:=keywordRange, Text:=commentText
    Debug.Print "  - FALLBACK: Success. Added comment at position " & keywordRange.Start & " for keyword '" & keyword & "'"

    HandleNotFoundContext = True ' Signal success
    Exit Function

ErrorHandler:
    Debug.Print "An error occurred in HandleNotFoundContext: " & Err.Description
    HandleNotFoundContext = False
End Function

Private Sub ExecuteSingleAction(ByRef overallContextRange As Range, ByVal actionObject As Object, ByVal topLevelSuggestion As Object, ByVal matchCase As Boolean)
    On Error GoTo ErrorHandler
    Dim target As String, action As String, replaceText As String, explanation As String
    Dim actionRange As Range
    Dim targetStyle As Style
    Dim styleLookupError As Long
    Dim newTable As Table
    Dim actionValue As Variant, targetValue As Variant, replaceValue As Variant, explanationValue As Variant
    Dim occurrenceIndex As Long
    Dim targetForSearch As String

    ' 1. Extract details from the actionObject
    If Not actionObject.Exists("action") Then
        Err.Raise vbObjectError + 512, "ExecuteSingleAction", "Sub-action is missing the required 'action' field."
    End If

    actionValue = actionObject("action")
    If IsNull(actionValue) Then
        Err.Raise vbObjectError + 512, "ExecuteSingleAction", "Sub-action 'action' value cannot be Null."
    End If
    action = Trim$(CStr(actionValue))
    If Len(action) = 0 Then
        Err.Raise vbObjectError + 512, "ExecuteSingleAction", "Sub-action 'action' value cannot be empty."
    End If

    If actionObject.Exists("target") Then
        targetValue = actionObject("target")
        If IsNull(targetValue) Then
            target = ""
        Else
            target = CStr(targetValue)
        End If
    Else
        target = ""
    End If

    targetForSearch = ""
    If Len(target) > 0 Then targetForSearch = NormalizeForDocument(target)

    occurrenceIndex = 1
    If actionObject.Exists("occurrenceIndex") Then
        If Not IsNull(actionObject("occurrenceIndex")) Then occurrenceIndex = CLng(Val(CStr(actionObject("occurrenceIndex"))))
    ElseIf HasDictionaryKey(topLevelSuggestion, "occurrenceIndex") Then
        If Not IsNull(topLevelSuggestion("occurrenceIndex")) Then occurrenceIndex = CLng(Val(CStr(topLevelSuggestion("occurrenceIndex"))))
    End If

    If actionObject.Exists("replace") Then
        replaceValue = actionObject("replace")
        If IsNull(replaceValue) Then
            replaceText = ""
        Else
            replaceText = CStr(replaceValue)
        End If
    Else
        replaceText = ""
    End If

    If actionObject.Exists("explanation") Then
        explanationValue = actionObject("explanation")
    ElseIf HasDictionaryKey(topLevelSuggestion, "explanation") Then
        explanationValue = topLevelSuggestion("explanation")
    Else
        explanationValue = ""
    End If
    If IsNull(explanationValue) Then
        explanation = ""
    Else
        explanation = CStr(explanationValue)
    End If

    ' 2. Define the Action Range
    ' If tableCell is present, DO NOT rely on overallContextRange (it may point into TOC/headings).
    ' Instead, locate the cell directly by scanning forward in the document.
    If HasDictionaryKey(actionObject, "tableCell") Then
        Dim tableSearchRange As Range
        Dim startPos As Long
        startPos = overallContextRange.End
        If startPos < 0 Then startPos = 0
        If startPos > ActiveDocument.Content.End Then startPos = ActiveDocument.Content.End
        Set tableSearchRange = ActiveDocument.Range(startPos, ActiveDocument.Content.End)

        ' Allow per-action context anchors (preferred) and fall back to top-level context for table preference.
        If Not HasDictionaryKey(actionObject, "context") Then
            If HasDictionaryKey(topLevelSuggestion, "context") Then
                On Error Resume Next
                actionObject("context") = topLevelSuggestion("context")
                On Error GoTo ErrorHandler
            End If
        End If

        Set actionRange = FindTableCell(actionObject, tableSearchRange)
        If actionRange Is Nothing Then
            Err.Raise vbObjectError + 517, "ExecuteSingleAction", "Table cell not found for this action."
        End If
    Else
        ' Non-table actions operate within the overall context range
        If target <> "" Then
            If occurrenceIndex <= 1 Then
                Set actionRange = FindWithProgressiveFallback(targetForSearch, overallContextRange, matchCase, actionObject)
                If actionRange Is Nothing Then
                    Err.Raise vbObjectError + 517, "ExecuteSingleAction", "Target '" & target & "' not found within its context."
                End If
            Else
                Dim searchSubRange As Range
                Dim foundRange As Range
                Dim n As Long
                Set searchSubRange = overallContextRange.Duplicate
                For n = 1 To occurrenceIndex
                    Set foundRange = FindWithProgressiveFallback(targetForSearch, searchSubRange, matchCase, actionObject)
                    If foundRange Is Nothing Then
                        Err.Raise vbObjectError + 517, "ExecuteSingleAction", "Target occurrence " & occurrenceIndex & " not found within its context."
                    End If
                    If n < occurrenceIndex Then
                        searchSubRange.Start = foundRange.End
                        If searchSubRange.Start >= searchSubRange.End Then
                            Err.Raise vbObjectError + 517, "ExecuteSingleAction", "Target occurrence " & occurrenceIndex & " not found within its context (range exhausted)."
                        End If
                    End If
                Next n
                Set actionRange = foundRange
            End If
        Else
            Set actionRange = overallContextRange.Duplicate
        End If
    End If

    ' 3. Handle context located inside a table cell
    ' Logic removed to allow partial replacements within table cells
    ' If actionRange.Information(wdWithinTable) Then
    '    Set actionRange = actionRange.Cells(1).Range
    '    actionRange.End = actionRange.End - 1 ' Trim end-of-cell marker
    '    Debug.Print "Adjusted for table cell. New range: " & actionRange.Start & "-" & actionRange.End
    ' End If

    ' 3.1 If a tableCell structure was used to locate the range, refine to the specific target within that cell
    If HasDictionaryKey(actionObject, "tableCell") Then
        If target <> "" Then
            Dim cellSearchRange As Range
            Dim cellFoundRange As Range
            Dim cellN As Long

            Set cellSearchRange = actionRange.Duplicate
            Set cellFoundRange = Nothing

            For cellN = 1 To occurrenceIndex
                Set cellFoundRange = FindLongString(targetForSearch, cellSearchRange, matchCase)
                If cellFoundRange Is Nothing Then
                    Err.Raise vbObjectError + 517, "ExecuteSingleAction", "Target '" & target & "' not found within the specified table cell."
                End If

                If cellN < occurrenceIndex Then
                    cellSearchRange.Start = cellFoundRange.End
                    If cellSearchRange.Start >= cellSearchRange.End Then
                        Err.Raise vbObjectError + 517, "ExecuteSingleAction", "Target occurrence " & occurrenceIndex & " not found within the specified table cell (range exhausted)."
                    End If
                End If
            Next cellN

            Set actionRange = cellFoundRange
        End If
    End If

    ' 4. Perform the specified action
    Select Case LCase(action)
        Case "change", "replace", "set_table_cell"
            If LCase(action) = "set_table_cell" Then
                If Not HasDictionaryKey(actionObject, "tableCell") Then
                    Err.Raise vbObjectError + 517, "ExecuteSingleAction", "set_table_cell requires 'tableCell' structure."
                End If
                If Len(Trim$(target)) > 0 Then
                    Err.Raise vbObjectError + 517, "ExecuteSingleAction", "set_table_cell requires empty 'target'."
                End If
            End If
            ' Check if the formatting is already applied (skip if it is)
            If IsFormattingAlreadyApplied(actionRange, replaceText) Then
                Debug.Print "Action 'change': SKIPPED - Formatting already matches for '" & actionRange.Text & "'"
                ' Don't apply or add comment - formatting is already correct
            Else
                Debug.Print "Action 'change': Replacing '" & actionRange.Text & "' with '" & replaceText & "'"
                ApplyFormattedReplacement actionRange, replaceText
                If explanation <> "" Then
                    'ActiveDocument.Comments.Add Range:=actionRange, Text:="AI Suggestion: " & explanation
                End If
            End If

        Case "comment"
            ' Special case: a comment action always applies to the OVERALL context range.
            If Len(Trim$(explanation)) = 0 Then
                Err.Raise vbObjectError + 516, "ExecuteSingleAction", "Comment actions require an explanation."
            End If
            Debug.Print "Action 'comment': Adding comment to '" & overallContextRange.Text & "'"
            ActiveDocument.Comments.Add Range:=overallContextRange, Text:=explanation
            
        Case "apply_heading_style"
            Debug.Print "Action 'apply_heading_style': Checking style '" & replaceText & "' for '" & actionRange.Text & "'"
            If Len(Trim$(replaceText)) = 0 Then
                Err.Raise vbObjectError + 518, "ExecuteSingleAction", "Heading style actions require a style name in the 'replace' field."
            End If
            replaceText = Trim$(replaceText)

            ' NO-OP CHECK: Skip if style is already applied
            On Error Resume Next
            If actionRange.Style = replaceText Then
                Debug.Print "Action 'apply_heading_style': SKIPPED - Style already matches"
                On Error GoTo ErrorHandler
                Exit Sub
            End If
            On Error GoTo ErrorHandler

            ' NEW: Try house style function first, fall back to direct style application
            If ApplyHouseStyle(actionRange, replaceText) Then
                Debug.Print "Action 'apply_heading_style': Applied via house style function"
            Else
                ' Fallback: Direct style application
                Err.Clear
                On Error Resume Next ' Temporarily handle missing style error
                Set targetStyle = ActiveDocument.Styles(replaceText)
                styleLookupError = Err.Number
                On Error GoTo ErrorHandler ' Restore main handler
                If styleLookupError <> 0 Or targetStyle Is Nothing Then
                    Err.Raise vbObjectError + 513, "ExecuteSingleAction", "Style '" & replaceText & "' not found in the document."
                End If

                Debug.Print "Action 'apply_heading_style': Applying style '" & replaceText & "'"
                actionRange.Style = targetStyle
            End If

            If explanation <> "" Then
                'ActiveDocument.Comments.Add Range:=actionRange, Text:="AI Suggestion: " & explanation
            End If

        Case "replace_with_table"
            Debug.Print "Action 'replace_with_table': Replacing content with a table."
            If Len(Trim$(replaceText)) = 0 Then
                Err.Raise vbObjectError + 519, "ExecuteSingleAction", "Table actions require markdown content in the 'replace' field."
            End If

            ' ConvertMarkdownToTable now handles table detection and deletion internally
            Set newTable = ConvertMarkdownToTable(actionRange, replaceText)
            If newTable Is Nothing Then
                Err.Raise vbObjectError + 521, "ExecuteSingleAction", "Markdown table conversion failed or returned no data."
            End If

            Set actionRange = newTable.Range
            Set overallContextRange = newTable.Range
            If explanation <> "" Then
                'ActiveDocument.Comments.Add Range:=actionRange, Text:="AI Suggestion: " & explanation
            End If

        Case "insert_table_row"
            Debug.Print "Action 'insert_table_row': Adding new row to table."
            If Not HasDictionaryKey(actionObject, "tableCell") Then
                Err.Raise vbObjectError + 522, "ExecuteSingleAction", "insert_table_row requires 'tableCell' structure."
            End If

            Dim insertPos As String
            insertPos = "after" ' default
            If HasDictionaryKey(actionObject, "insertPosition") Then
                insertPos = LCase$(Trim$(GetSuggestionText(actionObject, "insertPosition", "after")))
            End If

            Call InsertTableRow(actionObject, topLevelSuggestion, matchCase, insertPos, replaceText)

        Case "delete_table_row"
            Debug.Print "Action 'delete_table_row': Removing row from table."
            If Not HasDictionaryKey(actionObject, "tableCell") Then
                Err.Raise vbObjectError + 523, "ExecuteSingleAction", "delete_table_row requires 'tableCell' structure."
            End If

            Call DeleteTableRow(actionObject, topLevelSuggestion, matchCase)

        Case Else
            Err.Raise vbObjectError + 514, "ExecuteSingleAction", "Unsupported action type: '" & action & "'"
    End Select
    
    Exit Sub ' Success
    
ErrorHandler:
    ' Let the error bubble up to the caller (ProcessSuggestion) to be handled there.
    Err.Raise Err.Number, "ExecuteSingleAction: " & Err.source, Err.Description
End Sub

Private Sub pv_EatWhitespace()
    On Error GoTo ErrorHandler

'    Debug.Print "    >> pv_EatWhitespace (start index=" & p_Index & ")"

    Do While p_Index <= Len(p_Json)
        Select Case Mid$(p_Json, p_Index, 1)
            Case " ", vbCr, vbLf, vbTab
                p_Index = p_Index + 1
            Case Else
                Exit Do
        End Select

    Loop

'    Debug.Print "    << pv_EatWhitespace (end index=" & p_Index & ")"

    Exit Sub
ErrorHandler:
    HandleError "pv_EatWhitespace", Err
End Sub

Private Function pv_ParseObject() As Object
    On Error GoTo ErrorHandler
    Dim dict As Object
    Set dict = NewDictionary()
    p_Index = p_Index + 1 ' skip "{"
    Do
        pv_EatWhitespace
        If p_Index > Len(p_Json) Then
            p_ParseError = PARSE_UNEXPECTED_END_OF_INPUT
            Exit Function
        End If

        Dim NextChar As String
        NextChar = Mid$(p_Json, p_Index, 1)
        If NextChar = "}" Then
            p_Index = p_Index + 1
            Exit Do
        End If

        Dim Key As String
        Key = pv_ParseString()
        If p_ParseError <> PARSE_SUCCESS Then Exit Function

'        Debug.Print "       object: key='" & Key & "' at index=" & p_Index

        pv_EatWhitespace

'        Debug.Print "       object: expecting ':' got '" & Mid$(p_Json, p_Index, 1) & "' at index=" & p_Index

        If Mid$(p_Json, p_Index, 1) <> ":" Then
            p_ParseError = PARSE_INVALID_CHARACTER
            Exit Function
        End If

        p_Index = p_Index + 1
        Dim vVal As Variant

'        Debug.Print "       object: parsing value for key '" & Key & "' at index=" & p_Index

        llm_ParseValue vVal
        If p_ParseError <> PARSE_SUCCESS Then Exit Function
        dict.Add Key, vVal
        pv_EatWhitespace
        NextChar = Mid$(p_Json, p_Index, 1)

'        Debug.Print "       object: post-value next='" & NextChar & "' at index=" & p_Index

        If NextChar = "," Then
            p_Index = p_Index + 1
        ElseIf NextChar <> "}" Then
            p_ParseError = PARSE_INVALID_CHARACTER
            Exit Function
        End If

    Loop
    Set pv_ParseObject = dict
    Exit Function
ErrorHandler:
    HandleError "pv_ParseObject", Err
End Function

Private Function pv_ParseArray() As Object
    On Error GoTo ErrorHandler
    Dim arr As VBA.Collection
    Set arr = New VBA.Collection
    p_Index = p_Index + 1 ' skip "["
    Do
        pv_EatWhitespace
        If p_Index > Len(p_Json) Then
            p_ParseError = PARSE_UNEXPECTED_END_OF_INPUT
            Exit Function
        End If

        Dim NextChar As String
        NextChar = Mid$(p_Json, p_Index, 1)
        If NextChar = "]" Then
            p_Index = p_Index + 1
            Exit Do
        End If

        Dim vElem As Variant
        llm_ParseValue vElem
        If p_ParseError <> PARSE_SUCCESS Then Exit Function

'        Debug.Print "       array: adding element"

        arr.Add vElem
        If Err.Number <> 0 Then
            Debug.Print "       array: arr.Add failed Err " & Err.Number & ": " & Err.Description
            On Error GoTo ErrorHandler
            HandleError "pv_ParseArray_Add", Err
            Exit Function
        End If

        On Error GoTo ErrorHandler
        pv_EatWhitespace
        NextChar = Mid$(p_Json, p_Index, 1)

'        Debug.Print "       array: post-elem next='" & NextChar & "' at index=" & p_Index

        If NextChar = "," Then
            p_Index = p_Index + 1
        ElseIf NextChar = "]" Then
            p_Index = p_Index + 1
            Exit Do
        Else
            p_ParseError = PARSE_INVALID_CHARACTER
            Exit Function
        End If

    Loop
    Set pv_ParseArray = arr
    Exit Function
ErrorHandler:
    HandleError "pv_ParseArray", Err
End Function

Private Function pv_ParseString() As String
    On Error GoTo ErrorHandler

    ' advance past opening quote

    p_Index = p_Index + 1
    Dim EndIndex As Long
    EndIndex = InStr(p_Index, p_Json, Chr$(34))
    Do While EndIndex > 0
        If EndIndex <= 1 Then Exit Do
        If Mid$(p_Json, EndIndex - 1, 1) <> "\" Then Exit Do
        EndIndex = InStr(EndIndex + 1, p_Json, Chr$(34))
    Loop
    pv_ParseString = Mid$(p_Json, p_Index, EndIndex - p_Index)
    p_Index = EndIndex + 1

    ' Use explicit characters to avoid any name/escape ambiguity

    Dim bs As String: bs = Chr$(92) ' backslash
    pv_ParseString = VBA.Replace(pv_ParseString, bs & Chr$(34), Chr$(34)) ' \" -> "
    pv_ParseString = VBA.Replace(pv_ParseString, bs & bs, bs)             ' \\ -> \
    pv_ParseString = VBA.Replace(pv_ParseString, bs & "/", "/")         ' \/ -> /
    pv_ParseString = VBA.Replace(pv_ParseString, bs & "b", vbBack)
    pv_ParseString = VBA.Replace(pv_ParseString, bs & "f", vbFormFeed)
    pv_ParseString = VBA.Replace(pv_ParseString, bs & "n", vbLf)
    pv_ParseString = VBA.Replace(pv_ParseString, bs & "r", vbCr)
    pv_ParseString = VBA.Replace(pv_ParseString, bs & "t", vbTab)
    Exit Function
ErrorHandler:
    HandleError "pv_ParseString", Err
End Function

Private Function pv_ParseBoolean() As Boolean
    If LCase$(Mid$(p_Json, p_Index, 4)) = "true" Then
        pv_ParseBoolean = True
        p_Index = p_Index + 4
    Else
        pv_ParseBoolean = False
        p_Index = p_Index + 5
    End If

End Function

Private Function pv_ParseNull() As Variant
    p_Index = p_Index + 4
    pv_ParseNull = Null
End Function

Private Function ConvertMarkdownToTable(ByVal targetRange As Range, ByVal markdown As String) As Table

    ' This helper function converts a Markdown table string into a Word table.

    On Error GoTo ErrorHandler
    Dim lines() As String
    Dim processedLines As String
    Dim line As Variant
    Dim i As Long
    Dim tempLine As String
    Dim normalizedMarkdown As String
    Dim newTable As Table
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim wasInTable As Boolean
    Dim preservedStyle As Variant

    ' Store the original range boundaries
    rangeStart = targetRange.Start
    rangeEnd = targetRange.End

    ' Check if the target range is within a table
    wasInTable = targetRange.Information(wdWithinTable)

    ' If the range is within a table, we need to delete the entire table first
    If wasInTable Then
        Debug.Print "  - Target is within a table. Deleting the entire table before inserting new one."
        Dim oldTable As Table
        Set oldTable = targetRange.Tables(1)

        ' Preserve the table style before deleting
        On Error Resume Next
        preservedStyle = oldTable.Style
        Debug.Print "  - Preserving table style: " & CStr(preservedStyle)
        On Error GoTo ErrorHandler

        ' Get the range that the table occupies
        rangeStart = oldTable.Range.Start
        rangeEnd = oldTable.Range.End

        ' Delete the table completely
        oldTable.Delete

        ' Create a new range at the position where the table was
        Set targetRange = ActiveDocument.Range(rangeStart, rangeStart)
    End If

    ' 1. Normalize incoming line endings and split into individual lines.

    normalizedMarkdown = Replace(markdown, vbCrLf, vbLf)
    normalizedMarkdown = Replace(normalizedMarkdown, vbCr, vbLf)
    lines = Split(normalizedMarkdown, vbLf)
    processedLines = ""
    i = 0
    For Each line In lines
        tempLine = Replace(line, vbCr, "")
        tempLine = Trim$(tempLine)

        ' 2. Skip the markdown separator line (e.g., |---|---| or |:---|:---:|)

        If InStr(1, tempLine, "---") > 0 And Left$(tempLine, 1) = "|" Then

            ' Skip separator rows

        ElseIf Len(tempLine) > 0 Then

            ' 3. Process a content line

            If Left$(tempLine, 1) = "|" Then tempLine = Mid$(tempLine, 2)
            If Right$(tempLine, 1) = "|" Then tempLine = Left$(tempLine, Len(tempLine) - 1)
            tempLine = Replace(tempLine, "|", vbTab)
            Dim parts() As String
            Dim j As Long
            parts = Split(tempLine, vbTab)
            For j = LBound(parts) To UBound(parts)
                parts(j) = Trim$(parts(j))
            Next j
            tempLine = Join(parts, vbTab)
            If i > 0 Then processedLines = processedLines & vbCr
            processedLines = processedLines & tempLine
            i = i + 1
        End If

    Next line
    If Len(processedLines) = 0 Then
        Err.Raise vbObjectError + 520, "ConvertMarkdownToTable", "No usable markdown rows were found in the replacement text."
    End If

    Debug.Print "  - Inserting processed table data at position " & targetRange.Start
    targetRange.Text = processedLines
    Set newTable = targetRange.ConvertToTable(Separator:=vbTab, AutoFitBehavior:=wdAutoFitContent)

    ' Apply table style: use preserved style if available, otherwise use default
    On Error Resume Next
    If Not IsEmpty(preservedStyle) Then
        Debug.Print "  - Restoring preserved table style: " & CStr(preservedStyle)
        newTable.Style = preservedStyle
    Else
        Debug.Print "  - Applying default table style: Table Grid"
        newTable.Style = "Table Grid"
    End If
    On Error GoTo ErrorHandler

    ApplyFormattingToTableCells newTable
    Debug.Print "  - New table created successfully with " & GetSafeTableRowCount(newTable) & " rows and " & GetSafeTableColCount(newTable) & " columns."
    Set ConvertMarkdownToTable = newTable
    Exit Function
ErrorHandler:
    Debug.Print "An error occurred in ConvertMarkdownToTable: " & Err.Description
    Set ConvertMarkdownToTable = Nothing
    Err.Raise Err.Number, "ConvertMarkdownToTable", Err.Description
End Function

Private Sub ApplyFormattingToTableCells(ByVal sourceTable As Table)
    Dim tblCell As Cell
    Dim cellContent As Range
    Dim cellText As String
    For Each tblCell In sourceTable.Range.Cells
        Set cellContent = tblCell.Range
        If cellContent.End > cellContent.Start Then
            cellContent.End = cellContent.End - 1
        End If

        If cellContent.End > cellContent.Start And Right$(cellContent.Text, 1) = Chr$(13) Then
            cellContent.End = cellContent.End - 1
        End If

        cellText = cellContent.Text
        If Len(cellText) > 0 Then
            ApplyFormattedReplacement cellContent, cellText
        End If

    Next tblCell
End Sub

' =========================================================================================
' === TABLE CELL FINDING FUNCTIONS =======================================================
' =========================================================================================

Private Function GetCellText(ByVal c As Cell) As String
    ' Extracts normalized text from a cell
    On Error Resume Next

    Dim txt As String
    txt = GetCellTextFinalView(c)

    ' Normalize whitespace
    txt = NormalizeForDocument(txt)

    GetCellText = txt
End Function

Private Function FindCellByAdjacentContent(ByVal tbl As Table, _
                                           ByVal adjacentCells As Object, _
                                           ByVal hintRow As Long, _
                                           ByVal hintCol As Long, _
                                           ByVal hasHintRow As Boolean, _
                                           ByVal hasHintCol As Boolean) As Cell
    ' Uses adjacent cell content to pinpoint the exact cell

    On Error GoTo Fail

    Dim cellContent As String
    Dim aboveText As String
    Dim belowText As String
    Dim leftText As String
    Dim rightText As String

    ' Extract what we're looking for
    cellContent = GetSuggestionText(adjacentCells, "cellContent", "")
    aboveText = GetSuggestionText(adjacentCells, "above", "")
    belowText = GetSuggestionText(adjacentCells, "below", "")
    leftText = GetSuggestionText(adjacentCells, "left", "")
    rightText = GetSuggestionText(adjacentCells, "right", "")

    ' If we have row/col hints, start there
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim maxRows As Long
    Dim maxCols As Long

    maxRows = GetSafeTableRowCount(tbl)
    maxCols = GetSafeTableColCount(tbl)

    If maxRows = 0 Or maxCols = 0 Then
        Set FindCellByAdjacentContent = Nothing
        Exit Function
    End If

    If hasHintRow Then
        startRow = hintRow
        endRow = hintRow
    Else
        startRow = 1
        endRow = maxRows
    End If

    If hasHintCol Then
        startCol = hintCol
        endCol = hintCol
    Else
        startCol = 1
        endCol = maxCols
    End If

    If startRow < 1 Then startRow = 1
    If endRow > maxRows Then endRow = maxRows
    If startCol < 1 Then startCol = 1
    If endCol > maxCols Then endCol = maxCols
    If startRow > endRow Or startCol > endCol Then
        Set FindCellByAdjacentContent = Nothing
        Exit Function
    End If

    ' Search through the candidate cells
    Dim r As Long
    Dim c As Long

    Dim hasOtherAnchors As Boolean
    hasOtherAnchors = (Len(aboveText) > 0 Or Len(belowText) > 0 Or Len(leftText) > 0 Or Len(rightText) > 0 Or hasHintRow Or hasHintCol)

    Dim pass As Long
    Dim requireCellContent As Boolean
    Dim bestCandidate As Cell
    Set bestCandidate = Nothing

    For pass = 1 To 2
        requireCellContent = (pass = 1)
        If pass = 2 Then
            If Len(cellContent) = 0 Or Not hasOtherAnchors Then Exit For
            Set bestCandidate = Nothing
        End If

        For r = startRow To endRow
            For c = startCol To endCol
                Dim currentCell As Cell
                If Not TryGetTableCell(tbl, r, c, currentCell) Then GoTo NextCell

                ' Check if this cell matches all criteria
                Dim matches As Boolean
                matches = True

                Dim currentText As String
                currentText = GetCellText(currentCell)

                ' Check cell content (Option B: treat as a hint when other anchors exist)
                If requireCellContent Then
                    If Len(cellContent) > 0 Then
                        If Not TextMatchesHeuristic(cellContent, currentText) Then
                            matches = False
                        End If
                    End If
                End If

                ' Check above
                If matches And Len(aboveText) > 0 And r > 1 Then
                    Dim aboveCell As Cell
                    If TryGetTableCell(tbl, r - 1, c, aboveCell) Then
                        If Not TextMatchesHeuristic(aboveText, GetCellText(aboveCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                End If

                ' Check below
                If matches And Len(belowText) > 0 And r < maxRows Then
                    Dim belowCell As Cell
                    If TryGetTableCell(tbl, r + 1, c, belowCell) Then
                        If Not TextMatchesHeuristic(belowText, GetCellText(belowCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                End If

                ' Check left
                If matches And Len(leftText) > 0 And c > 1 Then
                    Dim leftCell As Cell
                    If TryGetTableCell(tbl, r, c - 1, leftCell) Then
                        If Not TextMatchesHeuristic(leftText, GetCellText(leftCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                End If

                ' Check right
                If matches And Len(rightText) > 0 And c < maxCols Then
                    Dim rightCell As Cell
                    If TryGetTableCell(tbl, r, c + 1, rightCell) Then
                        If Not TextMatchesHeuristic(rightText, GetCellText(rightCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                End If

                If matches Then
                    If pass = 1 Then
                        Set FindCellByAdjacentContent = currentCell
                        Exit Function
                    End If

                    If Len(cellContent) = 0 Then
                        Set FindCellByAdjacentContent = currentCell
                        Exit Function
                    End If

                    If TextMatchesHeuristic(cellContent, currentText) Then
                        Set FindCellByAdjacentContent = currentCell
                        Exit Function
                    End If

                    If bestCandidate Is Nothing Then
                        Set bestCandidate = currentCell
                    End If
                End If

NextCell:
            Next c
        Next r

        If pass = 2 Then
            If Not bestCandidate Is Nothing Then
                Set FindCellByAdjacentContent = bestCandidate
                Exit Function
            End If
        End If
    Next pass

    ' Not found
    Set FindCellByAdjacentContent = Nothing
    Exit Function

Fail:
    Set FindCellByAdjacentContent = Nothing
End Function

Private Function FindCellInTableRobust(ByVal tbl As Table, _
                                       ByVal rowHeader As String, _
                                       ByVal colHeader As String, _
                                       ByVal tableCellInfo As Object) As Range
    On Error GoTo Fail

    Dim foundColHeader As Range
    Dim finalCell As Range
    Dim hintRow As Long
    Dim hintCol As Long
    Dim hasHintRow As Boolean
    Dim hasHintCol As Boolean
    Dim targetRowNum As Long
    Dim hasTargetRow As Boolean
    Dim maxRows As Long
    Dim maxCols As Long

    hasHintRow = False
    hasHintCol = False
    hasTargetRow = False
    maxRows = GetSafeTableRowCount(tbl)
    maxCols = GetSafeTableColCount(tbl)

    If maxRows = 0 Or maxCols = 0 Then
        Set FindCellInTableRobust = Nothing
        Exit Function
    End If

    If Len(rowHeader) > 0 Then
        Dim r As Long
        For r = 1 To maxRows
            Dim rowFirstText As String
            rowFirstText = ""
            Dim rowFirstCell As Cell
            If TryGetFirstCellInRow(tbl, r, rowFirstCell) Then
                rowFirstText = GetCellText(rowFirstCell)
            End If

            If TextMatchesHeuristic(rowHeader, rowFirstText) Then
                targetRowNum = r
                hasTargetRow = True
                hintRow = r
                hasHintRow = True
                Exit For
            End If
        Next r
        If Not hasTargetRow Then
            Debug.Print "    -> Row Header '" & rowHeader & "' not found in table."
            Set FindCellInTableRobust = Nothing
            Exit Function
        End If
    End If

    If Len(colHeader) > 0 Then
        Dim headerCell As Cell
        Dim headerRowScan As Long
        Dim headerText As String

        For headerRowScan = 1 To maxRows
            If headerRowScan > 3 Then Exit For
            For Each headerCell In tbl.Range.Cells
                If headerCell.RowIndex = headerRowScan Then
                    headerText = GetCellText(headerCell)
                    If TextMatchesHeuristic(colHeader, headerText) Then
                        Set foundColHeader = headerCell.Range
                        hintCol = headerCell.ColumnIndex
                        hasHintCol = True
                        Exit For
                    End If
                End If
            Next headerCell
            If Not foundColHeader Is Nothing Then Exit For
        Next headerRowScan

        If foundColHeader Is Nothing Then
            Debug.Print "    -> Column Header '" & colHeader & "' not found in table."
        End If
    End If

    If HasDictionaryKey(tableCellInfo, "adjacentCells") Then
        Dim adj As Object
        Dim adjCell As Cell
        Set adj = tableCellInfo("adjacentCells")
        Set adjCell = FindCellByAdjacentContent(tbl, adj, hintRow, hintCol, hasHintRow, hasHintCol)
        If Not adjCell Is Nothing Then
            Set finalCell = adjCell.Range
        End If
    End If

    If finalCell Is Nothing And hasTargetRow Then
        Dim cell As Cell
        Dim cellIdx As Long
        cellIdx = 0

        For Each cell In tbl.Range.Cells
            If cell.RowIndex = targetRowNum Then
                cellIdx = cellIdx + 1

                If Not foundColHeader Is Nothing Then
                    If PositionsRoughlyAlign(cell.Range, foundColHeader) Then
                        Set finalCell = cell.Range
                        Exit For
                    End If
                End If

                If Len(colHeader) = 0 And cellIdx = 2 Then
                    Set finalCell = cell.Range
                End If
            End If
        Next cell
    End If

    If finalCell Is Nothing And Not foundColHeader Is Nothing Then
        ' As a fallback, scan all cells for the aligned column header position.
        For Each cell In tbl.Range.Cells
            If PositionsRoughlyAlign(cell.Range, foundColHeader) Then
                Set finalCell = cell.Range
                Exit For
            End If
        Next cell
    End If

    If Not finalCell Is Nothing Then
        Dim contentCell As Cell
        If finalCell.Cells.Count > 0 Then
            Set contentCell = finalCell.Cells(1)
            Dim contentRange As Range
            Set contentRange = GetCellContentRange(contentCell)
            If Not contentRange Is Nothing Then
                Set FindCellInTableRobust = contentRange
                Exit Function
            End If
        End If

        If finalCell.End > finalCell.Start Then finalCell.End = finalCell.End - 1
        Set FindCellInTableRobust = finalCell
    Else
        Debug.Print "    -> Cell not found after row/column matching."
        Set FindCellInTableRobust = Nothing
    End If
    Exit Function

Fail:
    Debug.Print "    -> Error in FindCellInTableRobust: " & Err.Description
    Set FindCellInTableRobust = Nothing
End Function

Private Function FindTableCell(ByVal suggestion As Object, ByVal searchRange As Range) As Range
    ' Returns the Range of the specific table cell, or Nothing if not found
    ' NEW: Prioritizes tableTitle field for deterministic table lookup

    On Error GoTo ErrorHandler

    ' 1. Check if suggestion has tableCell structure
    If Not HasDictionaryKey(suggestion, "tableCell") Then
        Set FindTableCell = Nothing
        Exit Function
    End If

    Dim tableCellInfo As Object
    Set tableCellInfo = suggestion("tableCell")

    ' 2. Extract search parameters
    Dim rowHeader As String
    Dim columnHeader As String
    Dim foundCell As Range
    Dim tableTitle As String

    rowHeader = GetSuggestionText(tableCellInfo, "rowHeader", "")
    columnHeader = GetSuggestionText(tableCellInfo, "columnHeader", "")
    tableTitle = GetSuggestionText(tableCellInfo, "tableTitle", "")

    Debug.Print "    -> Searching for table cell with:"
    If Len(tableTitle) > 0 Then Debug.Print "       tableTitle: " & tableTitle
    If Len(rowHeader) > 0 Then Debug.Print "       rowHeader: " & rowHeader
    If Len(columnHeader) > 0 Then Debug.Print "       columnHeader: " & columnHeader

    Dim contextAnchor As String
    contextAnchor = GetSuggestionText(suggestion, "context", "")

    Dim targetTable As Table
    Set targetTable = Nothing

    ' 3. NEW: Try tableTitle first (preferred method using indexed lookup)
    If Len(tableTitle) > 0 Then
        Dim matchCount As Long
        Set targetTable = FindTableByTitle(tableTitle, matchCount)

        If matchCount > 1 Then
            ' Multiple tables match - prompt user to select
            Debug.Print "    -> FindTableCell: Multiple tables (" & matchCount & ") match title '" & tableTitle & "'"
            Set targetTable = PromptUserToSelectTable(tableTitle, rowHeader, columnHeader)
            If targetTable Is Nothing Then
                Debug.Print "    -> FindTableCell: User cancelled or no selection made"
                Set FindTableCell = Nothing
                Exit Function
            End If
        ElseIf matchCount = 0 Then
            Debug.Print "    -> FindTableCell: No table found with title '" & tableTitle & "'"
            ' Fall through to try context anchor
        End If
    End If

    ' 4. Fallback: Try context anchor (original method)
    If targetTable Is Nothing And Len(contextAnchor) > 0 Then
        Dim ambiguousTitle As Boolean
        ambiguousTitle = False
        Set targetTable = ResolveTableFromContextAnchor(contextAnchor, searchRange, ambiguousTitle)
        If ambiguousTitle Then
            Debug.Print "    -> FindTableCell: Ambiguous table caption/title '" & contextAnchor & "'"
            ' Try to resolve with user selection
            Set targetTable = PromptUserToSelectTable(contextAnchor, rowHeader, columnHeader)
            If targetTable Is Nothing Then
                Set FindTableCell = Nothing
                Exit Function
            End If
        End If
    End If

    If targetTable Is Nothing Then
        Debug.Print "    -> FindTableCell: Could not locate a specific table"
        Set FindTableCell = Nothing
        Exit Function
    End If

    Set foundCell = FindCellInTableRobust(targetTable, rowHeader, columnHeader, tableCellInfo)

    If Not foundCell Is Nothing Then
        Set FindTableCell = foundCell
        Exit Function
    End If

    Debug.Print "    -> Table cell not found in target table"
    Set FindTableCell = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "    -> Error in FindTableCell: " & Err.Description
    Set FindTableCell = Nothing
End Function

Private Function CountTablesByCaptionAnchor(ByVal contextAnchor As String, ByVal searchRange As Range) As Long
    On Error GoTo Fail
    Dim c As Collection
    Set c = CollectTablesByCaptionAnchor(contextAnchor, searchRange)
    CountTablesByCaptionAnchor = c.Count
    Exit Function
Fail:
    CountTablesByCaptionAnchor = 0
End Function

Private Function CollectTablesByCaptionAnchor(ByVal contextAnchor As String, ByVal searchRange As Range) As Collection
    On Error GoTo Fail
    Dim matches As New Collection
    Dim dedupe As Object
    Set dedupe = NewDictionary()
    Dim tbl As Table
    For Each tbl In searchRange.Tables
        Dim foundCaption As Boolean
        foundCaption = False

        Dim p As Paragraph
        Dim i As Long
        Dim t As String

        On Error Resume Next
        Set p = tbl.Range.Paragraphs(1).Previous
        On Error GoTo Fail
        For i = 1 To 3
            If p Is Nothing Then Exit For
            t = Trim$(NormalizeForDocument(p.Range.Text))
            If Len(t) > 0 Then
                If TextMatchesHeuristic(contextAnchor, t) Then
                    foundCaption = True
                    Exit For
                End If
            End If
            On Error Resume Next
            Set p = p.Previous
            On Error GoTo Fail
        Next i

        If Not foundCaption Then
            On Error Resume Next
            Set p = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count).Next
            On Error GoTo Fail
            For i = 1 To 3
                If p Is Nothing Then Exit For
                t = Trim$(NormalizeForDocument(p.Range.Text))
                If Len(t) > 0 Then
                    If TextMatchesHeuristic(contextAnchor, t) Then
                        foundCaption = True
                        Exit For
                    End If
                End If
                On Error Resume Next
                Set p = p.Next
                On Error GoTo Fail
            Next i
        End If

        If foundCaption Then
            Dim k As String
            k = CStr(tbl.Range.Start) & ":" & CStr(tbl.Range.End)
            If Not dedupe.Exists(k) Then
                dedupe.Add k, True
                matches.Add tbl
            End If
        End If
    Next tbl
    Set CollectTablesByCaptionAnchor = matches
    Exit Function
Fail:
    Set CollectTablesByCaptionAnchor = New Collection
End Function

Private Function ResolveTableFromContextAnchor(ByVal contextAnchor As String, ByVal searchRange As Range, ByRef isAmbiguous As Boolean) As Table
    On Error GoTo Fail
    isAmbiguous = False
    Set ResolveTableFromContextAnchor = Nothing
    If Len(Trim$(contextAnchor)) = 0 Then Exit Function

    Dim captionMatches As Collection
    Set captionMatches = CollectTablesByCaptionAnchor(contextAnchor, searchRange)
    If captionMatches.Count = 1 Then
        Set ResolveTableFromContextAnchor = captionMatches(1)
        Exit Function
    ElseIf captionMatches.Count > 1 Then
        isAmbiguous = True
        Debug.Print "    -> ResolveTableFromContextAnchor: Caption/title matches multiple tables for '" & contextAnchor & "'"
        Dim t As Table
        For Each t In captionMatches
            Dim bt As String
            Dim at As String
            bt = "": at = ""
            Dim pp As Paragraph
            On Error Resume Next
            Set pp = t.Range.Paragraphs(1).Previous
            If Not pp Is Nothing Then bt = Trim$(NormalizeForDocument(pp.Range.Text))
            Set pp = t.Range.Paragraphs(t.Range.Paragraphs.Count).Next
            If Not pp Is Nothing Then at = Trim$(NormalizeForDocument(pp.Range.Text))
            On Error GoTo Fail
            If Len(bt) > 0 Then Debug.Print "       candidate caption(before): " & Left$(bt, 120) & IIf(Len(bt) > 120, "...", "")
            If Len(at) > 0 Then Debug.Print "       candidate caption(after):  " & Left$(at, 120) & IIf(Len(at) > 120, "...", "")
        Next t
        Exit Function
    End If

    Dim normAnchor As String
    normAnchor = NormalizeForDocument(contextAnchor)

    Dim work As Range
    Set work = searchRange.Duplicate
    Dim foundTable As Table
    Set foundTable = Nothing
    Dim resolveLoopCounter As Long
    Dim resolveLastPos As Long
    resolveLoopCounter = 0
    resolveLastPos = -1
    Do
        resolveLoopCounter = resolveLoopCounter + 1
        If resolveLoopCounter > LOOP_SAFETY_LIMIT Then
            Debug.Print "    -> SAFETY EXIT: ResolveTableFromContextAnchor exceeded " & LOOP_SAFETY_LIMIT & " iterations"
            Exit Do
        End If
        Dim anchorRange As Range
        Set anchorRange = FindLongString(normAnchor, work, False)
        If anchorRange Is Nothing Then Exit Do
        ' Prevent infinite loop if Find returns same position
        If anchorRange.Start = resolveLastPos Then
            Debug.Print "    -> SAFETY EXIT: ResolveTableFromContextAnchor stuck at position " & resolveLastPos
            Exit Do
        End If
        resolveLastPos = anchorRange.Start

        Dim candidateTable As Table
        Set candidateTable = GetTableFromAnchorRange(anchorRange)
        If Not candidateTable Is Nothing Then
            If foundTable Is Nothing Then
                Set foundTable = candidateTable
            Else
                If (candidateTable.Range.Start <> foundTable.Range.Start) Or (candidateTable.Range.End <> foundTable.Range.End) Then
                    isAmbiguous = True
                    Exit Do
                End If
            End If
        End If

        work.Start = anchorRange.End
        If work.Start >= work.End Then Exit Do
    Loop

    If isAmbiguous Then
        Set ResolveTableFromContextAnchor = Nothing
    Else
        Set ResolveTableFromContextAnchor = foundTable
    End If
    Exit Function
Fail:
    isAmbiguous = False
    Set ResolveTableFromContextAnchor = Nothing
End Function

Private Function GetTableFromAnchorRange(ByVal anchorRange As Range) As Table
    On Error GoTo Fail
    Set GetTableFromAnchorRange = Nothing
    If anchorRange Is Nothing Then Exit Function

    If anchorRange.Information(wdWithinTable) Then
        Set GetTableFromAnchorRange = anchorRange.Tables(1)
        Exit Function
    End If

    Dim p As Paragraph
    Dim i As Long

    On Error Resume Next
    Set p = anchorRange.Paragraphs(1).Previous
    On Error GoTo Fail
    For i = 1 To 3
        If p Is Nothing Then Exit For
        If p.Range.Information(wdWithinTable) Then
            Set GetTableFromAnchorRange = p.Range.Tables(1)
            Exit Function
        End If
        On Error Resume Next
        Set p = p.Previous
        On Error GoTo Fail
    Next i

    On Error Resume Next
    Set p = anchorRange.Paragraphs(anchorRange.Paragraphs.Count).Next
    On Error GoTo Fail
    For i = 1 To 3
        If p Is Nothing Then Exit For
        If p.Range.Information(wdWithinTable) Then
            Set GetTableFromAnchorRange = p.Range.Tables(1)
            Exit Function
        End If
        On Error Resume Next
        Set p = p.Next
        On Error GoTo Fail
    Next i
    Exit Function
Fail:
    Set GetTableFromAnchorRange = Nothing
End Function

Private Function StripTableNumberPrefix(ByVal captionText As String) As String
    On Error GoTo CleanFail
    Dim s As String
    s = Trim$(NormalizeForDocument(captionText))
    If Len(s) = 0 Then
        StripTableNumberPrefix = ""
        Exit Function
    End If

    Dim sLower As String
    sLower = LCase$(s)
    If Left$(sLower, 5) <> "table" Then
        StripTableNumberPrefix = s
        Exit Function
    End If

    s = Trim$(Mid$(s, 6))
    If Len(s) = 0 Then
        StripTableNumberPrefix = ""
        Exit Function
    End If

    Dim i As Long
    i = 1
    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    s = Trim$(Mid$(s, i))

    Do While Len(s) > 0
        ch = Left$(s, 1)
        If ch = "-" Or ch = ":" Then
            s = Trim$(Mid$(s, 2))
        Else
            Exit Do
        End If
    Loop

    If Len(s) = 0 Then
        StripTableNumberPrefix = Trim$(NormalizeForDocument(captionText))
    Else
        StripTableNumberPrefix = s
    End If
    Exit Function
CleanFail:
    StripTableNumberPrefix = Trim$(NormalizeForDocument(captionText))
End Function

Private Sub WarnUnusedAmbiguousTableTitles(ByVal searchRange As Range, ByVal usedCaptionKeys As Object)
    On Error GoTo CleanExit
    Dim captionCounts As Object
    Set captionCounts = BuildCaptionKeyCounts(searchRange)
    Dim k As Variant
    For Each k In captionCounts.Keys
        If CLng(captionCounts(k)) > 1 Then
            If usedCaptionKeys Is Nothing Then
                Debug.Print "  - WARNING: Ambiguous table title exists in document: '" & CStr(k) & "' (" & CLng(captionCounts(k)) & " occurrences)"
            ElseIf Not usedCaptionKeys.Exists(CStr(k)) Then
                Debug.Print "  - WARNING: Ambiguous table title exists but is not referenced by any suggestion: '" & CStr(k) & "' (" & CLng(captionCounts(k)) & " occurrences)"
            End If
        End If
    Next k
CleanExit:
End Sub

Private Function BuildCaptionKeyCounts(ByVal searchRange As Range) As Object
    On Error GoTo Fail
    Dim d As Object
    Set d = NewDictionary()
    Dim tbl As Table
    For Each tbl In searchRange.Tables
        Dim seen As Object
        Set seen = NewDictionary()

        Dim p As Paragraph
        Dim i As Long
        Dim t As String
        Dim keyText As String

        On Error Resume Next
        Set p = tbl.Range.Paragraphs(1).Previous
        On Error GoTo Fail
        For i = 1 To 3
            If p Is Nothing Then Exit For
            t = Trim$(NormalizeForDocument(p.Range.Text))
            If Len(t) > 0 Then
                keyText = Trim$(NormalizeForDocument(StripTableNumberPrefix(t)))
                If Len(keyText) > 0 Then
                    If Not seen.Exists(keyText) Then
                        seen.Add keyText, True
                        Call IncrementDictCount(d, keyText)
                    End If
                End If
            End If
            On Error Resume Next
            Set p = p.Previous
            On Error GoTo Fail
        Next i

        On Error Resume Next
        Set p = tbl.Range.Paragraphs(tbl.Range.Paragraphs.Count).Next
        On Error GoTo Fail
        For i = 1 To 3
            If p Is Nothing Then Exit For
            t = Trim$(NormalizeForDocument(p.Range.Text))
            If Len(t) > 0 Then
                keyText = Trim$(NormalizeForDocument(StripTableNumberPrefix(t)))
                If Len(keyText) > 0 Then
                    If Not seen.Exists(keyText) Then
                        seen.Add keyText, True
                        Call IncrementDictCount(d, keyText)
                    End If
                End If
            End If
            On Error Resume Next
            Set p = p.Next
            On Error GoTo Fail
        Next i
    Next tbl
    Set BuildCaptionKeyCounts = d
    Exit Function
Fail:
    Set BuildCaptionKeyCounts = NewDictionary()
End Function

Private Sub IncrementDictCount(ByVal d As Object, ByVal keyText As String)
    On Error GoTo CleanExit
    If d Is Nothing Then Exit Sub
    Dim k As String
    k = Trim$(NormalizeForDocument(keyText))
    If Len(k) = 0 Then Exit Sub
    If d.Exists(k) Then
        d(k) = CLng(d(k)) + 1
    Else
        d.Add k, 1
    End If
CleanExit:
End Sub

Private Function FindTextInRange(ByVal text As String, ByVal searchRange As Range) As Range
    Dim r As Range
    Set r = searchRange.Duplicate
    With r.Find
        .ClearFormatting
        .Text = text
        .MatchWholeWord = True
        .Wrap = wdFindStop
        If .Execute Then
            Set FindTextInRange = r
        Else
            Set FindTextInRange = Nothing
        End If
    End With
End Function

Private Function PositionsRoughlyAlign(ByVal rangeA As Range, ByVal rangeB As Range) As Boolean
    On Error Resume Next
    Dim posA As Single
    Dim posB As Single
    posA = rangeA.Information(wdHorizontalPositionRelativeToPage)
    posB = rangeB.Information(wdHorizontalPositionRelativeToPage)

    If Err.Number <> 0 Then
        PositionsRoughlyAlign = False
    Else
        PositionsRoughlyAlign = (Abs(posA - posB) <= 5)
    End If
    Err.Clear
End Function

Private Function FindPipeDelimitedContext(ByVal searchString As String, ByVal searchRange As Range, ByVal matchCase As Boolean) As Range
    ' Strategy 0.5: Handles pipe-delimited strings (common in table representations)
    ' Splits by | and newlines, then searches for the sequence of segments.
    
    If InStr(searchString, "|") = 0 Then
        Set FindPipeDelimitedContext = Nothing
        Exit Function
    End If
    
    Debug.Print "    [PipeSearch] Input(len=" & Len(searchString) & ") '" & Left$(Replace(searchString, Chr(13), "\\r"), 120) & IIf(Len(searchString) > 120, "...", "") & "'"
    
    ' Prepare the list of segments
    Dim rawSegments() As String
    ' Replace pipe with a unique delimiter for splitting
    Dim tempStr As String
    tempStr = Replace(searchString, "|", Chr(1))
    ' Also handle newlines as separators for table rows
    tempStr = Replace(tempStr, vbCrLf, Chr(1))
    tempStr = Replace(tempStr, vbCr, Chr(1))
    tempStr = Replace(tempStr, vbLf, Chr(1))
    
    rawSegments = Split(tempStr, Chr(1))
    
    Dim segments As New Collection
    Dim i As Long
    Dim seg As String
    
    For i = LBound(rawSegments) To UBound(rawSegments)
        seg = Trim(rawSegments(i))
        If Len(seg) > 0 Then segments.Add seg
    Next i
    
    Debug.Print "    [PipeSearch] Segments=" & segments.Count
    For i = 1 To segments.Count
        Debug.Print "    [PipeSearch]  seg(" & i & ")='" & Left$(segments(i), 80) & IIf(Len(segments(i)) > 80, "...", "") & "'"
    Next i
    
    If segments.Count = 0 Then
        Set FindPipeDelimitedContext = Nothing
        Exit Function
    End If
    
    ' Begin search for the sequence
    Dim firstSeg As String
    firstSeg = segments(1)
    Dim firstSegNorm As String
    firstSegNorm = NormalizeForDocument(firstSeg)
    
    Dim searchCursor As Range
    Set searchCursor = searchRange.Duplicate
    
    Dim matchStart As Range
    Dim currentPos As Range
    Dim seqIdx As Long
    Dim allFound As Boolean
    Dim partRange As Range
    Dim loopCounter As Long
    Dim lastMatchPos As Long
    Dim candidateContainer As Range
    Dim advanceStart As Long
    
    loopCounter = 0
    lastMatchPos = -1
    
    Do
        ' Safety check: prevent infinite loops
        loopCounter = loopCounter + 1
        If loopCounter > 50 Then
            Debug.Print "    -> SAFETY EXIT: Loop counter exceeded in FindPipeDelimitedContext"
            Exit Do
        End If
        
        ' Find candidate start
        Debug.Print "    [PipeSearch] Iter=" & loopCounter & " cursor=" & searchCursor.Start & "-" & searchCursor.End
        Set matchStart = FindLongString(firstSegNorm, searchCursor, matchCase)
        If matchStart Is Nothing Then Exit Do
        Debug.Print "    [PipeSearch]  firstMatch=" & matchStart.Start & "-" & matchStart.End & " text='" & Replace(Left$(NormalizeForDocument(matchStart.Text), 80), Chr(13), "\\r") & IIf(Len(matchStart.Text) > 80, "...", "") & "'"

        ' Check if this match is within TOC - skip if so
        If IsRangeInTOC(matchStart) Then
            Debug.Print "    [PipeSearch]  Skipping TOC match, advancing cursor"
            advanceStart = matchStart.End
            If advanceStart <= searchCursor.Start Then advanceStart = searchCursor.Start + 1
            searchCursor.Start = advanceStart
            If searchCursor.Start >= searchCursor.End Then Exit Do
            GoTo ContinueDo
        End If

        If matchStart.Start < searchCursor.Start Then
            Debug.Print "    [PipeSearch]  WARNING: matchStart before cursor. Forcing cursor advance."
            searchCursor.Start = searchCursor.Start + 1
            If searchCursor.Start >= searchCursor.End Then Exit Do
            GoTo ContinueDo
        End If

        Set candidateContainer = Nothing
        If matchStart.Information(wdWithinTable) Then
            On Error Resume Next
            Set candidateContainer = matchStart.Cells(1).Row.Range
            If Not candidateContainer Is Nothing Then
                If candidateContainer.End > candidateContainer.Start Then candidateContainer.End = candidateContainer.End - 1
            End If
            On Error GoTo 0
        End If
        If candidateContainer Is Nothing Then
            Set candidateContainer = searchRange.Duplicate
        End If
        
        ' Safety check: ensure we're making progress
        If matchStart.Start = lastMatchPos Then
            Debug.Print "    [PipeSearch]  WARNING: Same position matched twice. Advancing cursor and continuing."
            advanceStart = matchStart.End
            If matchStart.Information(wdWithinTable) Then
                If Not candidateContainer Is Nothing Then advanceStart = candidateContainer.End
            End If
            If advanceStart <= searchCursor.Start Then advanceStart = searchCursor.Start + 1
            searchCursor.Start = advanceStart
            If searchCursor.Start >= searchCursor.End Then Exit Do
            GoTo ContinueDo
        End If
        lastMatchPos = matchStart.Start
        
        ' Candidate found, verify subsequent segments
        Set currentPos = matchStart.Duplicate
        allFound = True
        
        If segments.Count > 1 Then
            For seqIdx = 2 To segments.Count
                Dim nextSeg As String
                nextSeg = NormalizeForDocument(segments(seqIdx))
                Debug.Print "    [PipeSearch]   seek seg(" & seqIdx & ")='" & Left$(nextSeg, 80) & IIf(Len(nextSeg) > 80, "...", "") & "'"
                
                ' Look ahead from currentPos.End
                Dim lookAhead As Range
                Set lookAhead = candidateContainer.Duplicate
                lookAhead.Start = currentPos.End
                
                ' Safety check for range validity
                If lookAhead.Start >= lookAhead.End Then
                    allFound = False
                    Exit For
                End If
                
                ' Limit lookahead distance (e.g. 500 chars) to prevent false matches across the doc
                If lookAhead.End - lookAhead.Start > 500 Then
                    lookAhead.End = lookAhead.Start + 500
                End If
                Debug.Print "    [PipeSearch]    lookAhead=" & lookAhead.Start & "-" & lookAhead.End & " preview='" & Replace(Left$(NormalizeForDocument(lookAhead.Text), 120), Chr(13), "\\r") & IIf(Len(lookAhead.Text) > 120, "...", "") & "'"

                If IsSkippablePipeSegment(nextSeg) Then
                    Debug.Print "    [PipeSearch]    SKIP seg(" & seqIdx & ") (skippable)"
                Else
                    Set partRange = FindSegmentWithFallback(nextSeg, lookAhead, matchCase)
                End If
                
                If Not IsSkippablePipeSegment(nextSeg) Then
                    If partRange Is Nothing Then
                    Debug.Print "    [PipeSearch]    FAIL seg(" & seqIdx & ") not found"
                    allFound = False
                    Exit For
                    Else
                    ' Found the next part. Update currentPos to this part.
                    Set currentPos = partRange
                    Debug.Print "    [PipeSearch]    OK seg(" & seqIdx & ") match=" & partRange.Start & "-" & partRange.End
                    End If
                End If
            Next seqIdx
        End If
        
        If allFound Then
            ' Return the full range from start of first segment to end of last segment
            Dim resultRange As Range
            Set resultRange = matchStart.Duplicate
            resultRange.End = currentPos.End
            
            Debug.Print "    -> SUCCESS: Found pipe-delimited sequence."
            Set FindPipeDelimitedContext = resultRange
            Exit Function
        End If
        
        ' If sequence mismatch, advance searchCursor past the first match and try again
        ' CRITICAL: Ensure we always advance at least 1 character
        advanceStart = matchStart.End
        If matchStart.Information(wdWithinTable) Then
            If Not candidateContainer Is Nothing Then advanceStart = candidateContainer.End
        End If
        If advanceStart <= searchCursor.Start Then advanceStart = searchCursor.Start + 1
        searchCursor.Start = advanceStart
        
        If searchCursor.Start >= searchCursor.End Then Exit Do
ContinueDo:
    Loop
    
    Set FindPipeDelimitedContext = Nothing
End Function

' =========================================================================================

Private Function FindWithProgressiveFallback(ByVal searchString As String, ByVal searchRange As Range, _
    Optional ByVal matchCase As Boolean = False, Optional ByVal suggestion As Object = Nothing) As Range
    ' HIGH-IMPACT FEATURE: Tries multiple strategies to find text, reducing "not found" errors
    ' 0. Table cell structure (if provided in suggestion)
    ' 1. Exact match (normalized)
    ' 2. Progressive shortening (80%, 60%, 40% of context from start)
    ' 3. Case-insensitive if case-sensitive was requested
    ' 4. Anchor word search as last resort

    On Error GoTo ErrorHandler

    Dim result As Range
    Dim shortenedContext As String
    Dim searchLen As Long
    Dim cutoffPercent As Variant
    Dim cutoffPercentages As Variant

    ' Strategy 0: Check for tableCell structure (NEW!)
    If Not suggestion Is Nothing Then
        If HasDictionaryKey(suggestion, "tableCell") Then
            Debug.Print "  - Strategy 0: Using tableCell structure..."
            Set result = FindTableCell(suggestion, searchRange)
            If Not result Is Nothing Then
                Debug.Print "    -> SUCCESS: Found via tableCell structure"
                Set FindWithProgressiveFallback = result
                Exit Function
            End If
            Debug.Print "    -> tableCell search failed, falling back to text search"
        End If
    End If

    ' Strategy 0.5: Check for pipe-delimited table row (NEW!)
    If InStr(searchString, "|") > 0 Then
        Debug.Print "  - Strategy 0.5: Checking for pipe-delimited table sequence..."
        Set result = FindPipeDelimitedContext(searchString, searchRange, matchCase)
        If Not result Is Nothing Then
            Debug.Print "    -> SUCCESS: Found pipe-delimited sequence"
            Set FindWithProgressiveFallback = result
            Exit Function
        End If
    End If

    ' Strategy 1: Try exact match with normalization
    Debug.Print "  - Strategy 1: Exact match (normalized, " & IIf(matchCase, "case-sensitive", "case-insensitive") & ")"
    Set result = FindLongString(searchString, searchRange, matchCase)
    If Not result Is Nothing Then
        Debug.Print "    -> SUCCESS: Found with exact match"
        Set FindWithProgressiveFallback = result
        Exit Function
    End If
    
    ' Strategy 2: Progressive context shortening (for overly specific contexts)
    ' Try 90%, 75% of the original context length from the START
    ' REDUCED from 80%, 60%, 40%, 25% to minimize false positives
    If Len(searchString) > 50 Then ' Only for longer contexts
        Debug.Print "  - Strategy 2: Progressive shortening..."
        cutoffPercentages = Array(0.9, 0.75)

        For Each cutoffPercent In cutoffPercentages
            searchLen = CLng(Len(searchString) * cutoffPercent)
            If searchLen < 40 Then Exit For ' Require at least 40 chars for reliability
            
            shortenedContext = Left$(searchString, searchLen)
            ' Trim to word boundary for cleaner matching
            shortenedContext = TrimToWordBoundary(shortenedContext)
            
            Debug.Print "    -> Trying " & (cutoffPercent * 100) & "% (" & Len(shortenedContext) & " chars)..."
            Set result = FindLongString(shortenedContext, searchRange, matchCase)

            If Not result Is Nothing Then
                ' Verify match quality before accepting
                If VerifyMatchQuality(searchString, result, 2) Then
                    Debug.Print "    -> SUCCESS: Found with " & (cutoffPercent * 100) & "% context (verified)"
                    Set FindWithProgressiveFallback = result
                    Exit Function
                Else
                    Debug.Print "    -> Match found but REJECTED due to poor quality"
                    Set result = Nothing ' Continue searching
                End If
            End If
        Next cutoffPercent
    End If
    
    ' Strategy 3: Case-insensitive fallback (if original was case-sensitive)
    If matchCase Then
        Debug.Print "  - Strategy 3: Case-insensitive fallback..."
        Set result = FindLongString(searchString, searchRange, False)
        If Not result Is Nothing Then
            ' Verify match quality before accepting
            If VerifyMatchQuality(searchString, result, 3) Then
                Debug.Print "    -> SUCCESS: Found with case-insensitive match (verified)"
                Set FindWithProgressiveFallback = result
                Exit Function
            Else
                Debug.Print "    -> Match found but REJECTED due to poor quality"
                Set result = Nothing
            End If
        End If

        ' Try shortened contexts case-insensitive too
        If Len(searchString) > 50 Then
            cutoffPercentages = Array(0.9)
            For Each cutoffPercent In cutoffPercentages
                searchLen = CLng(Len(searchString) * cutoffPercent)
                If searchLen < 40 Then Exit For

                shortenedContext = TrimToWordBoundary(Left$(searchString, searchLen))
                Set result = FindLongString(shortenedContext, searchRange, False)

                If Not result Is Nothing Then
                    ' Verify match quality (using strategy 2 since it's a shortened context)
                    If VerifyMatchQuality(searchString, result, 2) Then
                        Debug.Print "    -> SUCCESS: Found with " & (cutoffPercent * 100) & "% context (case-insensitive, verified)"
                        Set FindWithProgressiveFallback = result
                        Exit Function
                    Else
                        Debug.Print "    -> Match found but REJECTED due to poor quality"
                        Set result = Nothing
                    End If
                End If
            Next cutoffPercent
        End If
    End If
    
    ' Strategy 4: Anchor word as last resort (DISABLED - high false positive risk)
    ' This strategy is disabled because matching a single word has extremely high
    ' false positive rates. It will almost always match the wrong location.
    ' If you need this fallback, the HandleNotFoundContext function will place
    ' a comment at a keyword location instead.
    Debug.Print "  - Strategy 4: Anchor word fallback SKIPPED (disabled to prevent false positives)"

    ' UNCOMMENT BELOW TO RE-ENABLE (not recommended):
    ' Dim anchorWord As String
    ' anchorWord = ExtractAnchorWord(searchString)
    ' If Len(anchorWord) >= 5 Then
    '     Debug.Print "    -> Searching for anchor word: '" & anchorWord & "'"
    '     Set result = FindLongString(anchorWord, searchRange, False)
    '     If Not result Is Nothing Then
    '         Debug.Print "    -> SUCCESS: Found anchor word (WARNING: May be imprecise)"
    '         Set FindWithProgressiveFallback = result
    '         Exit Function
    '     End If
    ' End If
    
    ' All strategies failed - provide helpful debugging info
    Debug.Print "  - FAILED: All matching strategies exhausted"
    Debug.Print "  - Searched for: '" & Left$(searchString, 100) & IIf(Len(searchString) > 100, "...", "") & "'"
    Debug.Print "  - TIPS: (1) Try shortening the context to a unique phrase"
    Debug.Print "          (2) Check for typos or punctuation differences"
    Debug.Print "          (3) Ensure text actually exists in the document"
    Debug.Print "          (4) Try using a distinctive word from the passage"
    
    ' Warn about long contexts
    If Len(searchString) > 200 Then
        Debug.Print "  - WARNING: Context is very long (" & Len(searchString) & " chars). Shorter contexts match more reliably."
    End If
    
    Set FindWithProgressiveFallback = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in FindWithProgressiveFallback: " & Err.Description
    Set FindWithProgressiveFallback = Nothing
End Function

Private Function TrimToWordBoundary(ByVal text As String) As String
    ' Trims text to the last complete word to avoid mid-word breaks
    Dim i As Long
    Dim ch As String
    
    ' Trim from the end until we hit whitespace or punctuation
    For i = Len(text) To 1 Step -1
        ch = Mid$(text, i, 1)
        If ch = " " Or ch = Chr(13) Or ch = "." Or ch = "," Or ch = ";" Or ch = ":" Then
            TrimToWordBoundary = Trim$(Left$(text, i))
            Exit Function
        End If
    Next i
    
    ' If no word boundary found, return as-is
    TrimToWordBoundary = Trim$(text)
End Function

Private Function FindLongString(ByVal searchString As String, ByVal searchRange As Range, _
    Optional ByVal matchCase As Boolean = False) As Range

    On Error GoTo ErrorHandler

    Dim findInRange As Range
    Set findInRange = searchRange.Duplicate

    ' --- Standard Find for short strings with enhanced matching ---
    If Len(searchString) <= 255 Then
        With findInRange.Find
            .ClearFormatting
            .Text = searchString
            .matchCase = matchCase
            .Wrap = wdFindStop
            .Forward = True
            .MatchWildcards = False
            .MatchWholeWord = False
            .MatchAllWordForms = False

            ' Loop to skip TOC matches
            Dim shortLoopCounter As Long
            Dim shortLastPos As Long
            shortLoopCounter = 0
            shortLastPos = -1
            Do While .Execute
                shortLoopCounter = shortLoopCounter + 1
                If shortLoopCounter > LOOP_SAFETY_LIMIT Then
                    Debug.Print "    -> SAFETY EXIT: FindLongString (short) exceeded " & LOOP_SAFETY_LIMIT & " iterations"
                    Exit Do
                End If
                ' Prevent infinite loop if Find returns same position
                If findInRange.Start = shortLastPos Then
                    Debug.Print "    -> SAFETY EXIT: FindLongString (short) stuck at position " & shortLastPos
                    Exit Do
                End If
                shortLastPos = findInRange.Start
                ' Check if this match is within a TOC
                If Not IsRangeInTOC(findInRange) Then
                    ' Valid match outside TOC
                    Set FindLongString = findInRange
                    Exit Function
                End If
                ' Continue searching past this TOC match
                findInRange.Collapse wdCollapseEnd
            Loop

            ' No valid match found outside TOC, try fuzzy matching
            Set FindLongString = FuzzyFindString(searchString, searchRange, matchCase)
        End With
        Exit Function
    End If

    ' --- REVISED LOGIC for strings > 255 characters ---
    Dim startPart As String
    Dim endPart As String
    
    startPart = Left$(searchString, 255)
    endPart = Mid$(searchString, 256)

    With findInRange.Find
        .ClearFormatting
        .Text = startPart
        .matchCase = matchCase
        .Wrap = wdFindStop
        .Forward = True
        .MatchWildcards = False
        .MatchWholeWord = False
        .MatchAllWordForms = False

        Dim longLoopCounter As Long
        Dim longLastPos As Long
        longLoopCounter = 0
        longLastPos = -1
        Do While .Execute
            longLoopCounter = longLoopCounter + 1
            If longLoopCounter > LOOP_SAFETY_LIMIT Then
                Debug.Print "    -> SAFETY EXIT: FindLongString (long) exceeded " & LOOP_SAFETY_LIMIT & " iterations"
                Exit Do
            End If
            ' Prevent infinite loop if Find returns same position
            If findInRange.Start = longLastPos Then
                Debug.Print "    -> SAFETY EXIT: FindLongString (long) stuck at position " & longLastPos
                Exit Do
            End If
            longLastPos = findInRange.Start
            ' Check if this match is within a TOC - skip if so
            If IsRangeInTOC(findInRange) Then
                findInRange.Collapse wdCollapseEnd
                GoTo ContinueLoop
            End If

            ' Found the first part. 'findInRange' now represents that part.
            ' Let's check if the text IMMEDIATELY following it matches the endPart.

            Dim checkRange As Range
            Set checkRange = findInRange.Duplicate
            checkRange.Collapse wdCollapseEnd ' Move to the end of the found text

            ' Expand the checkRange to be the same length as the string we're looking for
            checkRange.End = checkRange.Start + Len(endPart)

            ' SAFEGUARD: Ensure the checkRange does not exceed the document's boundaries
            If checkRange.End > searchRange.End Then GoTo ContinueLoop

            ' Normalize both strings for comparison to handle whitespace variations
            Dim normalizedCheck As String
            Dim normalizedEnd As String
            normalizedCheck = NormalizeForDocument(checkRange.Text)
            normalizedEnd = NormalizeForDocument(endPart)

            ' Use a direct string comparison
            If normalizedCheck = normalizedEnd Then
                ' Success! Combine the ranges.
                findInRange.End = checkRange.End
                Set FindLongString = findInRange
                Exit Function
            End If

ContinueLoop:
            ' It wasn't a match, so collapse and continue searching.
            findInRange.Collapse wdCollapseEnd
        Loop
    End With

    ' If exact match failed, try fuzzy matching for long strings
    Set FindLongString = FuzzyFindString(searchString, searchRange, matchCase)
    Exit Function

ErrorHandler:
    Debug.Print "An error occurred in FindLongString: " & Err.Description
    Set FindLongString = Nothing
End Function

Private Function FuzzyFindString(ByVal searchString As String, ByVal searchRange As Range, _
    Optional ByVal matchCase As Boolean = False) As Range
    ' Attempts to find text with flexible whitespace matching
    ' This handles cases where the document has different whitespace than the search string
    
    On Error GoTo ErrorHandler
    
    ' For very long strings, don't attempt fuzzy matching (too slow)
    If Len(searchString) > 500 Then
        Set FuzzyFindString = Nothing
        Exit Function
    End If
    
    ' Extract the first significant word (5+ chars) to use as an anchor
    Dim anchorWord As String
    anchorWord = ExtractAnchorWord(searchString)
    
    If Len(anchorWord) < 3 Then
        Set FuzzyFindString = Nothing
        Exit Function
    End If
    
    ' Search for the anchor word
    Dim findInRange As Range
    Set findInRange = searchRange.Duplicate
    
    With findInRange.Find
        .ClearFormatting
        .Text = anchorWord
        .matchCase = matchCase
        .Wrap = wdFindStop
        .Forward = True
        .MatchWildcards = False

        Dim fuzzyLoopCounter As Long
        Dim fuzzyLastPos As Long
        fuzzyLoopCounter = 0
        fuzzyLastPos = -1
        Do While .Execute
            fuzzyLoopCounter = fuzzyLoopCounter + 1
            If fuzzyLoopCounter > LOOP_SAFETY_LIMIT Then
                Debug.Print "    -> SAFETY EXIT: FuzzyFindString exceeded " & LOOP_SAFETY_LIMIT & " iterations"
                Exit Do
            End If
            ' Prevent infinite loop if Find returns same position
            If findInRange.Start = fuzzyLastPos Then
                Debug.Print "    -> SAFETY EXIT: FuzzyFindString stuck at position " & fuzzyLastPos
                Exit Do
            End If
            fuzzyLastPos = findInRange.Start
            ' Skip if this anchor match is within TOC
            If IsRangeInTOC(findInRange) Then
                findInRange.Collapse wdCollapseEnd
                GoTo ContinueFuzzySearch
            End If

            ' Found the anchor word. Now check if the surrounding text matches
            Dim testRange As Range
            Set testRange = findInRange.Duplicate

            ' Expand the range to approximately the length of the search string
            testRange.Start = testRange.Start - Len(searchString) / 4
            testRange.End = testRange.End + Len(searchString)

            ' Ensure we don't go out of bounds
            If testRange.Start < searchRange.Start Then testRange.Start = searchRange.Start
            If testRange.End > searchRange.End Then testRange.End = searchRange.End

            ' Normalize both strings and compare
            Dim normalizedTest As String
            Dim normalizedSearch As String
            normalizedTest = NormalizeForDocument(testRange.Text)
            normalizedSearch = NormalizeForDocument(searchString)

            ' Check if the search string is contained in the test range
            If InStr(1, normalizedTest, normalizedSearch, IIf(matchCase, vbBinaryCompare, vbTextCompare)) > 0 Then
                ' Find the exact position and length
                Dim startPos As Long
                startPos = InStr(1, normalizedTest, normalizedSearch, IIf(matchCase, vbBinaryCompare, vbTextCompare))

                ' Adjust the range to match exactly
                testRange.Start = testRange.Start + startPos - 1
                testRange.End = testRange.Start + Len(searchString)

                ' Double-check that the final result isn't in TOC
                If Not IsRangeInTOC(testRange) Then
                    Set FuzzyFindString = testRange
                    Debug.Print "  - FUZZY MATCH: Found using anchor word '" & anchorWord & "'"
                    Exit Function
                End If
            End If

ContinueFuzzySearch:
            ' Continue searching
            findInRange.Collapse wdCollapseEnd
        Loop
    End With
    
    Set FuzzyFindString = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "An error occurred in FuzzyFindString: " & Err.Description
    Set FuzzyFindString = Nothing
End Function

Private Function CalculateWordOverlap(ByVal searchText As String, ByVal foundText As String) As Double
    ' Calculates the proportion of significant words from searchText that appear in foundText
    ' Returns a value from 0.0 (no overlap) to 1.0 (perfect overlap)
    ' Used to verify match quality and reduce false positives

    On Error GoTo ErrorHandler

    Dim searchWords() As String
    Dim foundWords() As String
    Dim searchWord As Variant
    Dim matchCount As Long
    Dim totalSignificantWords As Long
    Dim cleanSearch As String
    Dim cleanFound As String
    Dim commonWords As String

    ' List of common words to ignore in overlap calculation
    commonWords = "|the|and|for|with|from|that|this|have|been|were|will|would|could|should|are|was|has|had|"

    ' Clean and normalize both texts
    cleanSearch = NormalizeForDocument(searchText)
    cleanFound = NormalizeForDocument(foundText)

    ' Remove punctuation
    cleanSearch = Replace(cleanSearch, ",", " ")
    cleanSearch = Replace(cleanSearch, ".", " ")
    cleanSearch = Replace(cleanSearch, ";", " ")
    cleanSearch = Replace(cleanSearch, ":", " ")
    cleanSearch = Replace(cleanSearch, "(", " ")
    cleanSearch = Replace(cleanSearch, ")", " ")
    cleanSearch = Replace(cleanSearch, """", " ")
    cleanSearch = Replace(cleanSearch, "'", " ")

    cleanFound = Replace(cleanFound, ",", " ")
    cleanFound = Replace(cleanFound, ".", " ")
    cleanFound = Replace(cleanFound, ";", " ")
    cleanFound = Replace(cleanFound, ":", " ")
    cleanFound = Replace(cleanFound, "(", " ")
    cleanFound = Replace(cleanFound, ")", " ")
    cleanFound = Replace(cleanFound, """", " ")
    cleanFound = Replace(cleanFound, "'", " ")

    ' Split into words
    searchWords = Split(Trim(cleanSearch), " ")

    ' Count matches for significant words only
    matchCount = 0
    totalSignificantWords = 0

    For Each searchWord In searchWords
        ' Skip empty words and common words
        If Len(searchWord) > 0 Then
            ' Only count words > 3 chars and not common words
            If Len(searchWord) > 3 And InStr(1, commonWords, "|" & LCase(searchWord) & "|", vbTextCompare) = 0 Then
                totalSignificantWords = totalSignificantWords + 1

                ' Check if this word appears in the found text
                If InStr(1, " " & cleanFound & " ", " " & searchWord & " ", vbTextCompare) > 0 Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next searchWord

    ' Calculate overlap ratio
    If totalSignificantWords > 0 Then
        CalculateWordOverlap = CDbl(matchCount) / CDbl(totalSignificantWords)
    Else
        CalculateWordOverlap = 0#
    End If

    Exit Function

ErrorHandler:
    Debug.Print "Error in CalculateWordOverlap: " & Err.Description
    CalculateWordOverlap = 0#
End Function

Private Function VerifyMatchQuality(ByVal searchString As String, _
                                    ByVal foundRange As Range, _
                                    ByVal strategyUsed As Long) As Boolean
    ' Verifies that a found match is of acceptable quality
    ' Returns True if the match should be accepted, False if it's likely a false positive
    ' strategyUsed: 1=exact, 2=shortened, 3=case-insensitive, 4=anchor (disabled)

    On Error GoTo ErrorHandler

    Dim overlap As Double
    Dim foundText As String

    Select Case strategyUsed
        Case 1  ' Exact match - always accept
            VerifyMatchQuality = True
            Debug.Print "    -> Match quality: EXACT (100% - always accepted)"

        Case 2  ' Shortened context - require good word overlap
            foundText = foundRange.Text
            overlap = CalculateWordOverlap(searchString, foundText)

            ' Require at least 70% of significant words to match
            If overlap >= 0.7 Then
                VerifyMatchQuality = True
                Debug.Print "    -> Match quality: GOOD (" & Format(overlap * 100, "0") & "% word overlap)"
            Else
                VerifyMatchQuality = False
                Debug.Print "    -> Match quality: POOR (" & Format(overlap * 100, "0") & "% word overlap - REJECTED)"
            End If

        Case 3  ' Case-insensitive - accept if it's truly the same text
            foundText = foundRange.Text
            overlap = CalculateWordOverlap(searchString, foundText)

            ' For case-insensitive, be more lenient since the text should still match
            If overlap >= 0.6 Then
                VerifyMatchQuality = True
                Debug.Print "    -> Match quality: ACCEPTABLE (" & Format(overlap * 100, "0") & "% word overlap)"
            Else
                VerifyMatchQuality = False
                Debug.Print "    -> Match quality: POOR (" & Format(overlap * 100, "0") & "% word overlap - REJECTED)"
            End If

        Case 4  ' Anchor word - this strategy is disabled, but if re-enabled, require very high confidence
            foundText = foundRange.Text
            overlap = CalculateWordOverlap(searchString, foundText)

            ' Require 80% overlap for anchor word matches
            If overlap >= 0.8 Then
                VerifyMatchQuality = True
                Debug.Print "    -> Match quality: HIGH (" & Format(overlap * 100, "0") & "% word overlap)"
            Else
                VerifyMatchQuality = False
                Debug.Print "    -> Match quality: INSUFFICIENT (" & Format(overlap * 100, "0") & "% word overlap - REJECTED)"
            End If

        Case Else
            VerifyMatchQuality = False
            Debug.Print "    -> Match quality: UNKNOWN strategy - REJECTED"
    End Select

    Exit Function

ErrorHandler:
    Debug.Print "Error in VerifyMatchQuality: " & Err.Description
    VerifyMatchQuality = False
End Function

Private Function ExtractAnchorWord(ByVal text As String) As String
    ' Extracts a significant word from the text to use as a search anchor
    ' Prefers words that are 5+ characters and not common words
    
    Dim words() As String
    Dim word As Variant
    Dim cleanText As String
    Dim commonWords As String
    
    ' List of common words to avoid
    commonWords = "|the|and|for|with|from|that|this|have|been|were|will|would|could|should|"
    
    ' Clean the text
    cleanText = text
    cleanText = Replace(cleanText, ",", " ")
    cleanText = Replace(cleanText, ".", " ")
    cleanText = Replace(cleanText, ";", " ")
    cleanText = Replace(cleanText, ":", " ")
    cleanText = Replace(cleanText, Chr(13), " ")
    cleanText = Replace(cleanText, vbTab, " ")
    
    words = Split(Trim(cleanText), " ")
    
    ' Find first word that is 5+ chars and not a common word
    For Each word In words
        If Len(word) >= 5 Then
            If InStr(1, commonWords, "|" & LCase(word) & "|", vbTextCompare) = 0 Then
                ExtractAnchorWord = CStr(word)
                Exit Function
            End If
        End If
    Next word
    
    ' If no good word found, return the first word > 3 chars
    For Each word In words
        If Len(word) > 3 Then
            ExtractAnchorWord = CStr(word)
            Exit Function
        End If
    Next word
    
    ' Last resort: return the first non-empty word
    For Each word In words
        If Len(word) > 0 Then
            ExtractAnchorWord = CStr(word)
            Exit Function
        End If
    Next word
    
    ExtractAnchorWord = ""
End Function

Private Function pv_ParseNumber() As Double
    Dim StartIndex As Long
    StartIndex = p_Index
    Do While p_Index <= Len(p_Json)
        Select Case Mid$(p_Json, p_Index, 1)
            Case "0" To "9", ".", "e", "E", "+", "-"
                p_Index = p_Index + 1
            Case Else
                Exit Do
        End Select

    Loop
    pv_ParseNumber = CDbl(Mid$(p_Json, StartIndex, p_Index - StartIndex))
End Function

' =========================================================================================

' === HELPER FUNCTIONS FOR INLINE FORMATTING TAGS ============================================

' =========================================================================================

Private Function IsFormattingAlreadyApplied(ByVal targetRange As Range, ByVal replaceText As String) As Boolean
    ' Checks if the current formatting in the range already matches what would be applied
    ' Returns True if formatting is already correct (skip the change)
    
    On Error GoTo ErrorHandler
    IsFormattingAlreadyApplied = False
    
    ' Parse the replacement text to get plain text and formatting segments
    Dim plainText As String
    Dim segments As Collection
    ParseFormattingTags replaceText, plainText, segments
    
    ' First check: Does the plain text match (normalized, case-insensitive)?
    If NormalizeForDocument(targetRange.Text) <> NormalizeForDocument(plainText) Then
        ' Text content is different, so we need to apply the change
        Debug.Print "    - Text differs: '" & targetRange.Text & "' vs '" & plainText & "'"
        IsFormattingAlreadyApplied = False
        Exit Function
    End If
    
    ' Text matches! Now check if formatting also matches
    If segments Is Nothing Or segments.Count = 0 Then
        ' No formatting requested, and text matches - already correct!
        Debug.Print "    - Text matches and no formatting needed"
        IsFormattingAlreadyApplied = True
        Exit Function
    End If
    
    ' Check each formatting segment
    Dim segment As Object
    Dim baseStart As Long
    baseStart = targetRange.Start
    
    For Each segment In segments
        Dim checkRange As Range
        Set checkRange = targetRange.Duplicate
        checkRange.Start = baseStart + CLng(segment("Start")) - 1
        checkRange.End = checkRange.Start + CLng(segment("Length"))
        
        ' Check if the formatting matches
        With checkRange.Font
            If CBool(segment("Bold")) And .Bold <> True Then
                Debug.Print "    - Bold formatting missing at position " & segment("Start")
                IsFormattingAlreadyApplied = False
                Exit Function
            End If
            
            If CBool(segment("Italic")) And .Italic <> True Then
                Debug.Print "    - Italic formatting missing at position " & segment("Start")
                IsFormattingAlreadyApplied = False
                Exit Function
            End If
            
            If CBool(segment("Subscript")) And .Subscript <> True Then
                Debug.Print "    - Subscript formatting missing at position " & segment("Start")
                IsFormattingAlreadyApplied = False
                Exit Function
            End If
            
            If CBool(segment("Superscript")) And .Superscript <> True Then
                Debug.Print "    - Superscript formatting missing at position " & segment("Start")
                IsFormattingAlreadyApplied = False
                Exit Function
            End If
        End With
    Next segment
    
    ' All checks passed - formatting is already applied!
    Debug.Print "    - Formatting already matches perfectly"
    IsFormattingAlreadyApplied = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in IsFormattingAlreadyApplied: " & Err.Description
    IsFormattingAlreadyApplied = False ' On error, assume we need to apply
End Function

Private Sub ApplyFormattedReplacement(ByVal targetRange As Range, ByVal replaceText As String)
    Debug.Print "    - Applying formatted replacement with granular diff support..."

    ' Step 1: Parse formatting tags from the replacement text
    Dim plainText As String
    Dim segments As Collection
    ParseFormattingTags replaceText, plainText, segments

    ' Step 2: Get current text in the target range
    Dim currentText As String
    currentText = targetRange.Text

    If segments Is Nothing Or segments.Count = 0 Then
        Dim normalizedNewText As String
        normalizedNewText = NormalizeReplacementBoundaries(targetRange, currentText, plainText)
        If normalizedNewText <> plainText Then
            plainText = normalizedNewText
        End If
    End If

    Debug.Print "    - Current text length: " & Len(currentText) & ", New text length: " & Len(plainText)

    ' Step 3: Check if only formatting changed (text is identical)
    If currentText = plainText Then
        Debug.Print "    - Text is identical, applying formatting only..."
        ApplyFormattingOnly targetRange, segments
        Exit Sub
    End If

    ' Step 4: Decide whether to use granular diff or fall back to wholesale replacement
    Dim useGranular As Boolean
    useGranular = ShouldUseGranularDiff(currentText, plainText, segments)

    If Not useGranular Then
        ' FALLBACK: Use the original wholesale replacement method
        Debug.Print "    - Using FALLBACK (wholesale replacement)..."
        targetRange.Text = plainText
        targetRange.Font.Reset
        ApplyFormattingToSegments targetRange, segments
        Debug.Print "    -> Formatting tags applied (fallback method)."
        Exit Sub
    End If

    ' GRANULAR: Use character-level diff for tracked changes
    Debug.Print "    - Using GRANULAR DIFF (character-level changes)..."

    ' Step 5: Compute the diff operations
    Dim diffOps As Collection
    Set diffOps = ComputeDiff(currentText, plainText)

    ' Step 6: Apply the diff operations with Track Changes enabled
    ApplyDiffOperations targetRange, diffOps, segments

    Debug.Print "    -> Granular diff complete."
End Sub

Private Function NormalizeReplacementBoundaries(ByVal targetRange As Range, ByVal oldText As String, ByVal newText As String) As String
    On Error GoTo CleanFail

    Dim result As String
    result = newText

    Dim prevChar As String
    prevChar = ""
    If targetRange.Start > 0 Then
        prevChar = ActiveDocument.Range(targetRange.Start - 1, targetRange.Start).Text
    End If

    Dim nextChar As String
    nextChar = ""
    If targetRange.End < ActiveDocument.Content.End Then
        nextChar = ActiveDocument.Range(targetRange.End, targetRange.End + 1).Text
    End If

    Dim leadLen As Long
    leadLen = 0
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(oldText)
        ch = Mid$(oldText, i, 1)
        If ch = " " Or ch = vbTab Then
            leadLen = leadLen + 1
        Else
            Exit For
        End If
    Next i

    If leadLen > 0 And Len(result) > 0 Then
        Dim firstNew As String
        firstNew = Left$(result, 1)
        If firstNew <> " " And firstNew <> vbTab Then
            If targetRange.Start = 0 Or prevChar = Chr$(13) Or prevChar = Chr$(12) Then
                result = Left$(oldText, leadLen) & result
            End If
        End If
    End If

    If Len(result) > 0 And prevChar = "." And Left$(result, 1) = "." Then
        result = Mid$(result, 2)
    End If

    If Len(result) > 0 And nextChar = "." And Right$(result, 1) = "." Then
        result = Left$(result, Len(result) - 1)
    End If

    NormalizeReplacementBoundaries = result
    Exit Function

CleanFail:
    NormalizeReplacementBoundaries = newText
End Function

' =========================================================================================
' === GRANULAR DIFF FUNCTIONS ============================================================
' =========================================================================================



' Computes the length of the common prefix between two strings
Private Function ComputeCommonPrefix(ByVal oldText As String, ByVal newText As String) As Long
    Dim minLen As Long
    Dim i As Long

    minLen = IIf(Len(oldText) < Len(newText), Len(oldText), Len(newText))

    For i = 1 To minLen
        If Mid$(oldText, i, 1) <> Mid$(newText, i, 1) Then
            ComputeCommonPrefix = i - 1
            Exit Function
        End If
    Next i

    ComputeCommonPrefix = minLen
End Function

' Computes the length of the common suffix between two strings
Private Function ComputeCommonSuffix(ByVal oldText As String, ByVal newText As String) As Long
    Dim minLen As Long
    Dim i As Long
    Dim oldLen As Long
    Dim newLen As Long

    oldLen = Len(oldText)
    newLen = Len(newText)
    minLen = IIf(oldLen < newLen, oldLen, newLen)

    For i = 0 To minLen - 1
        If Mid$(oldText, oldLen - i, 1) <> Mid$(newText, newLen - i, 1) Then
            ComputeCommonSuffix = i
            Exit Function
        End If
    Next i

    ComputeCommonSuffix = minLen
End Function

' Computes a collection of diff operations between oldText and newText
' Returns a Collection of Dictionary objects, each with:
'   - "Operation": DIFF_EQUAL, DIFF_INSERT, or DIFF_DELETE
'   - "Text": The text segment for this operation
Private Function ComputeDiff(ByVal oldText As String, ByVal newText As String) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim equalOp As Object
    Dim deleteOp As Object
    Dim insertOp As Object

    TraceLog "    [ComputeDiff] oldText length: " & Len(oldText) & ", newText length: " & Len(newText)

    ' Handle edge cases
    If oldText = newText Then
        ' Identical - single EQUAL operation
        Set equalOp = NewDictionary()
        equalOp("Operation") = DIFF_EQUAL
        equalOp("Text") = oldText
        result.Add equalOp
        TraceLog "    [ComputeDiff] Texts are identical."
        Set ComputeDiff = result
        Exit Function
    End If

    If Len(oldText) = 0 Then
        ' Only insertion
        Set insertOp = NewDictionary()
        insertOp("Operation") = DIFF_INSERT
        insertOp("Text") = newText
        result.Add insertOp
        TraceLog "    [ComputeDiff] Old text empty, full insert."
        Set ComputeDiff = result
        Exit Function
    End If

    If Len(newText) = 0 Then
        ' Only deletion
        Set deleteOp = NewDictionary()
        deleteOp("Operation") = DIFF_DELETE
        deleteOp("Text") = oldText
        result.Add deleteOp
        TraceLog "    [ComputeDiff] New text empty, full delete."
        Set ComputeDiff = result
        Exit Function
    End If

    ' Find common prefix
    Dim prefixLen As Long
    prefixLen = ComputeCommonPrefix(oldText, newText)

    If prefixLen > 0 Then
        Set equalOp = NewDictionary()
        equalOp("Operation") = DIFF_EQUAL
        equalOp("Text") = Left$(oldText, prefixLen)
        result.Add equalOp
        TraceLog "    [ComputeDiff] Common prefix length: " & prefixLen
    End If

    ' Work on middle section (after prefix, before suffix)
    Dim oldMiddle As String
    Dim newMiddle As String
    oldMiddle = Mid$(oldText, prefixLen + 1)
    newMiddle = Mid$(newText, prefixLen + 1)

    ' Find common suffix
    Dim suffixLen As Long
    suffixLen = ComputeCommonSuffix(oldMiddle, newMiddle)

    Dim suffixText As String
    If suffixLen > 0 Then
        suffixText = Right$(oldMiddle, suffixLen)
        oldMiddle = Left$(oldMiddle, Len(oldMiddle) - suffixLen)
        newMiddle = Left$(newMiddle, Len(newMiddle) - suffixLen)
        TraceLog "    [ComputeDiff] Common suffix length: " & suffixLen
    End If

    ' Process the differing middle section
    If Len(oldMiddle) > 0 Then
        Set deleteOp = NewDictionary()
        deleteOp("Operation") = DIFF_DELETE
        deleteOp("Text") = oldMiddle
        result.Add deleteOp
        TraceLog "    [ComputeDiff] Delete: '" & oldMiddle & "'"
    End If

    If Len(newMiddle) > 0 Then
        Set insertOp = NewDictionary()
        insertOp("Operation") = DIFF_INSERT
        insertOp("Text") = newMiddle
        result.Add insertOp
        TraceLog "    [ComputeDiff] Insert: '" & newMiddle & "'"
    End If

    ' Add common suffix
    If suffixLen > 0 Then
        Set equalOp = NewDictionary()
        equalOp("Operation") = DIFF_EQUAL
        equalOp("Text") = suffixText
        result.Add equalOp
    End If

    Set ComputeDiff = result
    TraceLog "    [ComputeDiff] Total operations: " & result.Count
End Function

' Determines whether to use granular diff or fall back to wholesale replacement
Private Function ShouldUseGranularDiff(ByVal oldText As String, _
                                       ByVal newText As String, _
                                       ByVal formatSegments As Collection) As Boolean
    ' Check global feature flag first
    If Not USE_GRANULAR_DIFF Then
        TraceLog "    [ShouldUseGranularDiff] Feature disabled globally."
        ShouldUseGranularDiff = False
        Exit Function
    End If

    ' Rule 1: Text too long? Use fallback for performance
    If Len(oldText) > 1000 Or Len(newText) > 1000 Then
        TraceLog "    [ShouldUseGranularDiff] Text too long, using fallback."
        ShouldUseGranularDiff = False
        Exit Function
    End If

    ' Rule 2: No formatting or simple formatting? Use granular
    If formatSegments Is Nothing Or formatSegments.Count = 0 Then
        TraceLog "    [ShouldUseGranularDiff] No formatting, using granular diff."
        ShouldUseGranularDiff = True
        Exit Function
    End If

    ' Rule 3: Simple formatting (<=  3 segments)? Use granular
    If formatSegments.Count <= 3 Then
        TraceLog "    [ShouldUseGranularDiff] Simple formatting (" & formatSegments.Count & " segments), using granular diff."
        ShouldUseGranularDiff = True
        Exit Function
    End If

    ' Rule 4: Complex formatting? Use fallback
    TraceLog "    [ShouldUseGranularDiff] Complex formatting (" & formatSegments.Count & " segments), using fallback."
    ShouldUseGranularDiff = False
End Function

' Applies diff operations to a range with Track Changes enabled
Private Sub ApplyDiffOperations(ByVal targetRange As Range, _
                                ByVal diffOps As Collection, _
                                ByVal formatSegments As Collection)
    TraceLog "    [ApplyDiffOperations] Applying " & diffOps.Count & " diff operations..."

    ' Track position as we modify the document
    Dim currentPos As Long
    currentPos = targetRange.Start

    Dim op As Object
    Dim opIndex As Long
    opIndex = 0

    For Each op In diffOps
        opIndex = opIndex + 1
        Dim opType As String
        opType = op("Operation")
        Dim opText As String
        opText = op("Text")
        Dim opLen As Long
        opLen = Len(opText)

        TraceLog "    [ApplyDiffOperations] Op #" & opIndex & ": " & opType & " (length=" & opLen & ")"

        Select Case opType
            Case DIFF_EQUAL
                ' Skip over unchanged text
                currentPos = currentPos + opLen
                TraceLog "    [ApplyDiffOperations] EQUAL: Skipping " & opLen & " chars, now at " & currentPos

            Case DIFF_DELETE
                ' Delete this text (will show as tracked deletion)
                If opLen > 0 Then
                    Dim delRange As Range
                    Set delRange = ActiveDocument.Range(currentPos, currentPos + opLen)
                    TraceLog "    [ApplyDiffOperations] DELETE: Deleting range " & currentPos & " to " & (currentPos + opLen) & " ('" & Left$(opText, 20) & "')"
                    delRange.Delete
                    ' Note: currentPos stays same after deletion
                End If

            Case DIFF_INSERT
                ' Insert new text (will show as tracked insertion)
                If opLen > 0 Then
                    Dim insRange As Range
                    Set insRange = ActiveDocument.Range(currentPos, currentPos)
                    TraceLog "    [ApplyDiffOperations] INSERT: Inserting at " & currentPos & " ('" & Left$(opText, 20) & "')"
                    insRange.Text = opText
                    currentPos = currentPos + opLen
                End If

            Case Else
                TraceLog "    [ApplyDiffOperations] WARNING: Unknown operation type: " & opType
        End Select
    Next op

    ' Apply formatting after text changes are complete
    TraceLog "    [ApplyDiffOperations] Text changes complete. Final position: " & currentPos

    ' Determine the new range after all modifications
    Dim finalRange As Range
    Set finalRange = ActiveDocument.Range(targetRange.Start, currentPos)

    ' Only reset font and apply formatting when explicit formatting segments exist.
    ' For plain text replacements (no segments), preserve existing document formatting.
    If Not formatSegments Is Nothing Then
        If formatSegments.Count > 0 Then
            finalRange.Font.Reset
            ApplyFormattingToSegments finalRange, formatSegments
        End If
    End If

    TraceLog "    [ApplyDiffOperations] Diff operations complete."
End Sub

' Applies formatting segments to a range (extracted from old ApplyFormattedReplacement)
Private Sub ApplyFormattingToSegments(ByVal targetRange As Range, ByVal segments As Collection)
    If segments Is Nothing Then
        TraceLog "    [ApplyFormattingToSegments] No segments to apply."
        Exit Sub
    End If

    If segments.Count = 0 Then
        TraceLog "    [ApplyFormattingToSegments] Segments collection is empty."
        Exit Sub
    End If

    Dim baseStart As Long
    baseStart = targetRange.Start
    Dim segment As Object

    TraceLog "    [ApplyFormattingToSegments] Applying " & segments.Count & " formatting segments..."

    For Each segment In segments
        Dim formattedRange As Range
        Set formattedRange = targetRange.Duplicate
        formattedRange.Start = baseStart + CLng(segment("Start")) - 1
        formattedRange.End = formattedRange.Start + CLng(segment("Length"))

        With formattedRange.Font
            If segment("Bold") Then .Bold = True
            If segment("Italic") Then .Italic = True
            If segment("Subscript") And segment("Superscript") Then
                .Subscript = False
                .Superscript = True
            Else
                If segment("Subscript") Then .Subscript = True
                If segment("Superscript") Then .Superscript = True
            End If
        End With
    Next segment

    TraceLog "    [ApplyFormattingToSegments] Formatting applied."
End Sub

' Applies formatting when only formatting changes (no text changes)
Private Sub ApplyFormattingOnly(ByVal targetRange As Range, ByVal segments As Collection)
    TraceLog "    [ApplyFormattingOnly] Text identical, applying formatting only..."

    ' Reset formatting first
    targetRange.Font.Reset

    ' Apply formatting segments
    ApplyFormattingToSegments targetRange, segments
End Sub

Private Sub ParseFormattingTags(ByVal source As String, ByRef plainText As String, ByRef segments As Collection)
    Dim i As Long
    Dim ch As String
    Dim closePos As Long
    Dim tagContent As String
    Dim boldLevel As Long
    Dim italicLevel As Long
    Dim subLevel As Long
    Dim supLevel As Long
    plainText = ""
    Set segments = New Collection
    i = 1
    Do While i <= Len(source)
        ch = Mid$(source, i, 1)
        If ch = "<" Then
            closePos = InStr(i, source, ">")
            If closePos = 0 Then
                plainText = plainText & ch
                AddFormattedChar segments, Len(plainText), (boldLevel > 0), (italicLevel > 0), (subLevel > 0), (supLevel > 0)
                i = i + 1
            Else
                tagContent = Mid$(source, i + 1, closePos - i - 1)
                Select Case LCase$(tagContent)
                    Case "b"
                        boldLevel = boldLevel + 1
                        i = closePos + 1
                    Case "/b"
                        If boldLevel > 0 Then boldLevel = boldLevel - 1
                        i = closePos + 1
                    Case "i"
                        italicLevel = italicLevel + 1
                        i = closePos + 1
                    Case "/i"
                        If italicLevel > 0 Then italicLevel = italicLevel - 1
                        i = closePos + 1
                    Case "sub"
                        subLevel = subLevel + 1
                        i = closePos + 1
                    Case "/sub"
                        If subLevel > 0 Then subLevel = subLevel - 1
                        i = closePos + 1
                    Case "sup"
                        supLevel = supLevel + 1
                        i = closePos + 1
                    Case "/sup"
                        If supLevel > 0 Then supLevel = supLevel - 1
                        i = closePos + 1
                    Case Else
                        plainText = plainText & ch
                        AddFormattedChar segments, Len(plainText), (boldLevel > 0), (italicLevel > 0), (subLevel > 0), (supLevel > 0)
                        i = i + 1
                    End Select

            End If

        Else
            plainText = plainText & ch
            AddFormattedChar segments, Len(plainText), (boldLevel > 0), (italicLevel > 0), (subLevel > 0), (supLevel > 0)
            i = i + 1
        End If

    Loop
End Sub

Private Sub AddFormattedChar(ByRef segments As Collection, ByVal charIndex As Long, ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal isSubscript As Boolean, ByVal isSuperscript As Boolean)
    If segments Is Nothing Then Exit Sub
    If Not (isBold Or isItalic Or isSubscript Or isSuperscript) Then Exit Sub
    Dim segment As Object
    If segments.Count > 0 Then
        Set segment = segments(segments.Count)
        If CLng(segment("Start")) + CLng(segment("Length")) = charIndex _
            And CBool(segment("Bold")) = isBold _
            And CBool(segment("Italic")) = isItalic _
            And CBool(segment("Subscript")) = isSubscript _
            And CBool(segment("Superscript")) = isSuperscript Then
            segment("Length") = CLng(segment("Length")) + 1
            Exit Sub
        End If

    End If

    Set segment = NewDictionary()
    segment("Start") = charIndex
    segment("Length") = 1
    segment("Bold") = isBold
    segment("Italic") = isItalic
    segment("Subscript") = isSubscript
    segment("Superscript") = isSuperscript
    segments.Add segment
End Sub

Private Sub UndoLlmReview()
    ' MODIFIED: This sub now rejects changes for a user-specified author,
    ' defaulting to the current user.

    Dim rev As Revision
    Dim com As Comment
    Dim rejectedChangesCount As Long
    Dim deletedCommentsCount As Long
    Dim i As Long
    Dim userResponse As VbMsgBoxResult
    Dim authorToUndo As String

    On Error GoTo ErrorHandler

    ' --- 1. Ask the user which author to undo ---
    authorToUndo = InputBox("Please enter the author name whose changes you want to undo.", _
                            "Specify Author to Undo", Application.UserName)

    If authorToUndo = "" Then
        MsgBox "Operation cancelled. No author name provided.", vbInformation
        Exit Sub
    End If

    ' --- 2. Confirm with the user before making irreversible changes ---
    userResponse = MsgBox("Are you sure you want to reject all tracked changes and delete all comments made by '" & authorToUndo & "'?" & _
                          vbCrLf & vbCrLf & "This action cannot be undone.", _
                          vbQuestion + vbYesNo, "Confirm Undo Action")
                          
    If userResponse = vbNo Then
        MsgBox "Operation cancelled by user.", vbInformation
        Exit Sub
    End If

    Debug.Print "--- Starting UndoLlmReview process for author '" & authorToUndo & "' at " & Now() & " ---"
    Application.ScreenUpdating = False

    ' --- 3. Reject Tracked Changes from the specified author ---
    For i = ActiveDocument.Revisions.Count To 1 Step -1
        Set rev = ActiveDocument.Revisions(i)
        If rev.Author = authorToUndo Then
            rev.Reject
            rejectedChangesCount = rejectedChangesCount + 1
        End If
    Next i

    ' --- 4. Delete Comments from the specified author ---
    For i = ActiveDocument.Comments.Count To 1 Step -1
        Set com = ActiveDocument.Comments(i)
        If com.Author = authorToUndo Then
            com.Delete
            deletedCommentsCount = deletedCommentsCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Undo process complete!" & vbCrLf & vbCrLf & _
           "Summary for '" & authorToUndo & "':" & vbCrLf & _
           " - Revisions Rejected: " & rejectedChangesCount & vbCrLf & _
           " - Comments Deleted: " & deletedCommentsCount, _
           vbInformation, "Undo Complete"
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred during the undo process." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Undo Error"
    Application.ScreenUpdating = True
End Sub

' =========================================================================================
' === PreProcessJson (CLEANS THE RAW JSON STRING) =========================================
' =========================================================================================
Private Function PreProcessJson(ByVal jsonString As String) As String
    ' This function cleans common LLM output errors before parsing.
    Debug.Print "  - Pre-processing JSON string..."
    Dim temp As String
    temp = jsonString
    ' 1. Replace Word/HTML smart quotes with standard quotes (only when acting as JSON delimiters)
    temp = FixSmartJsonQuotes(temp)
    ' 2. Remove trailing commas that break parsers (e.g., [1, 2, ])
    temp = Replace(temp, ",]", "]")
    temp = Replace(temp, ",}", "}")
    ' 3. Merge adjacent top-level arrays like "][" into ","
    ' (This is a common error when multiple JSON snippets are pasted together)
    If InStr(1, temp, "][", vbBinaryCompare) > 0 Then
        temp = Replace(temp, "][", ",")
    End If
    PreProcessJson = temp
    Debug.Print "  -> Pre-processing complete."
End Function

Private Function FixSmartJsonQuotes(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    Dim inString As Boolean
    Dim esc As Boolean

    result = ""
    inString = False
    esc = False

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If inString Then
            If esc Then
                esc = False
            ElseIf ch = "\\" Then
                esc = True
            ElseIf ch = "\""" Then
                inString = False
            ElseIf IsSmartDoubleQuote(ch) Then
                Dim nextNonWs As String
                nextNonWs = NextNonWhitespaceChar(s, i + 1)
                If nextNonWs = "" Or InStr(1, ",:}]", nextNonWs, vbBinaryCompare) > 0 Then
                    ch = "\"""
                    inString = False
                End If
            End If
        Else
            If ch = "\""" Then
                inString = True
            ElseIf IsSmartDoubleQuote(ch) Then
                ch = "\"""
                inString = True
            End If
        End If

        result = result & ch
    Next i

    FixSmartJsonQuotes = result
End Function

Private Function IsSmartDoubleQuote(ByVal ch As String) As Boolean
    IsSmartDoubleQuote = (ch = Chr(147) Or ch = Chr(148) Or ch = ChrW(8220) Or ch = ChrW(8221))
End Function

Private Function NextNonWhitespaceChar(ByVal s As String, ByVal startPos As Long) As String
    Dim j As Long
    For j = startPos To Len(s)
        Dim c As String
        c = Mid$(s, j, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then
            NextNonWhitespaceChar = c
            Exit Function
        End If
    Next j
    NextNonWhitespaceChar = ""
End Function

' =========================================================================================
' === JSON PARSER (VBA-JSON by Tim Hall) - REQUIRED FOR LLM_ParseJson CALL ==============
' =========================================================================================
' Included directly for portability. 64-bit compatible.
' Source: https://github.com/VBA-tools/VBA-JSON
' License: MIT
'------------------------------------------------------------------------------------------
Private Function LLM_ParseJson(ByVal jsonString As String) As Object
    On Error GoTo ParseError
    If jsonString = "" Then Exit Function
    Debug.Print "  - Parsing JSON (len=" & Len(jsonString) & ")..."
    p_Json = jsonString
    p_Index = 1
    p_ParseError = PARSE_SUCCESS
    
    Dim vParsed As Variant
    llm_ParseValue vParsed
    
    If p_ParseError <> PARSE_SUCCESS Then
        Set LLM_ParseJson = Nothing
    ElseIf IsObject(vParsed) Then
        Set LLM_ParseJson = vParsed
    Else
        Set LLM_ParseJson = Nothing
    End If
    Debug.Print "  -> Parsing complete."
    Exit Function

ParseError:
    HandleError "LLM_ParseJson", Err
    Set LLM_ParseJson = Nothing
End Function

Private Sub llm_ParseValue(ByRef result As Variant)
    On Error GoTo ErrorHandler
    pv_EatWhitespace
    Dim FirstChar As String
    FirstChar = Mid$(p_Json, p_Index, 1)
    
    Select Case FirstChar
        Case "{"
            Set result = pv_ParseObject()
        Case "["
            Set result = pv_ParseArray()
        Case """"
            result = pv_ParseString()
        Case "t", "f"
            result = pv_ParseBoolean()
        Case "n"
            result = pv_ParseNull()
        Case Else
            If IsNumeric(FirstChar) Or FirstChar = "-" Then
                result = pv_ParseNumber()
            Else
                p_ParseError = PARSE_INVALID_JSON_TYPE
            End If
    End Select
    Exit Sub

ErrorHandler:
    HandleError "llm_ParseValue", Err
End Sub

' =========================================================================================
' === UTILITY SUBROUTINE TO CHECK FOR REMAINING CHANGES ===================================
' =========================================================================================

Private Sub FinalCheckForRemainingChanges()
    ' This macro checks if any revisions or comments by the current user
    ' are still present in the document and provides a report.

    Dim rev As Revision
    Dim com As Comment
    Dim revisionCount As Long
    Dim commentCount As Long
    Dim authorName As String
    
    ' Check for changes made by the current user, to align with the review form's logic.
    authorName = Application.UserName

    ' --- 1. Count remaining revisions by the current user ---
    For Each rev In ActiveDocument.Revisions
        If rev.Author = authorName Then
            revisionCount = revisionCount + 1
        End If
    Next rev

    ' --- 2. Count remaining comments by the current user ---
    For Each com In ActiveDocument.Comments
        If com.Author = authorName Then
            commentCount = commentCount + 1
        End If
    Next com

    ' --- 3. Report the results to the user ---
    If revisionCount = 0 And commentCount = 0 Then
        MsgBox "Verification complete: No remaining changes or comments found for user '" & authorName & "'.", _
               vbInformation, "Final Check Passed"
    Else
        MsgBox "Warning: There are still items from '" & authorName & "' in the document." & vbCrLf & vbCrLf & _
               "  - Remaining Revisions: " & revisionCount & vbCrLf & _
               "  - Remaining Comments: " & commentCount & vbCrLf & vbCrLf & _
               "You may want to continue reviewing.", _
               vbExclamation, "Final Check Warning"
    End If
End Sub

' =========================================================================================
' === TABLE ROW INSERTION/DELETION FUNCTIONS (ADD TO wordAIreviewer.bas) ================
' =========================================================================================
'
' Add these functions to wordAIreviewer.bas to support insert_table_row and delete_table_row actions
'

' Add these Case statements to ExecuteSingleAction (around line 1070):
'
'        Case "insert_table_row"
'            Debug.Print "Action 'insert_table_row': Adding new row to table."
'            If Not HasDictionaryKey(actionObject, "tableCell") Then
'                Err.Raise vbObjectError + 522, "ExecuteSingleAction", "insert_table_row requires 'tableCell' structure."
'            End If
'
'            Dim insertPos As String
'            insertPos = "after" ' default
'            If HasDictionaryKey(actionObject, "insertPosition") Then
'                insertPos = LCase$(Trim$(GetSuggestionText(actionObject, "insertPosition", "after")))
'            End If
'
'            Call InsertTableRow(actionObject, topLevelSuggestion, matchCase, insertPos, replaceText)
'
'        Case "delete_table_row"
'            Debug.Print "Action 'delete_table_row': Removing row from table."
'            If Not HasDictionaryKey(actionObject, "tableCell") Then
'                Err.Raise vbObjectError + 523, "ExecuteSingleAction", "delete_table_row requires 'tableCell' structure."
'            End If
'
'            Call DeleteTableRow(actionObject, topLevelSuggestion, matchCase)


' =========================================================================================
' New Function: InsertTableRow
' =========================================================================================
Private Sub InsertTableRow(ByVal actionObject As Object, _
                          ByVal topLevelSuggestion As Object, _
                          ByVal matchCase As Boolean, _
                          ByVal insertPosition As String, _
                          ByVal rowData As String)
    ' Inserts a new row into a table at the specified position
    ' Copies formatting from adjacent row
    ' Populates cells with pipe-separated values from rowData

    On Error GoTo ErrorHandler

    ' Step 1: Find the target cell to identify the row
    Dim context As String
    context = GetSuggestionContextText(topLevelSuggestion)
    Dim contextForSearch As String
    contextForSearch = NormalizeForDocument(context)

    Dim searchRange As Range
    Set searchRange = ActiveDocument.Content

    Dim targetCell As Range
    Set targetCell = FindTableCell(actionObject, searchRange)

    If targetCell Is Nothing Then
        Err.Raise vbObjectError + 524, "InsertTableRow", "Could not find target cell for row insertion."
    End If

    ' Step 2: Get the table and row number
    Dim tbl As Table
    Set tbl = targetCell.Tables(1)

    Dim targetRowNum As Long
    targetRowNum = targetCell.Cells(1).RowIndex

    Debug.Print "  - Inserting row " & IIf(insertPosition = "before", "before", "after") & " row " & targetRowNum

    ' Step 3: Insert new row
    Dim newRow As Row
    Dim rowCount As Long
    rowCount = GetSafeTableRowCount(tbl)

    If rowCount = 0 Then
        Err.Raise vbObjectError + 526, "InsertTableRow", "Unable to determine row count for this merged table."
    End If

    On Error Resume Next
    If insertPosition = "before" Then
        Set newRow = tbl.Rows.Add(tbl.Rows(targetRowNum))
        ' newRow is now BEFORE targetRowNum (insertion shifts indices)
        If Err.Number <> 0 Then
            Err.Clear
            Set newRow = tbl.Rows.Add
        End If
    Else ' "after"
        If targetRowNum < rowCount Then
            Set newRow = tbl.Rows.Add(tbl.Rows(targetRowNum + 1))
            If Err.Number <> 0 Then
                Err.Clear
                Set newRow = tbl.Rows.Add
            End If
        Else
            ' Adding after last row
            Set newRow = tbl.Rows.Add
        End If
    End If
    On Error GoTo ErrorHandler

    If newRow Is Nothing Then
        Err.Raise vbObjectError + 527, "InsertTableRow", "Unable to insert row in this merged table structure."
    End If

    ' Step 4: Copy formatting from adjacent row
    Dim sourceRowNum As Long
    If insertPosition = "before" Then
        sourceRowNum = targetRowNum + 1 ' The row we inserted before (now shifted)
    Else
        sourceRowNum = targetRowNum ' The row we inserted after
    End If

    Call CopyRowFormatting(tbl, sourceRowNum, newRow.Index)

    ' Step 5: Populate cells with data
    If Len(Trim$(rowData)) > 0 Then
        Dim cellValues() As String
        cellValues = Split(rowData, "|")

        Dim colIndex As Long
        For colIndex = 0 To UBound(cellValues)
            If colIndex + 1 <= newRow.Cells.Count Then
                Dim cellRange As Range
                Set cellRange = newRow.Cells(colIndex + 1).Range
                ' Remove cell end marker
                If cellRange.End > cellRange.Start Then
                    cellRange.End = cellRange.End - 1
                End If
                cellRange.Text = Trim$(cellValues(colIndex))
                ' Apply formatting if cell contains HTML tags
                If InStr(1, cellValues(colIndex), "<") > 0 Then
                    ApplyFormattedReplacement cellRange, Trim$(cellValues(colIndex))
                End If
            End If
        Next colIndex
    End If

    Debug.Print "  -> Row inserted successfully at position " & newRow.Index
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "InsertTableRow: " & Err.Source, Err.Description
End Sub

' =========================================================================================
' New Function: DeleteTableRow
' =========================================================================================
Private Sub DeleteTableRow(ByVal actionObject As Object, _
                          ByVal topLevelSuggestion As Object, _
                          ByVal matchCase As Boolean)
    ' Deletes a row from a table

    On Error GoTo ErrorHandler

    ' Step 1: Find the target cell to identify the row
    Dim context As String
    context = GetSuggestionContextText(topLevelSuggestion)
    Dim contextForSearch As String
    contextForSearch = NormalizeForDocument(context)

    Dim searchRange As Range
    Set searchRange = ActiveDocument.Content

    Dim targetCell As Range
    Set targetCell = FindTableCell(actionObject, searchRange)

    If targetCell Is Nothing Then
        Err.Raise vbObjectError + 525, "DeleteTableRow", "Could not find target cell for row deletion."
    End If

    ' Step 2: Get the table and row number
    Dim tbl As Table
    Set tbl = targetCell.Tables(1)

    Dim targetRowNum As Long
    targetRowNum = targetCell.Cells(1).RowIndex

    Debug.Print "  - Deleting row " & targetRowNum & " from table"

    ' Step 3: Safety check - don't delete header row unless explicitly confirmed
    If targetRowNum = 1 Then
        Debug.Print "  - WARNING: Attempting to delete first row (likely header)"
        ' Proceed anyway - user can undo if needed
    End If

    ' Step 4: Delete the row
    tbl.Rows(targetRowNum).Delete

    Debug.Print "  -> Row deleted successfully"
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "DeleteTableRow: " & Err.Source, Err.Description
End Sub

' =========================================================================================
' New Helper Function: CopyRowFormatting
' =========================================================================================
Private Sub CopyRowFormatting(ByVal tbl As Table, _
                              ByVal sourceRowIndex As Long, _
                              ByVal targetRowIndex As Long)
    ' Copies all formatting from source row to target row
    ' Includes: cell shading, borders, font, alignment, height

    On Error Resume Next ' Some formatting properties may not be available

    Debug.Print "  - Copying formatting from row " & sourceRowIndex & " to row " & targetRowIndex

    Dim sourceRow As Row
    Dim targetRow As Row
    Set sourceRow = tbl.Rows(sourceRowIndex)
    Set targetRow = tbl.Rows(targetRowIndex)

    ' Copy row height
    targetRow.Height = sourceRow.Height
    targetRow.HeightRule = sourceRow.HeightRule

    ' Copy cell-by-cell formatting
    Dim colIndex As Long
    Dim maxCols As Long
    maxCols = sourceRow.Cells.Count
    If targetRow.Cells.Count < maxCols Then maxCols = targetRow.Cells.Count

    For colIndex = 1 To maxCols
        Dim sourceCell As Cell
        Dim targetCell As Cell
        Set sourceCell = sourceRow.Cells(colIndex)
        Set targetCell = targetRow.Cells(colIndex)

        ' Copy cell shading
        targetCell.Shading.BackgroundPatternColor = sourceCell.Shading.BackgroundPatternColor

        ' Copy borders
        targetCell.Borders(wdBorderTop).LineStyle = sourceCell.Borders(wdBorderTop).LineStyle
        targetCell.Borders(wdBorderBottom).LineStyle = sourceCell.Borders(wdBorderBottom).LineStyle
        targetCell.Borders(wdBorderLeft).LineStyle = sourceCell.Borders(wdBorderLeft).LineStyle
        targetCell.Borders(wdBorderRight).LineStyle = sourceCell.Borders(wdBorderRight).LineStyle

        ' Copy cell content formatting (font, alignment)
        Dim sourceCellRange As Range
        Dim targetCellRange As Range
        Set sourceCellRange = sourceCell.Range
        Set targetCellRange = targetCell.Range

        ' Remove cell markers for proper range
        If sourceCellRange.End > sourceCellRange.Start Then
            sourceCellRange.End = sourceCellRange.End - 1
        End If
        If targetCellRange.End > targetCellRange.Start Then
            targetCellRange.End = targetCellRange.End - 1
        End If

        ' Copy font properties
        targetCellRange.Font.Name = sourceCellRange.Font.Name
        targetCellRange.Font.Size = sourceCellRange.Font.Size
        targetCellRange.Font.Bold = sourceCellRange.Font.Bold
        targetCellRange.Font.Italic = sourceCellRange.Font.Italic
        targetCellRange.Font.Color = sourceCellRange.Font.Color

        ' Copy paragraph alignment
        targetCellRange.ParagraphFormat.Alignment = sourceCellRange.ParagraphFormat.Alignment

        ' Copy vertical alignment
        targetCell.VerticalAlignment = sourceCell.VerticalAlignment
    Next colIndex

    On Error GoTo 0
    Debug.Print "  - Formatting copied successfully"
End Sub

' =========================================================================================
' INTEGRATION NOTES:
' =========================================================================================
'
' 1. Add the Case statements to ExecuteSingleAction around line 1070 (after "replace_with_table")
'
' 2. Add these three new functions (InsertTableRow, DeleteTableRow, CopyRowFormatting)
'    to the end of the module before the final End statements
'
' 3. Update the action sorting in ProcessSuggestion (around line 716) to include:
'    Case "insert_table_row"
'        cOther.Add actionObject  ' Or create new collection cRowOps if you want specific ordering
'    Case "delete_table_row"
'        cOther.Add actionObject
'
' 4. These actions will work with Track Changes enabled - insertions/deletions will be tracked
'
' 5. The tableCell structure is REQUIRED for both actions - this ensures precise row identification
'
' =========================================================================================


' === LAUNCHER FOR THE REVIEW TOOL ========================================================
 
' =========================================================================================
 
Private Sub StartAiReview()

    ' This sub launches the modeless review form.

    Dim reviewForm As New frmReviewer
    reviewForm.Show
End Sub

' =========================================================================================
' === PUBLIC ENTRY POINTS FOR DSM (V4 Architecture) ======================================
' =========================================================================================

Private Sub GenerateAndShowStructureMap()
    ' Public entry point to generate and display the Document Structure Map
    ' Can be called from ribbon, form, or directly for testing
    
    On Error GoTo ErrorHandler
    
    Debug.Print vbCrLf & "================ GENERATING DOCUMENT STRUCTURE MAP ================"
    
    ' Clear any existing map
    ClearDocumentStructureMap
    
    ' Build new map
    BuildDocumentStructureMap ActiveDocument
    
    If g_DocumentMapCount = 0 Then
        MsgBox "Document is empty or structure map could not be built.", vbExclamation, "Structure Map"
        Exit Sub
    End If
    
    ' Export as markdown
    Dim markdown As String
    markdown = ExportStructureMapAsMarkdown()
    
    ' Avoid printing very large DSM payloads to the Immediate window.
    Debug.Print "DSM markdown generated (" & Len(markdown) & " chars)"
    
    ' Copy to clipboard
    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText markdown
    dataObj.PutInClipboard
    
    MsgBox "Document Structure Map generated!" & vbCrLf & vbCrLf & _
           "Elements mapped: " & g_DocumentMapCount & vbCrLf & _
           "The map has been copied to your clipboard." & vbCrLf & vbCrLf & _
           "Paste it into your LLM prompt to get structured suggestions.", _
           vbInformation, "Structure Map Ready"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating structure map: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub TestTargetResolution()
    ' Test function to verify target resolution works correctly
    
    On Error Resume Next
    
    Debug.Print vbCrLf & "================ TESTING TARGET RESOLUTION ================"
    
    ' Build structure map
    If Not g_DocumentMapBuilt Then
        BuildDocumentStructureMap ActiveDocument
    End If
    
    ' Test various target formats
    Dim testTargets() As String
    testTargets = Split("P1,P5,T1,T1.R1,T1.R2.C1,T1.H.C1", ",")
    
    Dim target As Variant
    Dim rng As Range
    
    For Each target In testTargets
        Set rng = ResolveTargetToRange(CStr(target))
        If Not rng Is Nothing Then
            Debug.Print "✓ " & target & " -> Found at position " & rng.Start & "-" & rng.End
            Debug.Print "   Text preview: " & Left$(rng.Text, 50)
        Else
            Debug.Print "✗ " & target & " -> NOT FOUND"
        End If
    Next target
    
    Debug.Print "================ TEST COMPLETE ================"
End Sub

' =========================================================================================
' === V4 TOOL-CALL PIPELINE (STRICT SCHEMA) ===============================================
' =========================================================================================

Private Function JsonEscape(ByVal inputText As String) As String
    Dim s As String
    s = inputText
    s = Replace(s, "\", "\\")
    s = Replace(s, """", Chr$(92) & """")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

Private Function RangeHasDeletionRevisions(ByVal sourceRange As Range) As Boolean
    On Error GoTo Fail

    Dim rev As Revision
    For Each rev In sourceRange.Revisions
        If rev.Type = wdRevisionDelete Then
            If rev.Range.End > sourceRange.Start And rev.Range.Start < sourceRange.End Then
                RangeHasDeletionRevisions = True
                Exit Function
            End If
        End If
    Next rev

    RangeHasDeletionRevisions = False
    Exit Function

Fail:
    RangeHasDeletionRevisions = False
End Function

Private Function BuildTaggedTextFromRangeFinalView(ByVal sourceRange As Range) As String
    On Error GoTo Fail

    Dim taggedText As String
    Dim spansJson As String
    BuildTaggedAndSpansFromRangeFinalView sourceRange, taggedText, spansJson
    BuildTaggedTextFromRangeFinalView = taggedText
    Exit Function

Fail:
    BuildTaggedTextFromRangeFinalView = GetRangeTextFinalView(sourceRange)
End Function

Private Function BuildFormatSpansJsonFinalView(ByVal sourceRange As Range) As String
    On Error GoTo Fail

    Dim taggedText As String
    Dim spansJson As String
    BuildTaggedAndSpansFromRangeFinalView sourceRange, taggedText, spansJson
    BuildFormatSpansJsonFinalView = spansJson
    Exit Function

Fail:
    BuildFormatSpansJsonFinalView = "[]"
End Function

Private Function RangeHasAnyInlineFormatting(ByVal sourceRange As Range) As Boolean
    ' Fast-path check: if the whole range reports plain formatting, skip character scanning.

    On Error GoTo Fail

    If sourceRange Is Nothing Then
        RangeHasAnyInlineFormatting = False
        Exit Function
    End If

    With sourceRange.Font
        If .Bold = 0 And .Italic = 0 And .Subscript = 0 And .Superscript = 0 Then
            RangeHasAnyInlineFormatting = False
        Else
            RangeHasAnyInlineFormatting = True
        End If
    End With
    Exit Function

Fail:
    ' If uncertain, keep robust behavior by taking the detailed path.
    RangeHasAnyInlineFormatting = True
End Function

Private Sub BuildTaggedAndSpansFromRange(ByVal sourceRange As Range, ByRef taggedText As String, ByRef spansJson As String)
    On Error GoTo Fail

    Dim plainText As String
    plainText = sourceRange.Text
    If Len(plainText) = 0 Then
        taggedText = ""
        spansJson = "[]"
        Exit Sub
    End If

    If Not RangeHasAnyInlineFormatting(sourceRange) Then
        taggedText = plainText
        spansJson = "[]"
        Exit Sub
    End If

    Dim i As Long
    Dim out As String
    Dim spans As String
    Dim inSub As Boolean
    Dim inSup As Boolean
    Dim inBold As Boolean
    Dim inItalic As Boolean
    Dim hasRun As Boolean
    Dim runStart As Long
    Dim runLen As Long
    Dim runSub As Boolean
    Dim runSup As Boolean
    Dim runBold As Boolean
    Dim runItalic As Boolean

    For i = 1 To Len(plainText)
        Dim chRange As Range
        Set chRange = sourceRange.Duplicate
        chRange.Start = sourceRange.Start + i - 1
        chRange.End = chRange.Start + 1

        Dim isSub As Boolean
        Dim isSup As Boolean
        Dim isBold As Boolean
        Dim isItalic As Boolean

        isSub = (chRange.Font.Subscript = True)
        isSup = (chRange.Font.Superscript = True)
        isBold = (chRange.Font.Bold = True)
        isItalic = (chRange.Font.Italic = True)

        ' Build tagged text.
        If inSub And Not isSub Then out = out & "</sub>": inSub = False
        If inSup And Not isSup Then out = out & "</sup>": inSup = False
        If inBold And Not isBold Then out = out & "</b>": inBold = False
        If inItalic And Not isItalic Then out = out & "</i>": inItalic = False

        If isItalic And Not inItalic Then out = out & "<i>": inItalic = True
        If isBold And Not inBold Then out = out & "<b>": inBold = True
        If isSup And Not inSup Then out = out & "<sup>": inSup = True
        If isSub And Not inSub Then out = out & "<sub>": inSub = True

        out = out & Mid$(plainText, i, 1)

        ' Build spans.
        If Not (isSub Or isSup Or isBold Or isItalic) Then
            If hasRun Then
                spans = spans & IIf(Len(spans) > 0, ",", "") & _
                    "{""start"":" & runStart & ",""length"":" & runLen & ",""subscript"":" & LCase$(CStr(runSub)) & _
                    ",""superscript"":" & LCase$(CStr(runSup)) & ",""bold"":" & LCase$(CStr(runBold)) & _
                    ",""italic"":" & LCase$(CStr(runItalic)) & "}"
                hasRun = False
            End If
            GoTo ContinueLoop
        End If

        If Not hasRun Then
            runStart = i
            runLen = 1
            runSub = isSub
            runSup = isSup
            runBold = isBold
            runItalic = isItalic
            hasRun = True
        ElseIf runSub = isSub And runSup = isSup And runBold = isBold And runItalic = isItalic Then
            runLen = runLen + 1
        Else
            spans = spans & IIf(Len(spans) > 0, ",", "") & _
                "{""start"":" & runStart & ",""length"":" & runLen & ",""subscript"":" & LCase$(CStr(runSub)) & _
                ",""superscript"":" & LCase$(CStr(runSup)) & ",""bold"":" & LCase$(CStr(runBold)) & _
                ",""italic"":" & LCase$(CStr(runItalic)) & "}"
            runStart = i
            runLen = 1
            runSub = isSub
            runSup = isSup
            runBold = isBold
            runItalic = isItalic
            hasRun = True
        End If

ContinueLoop:
    Next i

    If inSub Then out = out & "</sub>"
    If inSup Then out = out & "</sup>"
    If inBold Then out = out & "</b>"
    If inItalic Then out = out & "</i>"

    If hasRun Then
        spans = spans & IIf(Len(spans) > 0, ",", "") & _
            "{""start"":" & runStart & ",""length"":" & runLen & ",""subscript"":" & LCase$(CStr(runSub)) & _
            ",""superscript"":" & LCase$(CStr(runSup)) & ",""bold"":" & LCase$(CStr(runBold)) & _
            ",""italic"":" & LCase$(CStr(runItalic)) & "}"
    End If

    taggedText = out
    spansJson = "[" & spans & "]"
    Exit Sub

Fail:
    taggedText = sourceRange.Text
    spansJson = "[]"
End Sub

Private Sub BuildTaggedAndSpansFromRangeFinalView(ByVal sourceRange As Range, ByRef taggedText As String, ByRef spansJson As String)
    On Error GoTo Fail

    If RangeHasDeletionRevisions(sourceRange) Then
        taggedText = GetRangeTextFinalView(sourceRange)
        spansJson = "[]"
    Else
        BuildTaggedAndSpansFromRange sourceRange, taggedText, spansJson
    End If
    Exit Sub

Fail:
    taggedText = GetRangeTextFinalView(sourceRange)
    spansJson = "[]"
End Sub

Private Function BuildTaggedTextFromRange(ByVal sourceRange As Range) As String
    On Error GoTo Fail

    Dim taggedText As String
    Dim spansJson As String
    BuildTaggedAndSpansFromRange sourceRange, taggedText, spansJson
    BuildTaggedTextFromRange = taggedText
    Exit Function

Fail:
    BuildTaggedTextFromRange = sourceRange.Text
End Function

Private Function BuildFormatSpansJson(ByVal sourceRange As Range) As String
    On Error GoTo Fail

    Dim taggedText As String
    Dim spansJson As String
    BuildTaggedAndSpansFromRange sourceRange, taggedText, spansJson
    BuildFormatSpansJson = spansJson
    Exit Function

Fail:
    BuildFormatSpansJson = "[]"
End Function

Private Function BuildJsonStringArray(ByVal items As Variant) As String
    Dim result As String
    Dim i As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    If IsEmpty(items) Then
        BuildJsonStringArray = "[]"
        Exit Function
    End If

    If Not IsArray(items) Then
        BuildJsonStringArray = "[]"
        Exit Function
    End If

    On Error GoTo EmptyArray
    lowerBound = LBound(items)
    upperBound = UBound(items)
    On Error GoTo 0

    result = "["
    For i = lowerBound To upperBound
        If i > lowerBound Then
            result = result & ","
        End If
        result = result & """" & JsonEscape(CStr(items(i))) & """"
    Next i
    result = result & "]"
    BuildJsonStringArray = result
    Exit Function

EmptyArray:
    BuildJsonStringArray = "[]"
End Function

Private Sub AddToolDefinition(ByVal name As String, ByVal action As String, ByVal description As String, ByVal targetRequirement As String, _
    ByVal requiredArgs As Variant, ByVal optionalArgs As Variant, ByVal exampleTarget As String, _
    ByVal exampleArgsJson As String, ByVal exampleExplanation As String, ByVal notes As Variant)

    Dim def As Object
    Set def = CreateObject("Scripting.Dictionary")

    def("name") = name
    def("action") = action
    def("description") = description
    def("target_requirement") = targetRequirement
    def("required_args") = requiredArgs
    def("optional_args") = optionalArgs
    def("example_target") = exampleTarget
    def("example_args_json") = exampleArgsJson
    def("example_explanation") = exampleExplanation
    def("notes") = notes

    g_ToolRegistry.Add LCase$(name), def
    g_ToolRegistryOrder.Add def
End Sub

Private Sub InitializeToolRegistry()
    If Not g_ToolRegistry Is Nothing Then Exit Sub

    Set g_ToolRegistry = CreateObject("Scripting.Dictionary")
    Set g_ToolRegistryOrder = New Collection

    AddToolDefinition "replace_text", "replace", _
        "Find and replace text strictly within the resolved target range.", _
        "Use paragraph IDs (P#) or table cell references (T#.R#.C#) so replacements stay scoped.", _
        Array("find", "replace"), _
        Array("match_case"), _
        "P12", _
        "{""find"":""recieved"",""replace"":""received""}", _
        "Correct spelling or update terminology without touching adjacent content.", _
        Array("Never rely on document-wide search—provide the exact text expected inside the target.", _
              "You may use <b>, <i>, <sub>, and <sup> tags in the replace string when formatting is required.")

    AddToolDefinition "apply_style", "apply_style", _
        "Apply a predefined style to the resolved range.", _
        "Works best with full-paragraph targets (P#).", _
        Array("style"), _
        Empty, _
        "P3", _
        "{""style"":""heading_l2""}", _
        "Promote or demote headings to match the report outline.", _
        Array("Supported tokens appear in the tooling.style_tokens list (e.g., heading_l2, table_heading).", _
              "If a VA macro is unavailable the macro falls back to direct Word style application.")

    AddToolDefinition "add_comment", "comment", _
        "Insert a Word comment anchored to the target range.", _
        "Target can be any paragraph, table, row, or cell reference.", _
        Empty, _
        Empty, _
        "P8", _
        "{""text"":""Clarify the propagation model assumptions.""}", _
        "Request clarifications or note issues that cannot be auto-fixed.", _
        Array("If args.text is omitted the macro uses the suggestion's explanation, but providing text is preferred.")

    AddToolDefinition "delete_range", "delete", _
        "Remove the resolved target range (paragraph, row, or cell).", _
        "Target must be a precise DSM ID; no partial matches are attempted.", _
        Empty, _
        Empty, _
        "P15", _
        "{}", _
        "Remove duplicated sentences or redundant cells.", _
        Array("Use sparingly—deleting entire sections should be accompanied by a comment explaining the rationale.")

    AddToolDefinition "replace_table", "replace_table", _
        "Replace an entire table with markdown content that will be converted to a Word table.", _
        "Target must be the table ID (T#).", _
        Array("markdown"), _
        Empty, _
        "T2", _
        "{""markdown"":""| Location | Leq dB |\\n|---|---|\\n| Roof | 55 |""}", _
        "Provide complete updated data when layout changes are easier than incremental edits.", _
        Array("Include a header row and keep the column count consistent with the intended output.")

    AddToolDefinition "insert_table_row", "insert_row", _
        "Insert a row into the specified table.", _
        "Target must be the table ID (T#).", _
        Array("data"), _
        Array("after_row"), _
        "T3", _
        "{""data"":[""Location 4"",""47 dB"",""Day""],""after_row"":3}", _
        "Add missing measurement rows while preserving header and formatting.", _
        Array("Provide one entry per column; omit after_row or set to 0 to insert at the top.")

    AddToolDefinition "delete_table_row", "delete_row", _
        "Delete a specific table row (including header rows when appropriate).", _
        "Target must be a row reference such as T2.R5 or T2.H.", _
        Empty, _
        Empty, _
        "T2.R5", _
        "{}", _
        "Use for duplicated or obsolete data rows.", _
        Array("This removes the entire row; update nearby references or totals if needed.")
End Sub

Private Function GetToolRegistry() As Object
    InitializeToolRegistry
    Set GetToolRegistry = g_ToolRegistry
End Function

Private Function GetToolDefinition(ByVal toolName As String) As Object
    Dim registry As Object
    Dim normalized As String

    normalized = LCase$(Trim$(toolName))
    Set registry = GetToolRegistry()

    If registry.Exists(normalized) Then
        Set GetToolDefinition = registry(normalized)
    Else
        Set GetToolDefinition = Nothing
    End If
End Function

Private Function BuildToolEntry(ByVal name As String, ByVal description As String, ByVal targetRequirement As String, _
    ByVal requiredArgs As Variant, ByVal optionalArgs As Variant, ByVal exampleTarget As String, _
    ByVal exampleArgsJson As String, ByVal exampleExplanation As String, ByVal notes As Variant) As String

    Dim entry As String
    entry = "{""name"":""" & JsonEscape(name) & """," & _
            """description"":""" & JsonEscape(description) & """," & _
            """target_requirements"":""" & JsonEscape(targetRequirement) & """," & _
            """args"":{""required"":" & BuildJsonStringArray(requiredArgs) & ",""optional"":" & BuildJsonStringArray(optionalArgs) & "}," & _
            """notes"":" & BuildJsonStringArray(notes) & "," & _
            """example"":{""tool"":""" & JsonEscape(name) & """,""target"":""" & JsonEscape(exampleTarget) & """,""args"":" & exampleArgsJson & ",""explanation"":""" & JsonEscape(exampleExplanation) & """}}"

    BuildToolEntry = entry
End Function

Private Function BuildTargetReferenceDocsJson() As String
    Dim entries As Variant
    Dim entry As Variant
    Dim i As Long
    Dim output As String

    entries = Array( _
        Array("paragraph", "P{n}", "Use for any paragraph listed in the DSM elements array. The numeric suffix matches the paragraph order.", "P5"), _
        Array("table", "T{n}", "References an entire table. Use this when replacing a whole table or inserting rows relative to it.", "T2"), _
        Array("table_row", "T{n}.R{r}", "References a specific row within table n (e.g., R1 = first data row).", "T2.R3"), _
        Array("table_header_row", "T{n}.H", "References the table header row (equivalent to row 1).", "T2.H"), _
        Array("table_cell", "T{n}.R{r}.C{c}", "References a single cell using row and column numbers.", "T2.R3.C2"), _
        Array("table_header_cell", "T{n}.H.C{c}", "Targets a header row cell for formatting or text edits.", "T2.H.C1") _
    )

    output = "["
    For i = LBound(entries) To UBound(entries)
        entry = entries(i)
        If i > LBound(entries) Then output = output & ","
        output = output & "{""kind"":""" & JsonEscape(entry(0)) & """," & _
                          """format"":""" & JsonEscape(entry(1)) & """," & _
                          """description"":""" & JsonEscape(entry(2)) & """," & _
                          """example"":""" & JsonEscape(entry(3)) & """}"
    Next i
    output = output & "]"
    BuildTargetReferenceDocsJson = output
End Function

Private Function BuildToolDefinitionJson() As String
    Dim output As String
    Dim i As Long
    Dim def As Object

    InitializeToolRegistry

    output = "["
    If Not g_ToolRegistryOrder Is Nothing Then
        For i = 1 To g_ToolRegistryOrder.Count
            Set def = g_ToolRegistryOrder(i)
            If i > 1 Then output = output & ","
            output = output & BuildToolEntry( _
                def("name"), _
                def("description"), _
                def("target_requirement"), _
                def("required_args"), _
                def("optional_args"), _
                def("example_target"), _
                def("example_args_json"), _
                def("example_explanation"), _
                def("notes"))
        Next i
    End If
    output = output & "]"
    BuildToolDefinitionJson = output
End Function

Private Function BuildResponseContractJson() As String
    Dim responseNotes As Variant

    responseNotes = Array( _
        "Return valid JSON only (no Markdown code fences or commentary).", _
        "Populate tool_calls with deterministic actions ordered by priority—earlier entries run first.", _
        "args must always be a JSON object even when no fields are required (use {}).", _
        "Set explanation to brief natural language so reviewers understand the reason for the change." _
    )

    BuildResponseContractJson = "{""root_object"":""" & JsonEscape("{ ""tool_calls"": [] }") & """," & _
        """required_fields"":" & BuildJsonStringArray(Array("tool_calls")) & "," & _
        """tool_call_entry"":{""required"":" & BuildJsonStringArray(Array("tool", "target", "args")) & ",""optional"":" & BuildJsonStringArray(Array("explanation")) & "}," & _
        """notes"":" & BuildJsonStringArray(responseNotes) & "}"
End Function

Private Function BuildManualFallbackJson() As String
    Dim steps As Variant
    Dim notes As Variant

    steps = Array( _
        "Review the macro summary or log to identify which tool calls failed.", _
        "Use the DSM target IDs (P#, T#, etc.) to locate the exact paragraphs or tables in Word.", _
        "Apply the edit manually or adjust the JSON entry, then rerun V4_ApplyToolCalls when ready." _
    )

    notes = Array( _
        "Failed items leave the document unchanged, so manual edits are safe to perform afterward.", _
        "Turn on Track Changes before manual fixes if you need an audit trail." _
    )

    BuildManualFallbackJson = "{""trigger"":""" & JsonEscape("One or more tool calls failed to apply automatically.") & """," & _
        """steps"":" & BuildJsonStringArray(steps) & "," & _
        """notes"":" & BuildJsonStringArray(notes) & "}"
End Function

Private Function BuildToolingDocumentationJson() As String
    Dim guidelines As Variant
    Dim styleTokens As Variant

    guidelines = Array( _
        "Always anchor targets using DSM IDs from the elements array; never rely on searching raw text.", _
        "Confine replacements to the resolved paragraph or cell to avoid accidental edits elsewhere.", _
        "Paragraph IDs refer to body paragraphs; table cell text should be targeted with T#.R#.C# references.", _
        "Favour table-specific tools (replace_table, insert_table_row, delete_table_row) for structured data changes.", _
        "Add comments when recommending manual review steps or when data is missing.", _
        "Keep tool calls atomic—one logical change per entry." _
    )

    styleTokens = Array("heading_l1", "heading_l2", "heading_l3", "heading_l4", "body_text", "bullet", "table_heading", "table_text", "table_title", "figure")

    BuildToolingDocumentationJson = _
        "{""overview"":""" & JsonEscape("Use these instructions to convert DSM v4.2 data into executable Word tool calls.") & """," & _
        """ordering"":""" & JsonEscape("Elements are listed in document order (top-to-bottom).") & """," & _
        """text_view"":""" & JsonEscape("Paragraph and table previews reflect Word's Final view (insertions included, deletions excluded).") & """," & _
        """guidelines"":" & BuildJsonStringArray(guidelines) & "," & _
        """style_tokens"":" & BuildJsonStringArray(styleTokens) & "," & _
        """target_reference_format"":" & BuildTargetReferenceDocsJson() & "," & _
        """response_contract"":" & BuildResponseContractJson() & "," & _
        """tools"":" & BuildToolDefinitionJson() & "," & _
        """manual_fallback"":" & BuildManualFallbackJson() & "}"
End Function

Private Function ExportStructureMapAsJsonV42() As String
    On Error GoTo ErrorHandler

    If Not g_DocumentMapBuilt Then
        BuildDocumentStructureMap ActiveDocument
    End If

    Dim output As String
    Dim i As Long
    Dim elem As DocumentElement
    Dim elementJson As String
    Dim elementParts() As String
    Dim docHasRevisions As Boolean

    docHasRevisions = False
    On Error Resume Next
    docHasRevisions = (ActiveDocument.Revisions.Count > 0)
    Err.Clear
    On Error GoTo ErrorHandler

    output = "{""version"":""4.2"",""document"":{""name"":""" & JsonEscape(ActiveDocument.Name) & """,""generated_at"":""" & _
             Format(Now, "yyyy-mm-dd\Thh:nn:ss") & """},""tooling"":" & BuildToolingDocumentationJson() & ",""elements"":["

    If g_DocumentMapCount > 0 Then
        ReDim elementParts(1 To g_DocumentMapCount)
    End If

    For i = 1 To g_DocumentMapCount
        elem = g_DocumentMap(i)
        elementJson = ""
        If i Mod 25 = 0 Then
            TraceLog "  -> DSM JSON export progress: " & i & "/" & g_DocumentMapCount
            DoEvents
        End If

        If elem.ElementType = "paragraph" Then
            Dim pRange As Range
            Dim pFinalText As String
            Dim pTaggedText As String
            Dim pSpansJson As String
            Set pRange = ActiveDocument.Range(elem.StartPos, elem.EndPos)
            If docHasRevisions Then
                pFinalText = GetRangeTextFinalView(pRange)
                BuildTaggedAndSpansFromRangeFinalView pRange, pTaggedText, pSpansJson
            Else
                pFinalText = pRange.Text
                BuildTaggedAndSpansFromRange pRange, pTaggedText, pSpansJson
            End If
            elementJson = "{""id"":""" & elem.ElementID & """,""kind"":""paragraph"",""style"":""" & JsonEscape(elem.StyleName) & """,""text_plain"":""" & _
                          JsonEscape(pFinalText) & """,""text_tagged"":""" & JsonEscape(pTaggedText) & _
                          """,""format_spans"":" & pSpansJson & ",""range"":{""start"":" & elem.StartPos & _
                          ",""end"":" & elem.EndPos & "},""section_number"":" & elem.SectionNumber & ",""page_number"":" & elem.PageNumber & _
                          ",""heading_level"":" & elem.HeadingLevel & ",""within_table"":" & LCase$(CStr(elem.WithinTable)) & "}"
        ElseIf elem.ElementType = "table" Then
            Dim tbl As Table
            Set tbl = ResolveTableByID(elem.ElementID)
            Dim cellsJson As String
            Dim cellFinalText As String
            Dim tableCell As Cell
            Dim cellRange As Range
            Dim exportedCellCount As Long
            Dim cellTagged As String
            Dim cellSpans As String
            Dim cellTaggedAll As String
            Dim cellSpansAll As String

            cellsJson = ""
            If Not tbl Is Nothing Then
                TraceLog "  -> Exporting " & elem.ElementID & " cells..."
                exportedCellCount = 0
                For Each tableCell In tbl.Range.Cells
                    Set cellRange = GetCellContentRange(tableCell)
                    If Not cellRange Is Nothing Then
                        If docHasRevisions Then
                            cellFinalText = Trim$(GetRangeTextFinalView(cellRange))
                        Else
                            cellFinalText = Trim$(cellRange.Text)
                        End If
                        cellTagged = ""
                        cellSpans = "[]"
                        If DSM_INCLUDE_TABLE_CELL_TAGGED_TEXT Or DSM_INCLUDE_TABLE_CELL_FORMAT_SPANS Then
                            If docHasRevisions Then
                                BuildTaggedAndSpansFromRangeFinalView cellRange, cellTaggedAll, cellSpansAll
                            Else
                                BuildTaggedAndSpansFromRange cellRange, cellTaggedAll, cellSpansAll
                            End If
                            If DSM_INCLUDE_TABLE_CELL_TAGGED_TEXT Then cellTagged = cellTaggedAll
                            If DSM_INCLUDE_TABLE_CELL_FORMAT_SPANS Then cellSpans = cellSpansAll
                        End If
                        cellsJson = cellsJson & IIf(Len(cellsJson) > 0, ",", "") & _
                            "{""id"":""" & elem.ElementID & ".R" & tableCell.RowIndex & ".C" & tableCell.ColumnIndex & """,""text_plain"":""" & JsonEscape(cellFinalText) & _
                            """,""text_tagged"":""" & JsonEscape(cellTagged) & """,""format_spans"":" & cellSpans & "}"
                    End If

                    exportedCellCount = exportedCellCount + 1
                    If exportedCellCount Mod DSM_CELL_PROGRESS_INTERVAL = 0 Then DoEvents
                Next tableCell
            End If

            elementJson = "{""id"":""" & elem.ElementID & """,""kind"":""table"",""rows"":" & elem.TableRowCount & ",""cols"":" & elem.TableColCount & _
                          ",""range"":{""start"":" & elem.StartPos & ",""end"":" & elem.EndPos & "},""section_number"":" & elem.SectionNumber & _
                          ",""page_number"":" & elem.PageNumber & ",""title_text"":""" & JsonEscape(elem.TableTitle) & """,""caption_text"":""" & _
                          JsonEscape(elem.TableCaption) & """,""within_table"":false,""cells"":[" & cellsJson & "]}"
        End If

        elementParts(i) = elementJson
    Next i

    If g_DocumentMapCount > 0 Then
        output = output & Join(elementParts, ",")
    End If
    output = output & "]}"
    ExportStructureMapAsJsonV42 = output
    Exit Function

ErrorHandler:
    ExportStructureMapAsJsonV42 = "{""version"":""4.2"",""error"":""" & JsonEscape(Err.Description) & """}"
End Function

Private Function IsAllowedToolName(ByVal toolName As String) As Boolean
    Dim registry As Object
    Dim normalized As String

    normalized = LCase$(Trim$(toolName))
    Set registry = GetToolRegistry()
    IsAllowedToolName = registry.Exists(normalized)
End Function

Private Function GetMissingArgsList(ByVal argsObj As Object, ByVal requiredArgs As Variant) As String
    Dim result As String
    Dim i As Long
    Dim argName As String

    If IsEmpty(requiredArgs) Then
        GetMissingArgsList = ""
        Exit Function
    End If

    On Error GoTo CleanFail
    For i = LBound(requiredArgs) To UBound(requiredArgs)
        argName = CStr(requiredArgs(i))
        If Not HasDictionaryKey(argsObj, argName) Then
            result = result & IIf(Len(result) > 0, ", ", "") & argName
        End If
    Next i
    GetMissingArgsList = result
    Exit Function

CleanFail:
    GetMissingArgsList = ""
End Function

Private Function IsDictionaryLikeObject(ByVal value As Variant) As Boolean
    On Error GoTo Fail

    If Not IsObject(value) Then
        IsDictionaryLikeObject = False
        Exit Function
    End If

    If TypeName(value) = "Collection" Then
        IsDictionaryLikeObject = False
        Exit Function
    End If

    Dim hasKey As Boolean
    hasKey = value.Exists("__probe__")
    Err.Clear
    IsDictionaryLikeObject = True
    Exit Function

Fail:
    Err.Clear
    IsDictionaryLikeObject = False
End Function

Private Function ParseV4ToolCalls(ByVal jsonString As String, ByRef toolCalls As Collection, ByRef validationErrors As String) As Boolean
    On Error GoTo Fail

    ParseV4ToolCalls = False
    Set toolCalls = New Collection
    validationErrors = ""

    Dim parsed As Object
    Dim registry As Object

    Set parsed = LLM_ParseJson(PreProcessJson(jsonString))
    If parsed Is Nothing Then
        validationErrors = "JSON parse failed."
        Exit Function
    End If

    If Not HasDictionaryKey(parsed, "tool_calls") Then
        validationErrors = "Root object must contain 'tool_calls'."
        Exit Function
    End If

    If TypeName(parsed("tool_calls")) <> "Collection" Then
        validationErrors = "'tool_calls' must be an array."
        Exit Function
    End If

    Set registry = GetToolRegistry()

    Dim callObj As Object
    Dim idx As Long
    idx = 0
    For Each callObj In parsed("tool_calls")
        idx = idx + 1

        If Not IsDictionaryLikeObject(callObj) Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": each tool_call must be an object."
            GoTo ContinueLoop
        End If

        If HasDictionaryKey(callObj, "context") Or HasDictionaryKey(callObj, "actions") Or HasDictionaryKey(callObj, "action") Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": legacy V3 keys detected."
            GoTo ContinueLoop
        End If

        If Not HasDictionaryKey(callObj, "tool") Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": missing 'tool'."
            GoTo ContinueLoop
        End If
        If Not HasDictionaryKey(callObj, "target") Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": missing 'target'."
            GoTo ContinueLoop
        End If
        If Not HasDictionaryKey(callObj, "args") Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": missing 'args'."
            GoTo ContinueLoop
        End If

        Dim toolName As String
        Dim toolDef As Object
        toolName = LCase$(Trim$(GetSuggestionText(callObj, "tool", "")))
        If Not registry.Exists(toolName) Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": unsupported tool '" & toolName & "'."
            GoTo ContinueLoop
        End If
        Set toolDef = registry(toolName)

        If Not IsDictionaryLikeObject(callObj("args")) Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & "Item " & idx & ": 'args' must be an object."
            GoTo ContinueLoop
        End If

        Dim argsObj As Object
        Set argsObj = callObj("args")

        Dim missingArgs As String
        missingArgs = GetMissingArgsList(argsObj, toolDef("required_args"))
        If Len(missingArgs) > 0 Then
            validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & _
                "Item " & idx & ": " & toolName & " requires args: " & missingArgs & "."
            GoTo ContinueLoop
        End If

        If toolName = "add_comment" Then
            If Not HasDictionaryKey(argsObj, "text") And Len(GetSuggestionText(callObj, "explanation", "")) = 0 Then
                validationErrors = validationErrors & IIf(Len(validationErrors) > 0, vbCrLf, "") & _
                    "Item " & idx & ": add_comment requires args.text or explanation."
                GoTo ContinueLoop
            End If
        End If

        toolCalls.Add callObj

ContinueLoop:
    Next callObj

    ParseV4ToolCalls = (Len(validationErrors) = 0)
    Exit Function

Fail:
    validationErrors = "Unexpected validator error: " & Err.Description
End Function

Private Function ConvertToolCallToSuggestion(ByVal callObj As Object) As Object
    Dim suggestion As Object
    Set suggestion = NewDictionary()

    Dim toolName As String
    toolName = LCase$(Trim$(GetSuggestionText(callObj, "tool", "")))
    Dim argsObj As Object
    If HasDictionaryKey(callObj, "args") And IsDictionaryLikeObject(callObj("args")) Then
        Set argsObj = callObj("args")
    Else
        Set argsObj = CreateObject("Scripting.Dictionary")
    End If
    Dim toolDef As Object
    Set toolDef = GetToolDefinition(toolName)

    suggestion("target") = GetSuggestionText(callObj, "target", "")
    suggestion("explanation") = GetSuggestionText(callObj, "explanation", "")
    suggestion("tool_name") = toolName
    If Not toolDef Is Nothing Then
        suggestion("action") = toolDef("action")
    End If

    Select Case toolName
        Case "replace_text"
            suggestion("find") = GetSuggestionText(argsObj, "find", "")
            suggestion("replace") = GetSuggestionText(argsObj, "replace", "")
            If HasDictionaryKey(argsObj, "match_case") Then suggestion("match_case") = CBool(argsObj("match_case"))
        Case "apply_style"
            suggestion("style") = GetSuggestionText(argsObj, "style", "")
        Case "add_comment"
            If HasDictionaryKey(argsObj, "text") Then suggestion("explanation") = GetSuggestionText(argsObj, "text", "")
        Case "delete_range"
        Case "replace_table"
            suggestion("replace") = GetSuggestionText(argsObj, "markdown", "")
        Case "insert_table_row"
            If HasDictionaryKey(argsObj, "after_row") Then suggestion("after_row") = CLng(Val(CStr(argsObj("after_row"))))
            If HasDictionaryKey(argsObj, "data") Then suggestion("data") = argsObj("data")
        Case "delete_table_row"
    End Select

    Set ConvertToolCallToSuggestion = suggestion
End Function

Private Function IsToolCallNoOp(ByVal callObj As Object, ByVal targetRange As Range) As Boolean
    On Error GoTo CleanFail
    IsToolCallNoOp = False

    Dim toolName As String
    toolName = LCase$(Trim$(GetSuggestionText(callObj, "tool", "")))
    If toolName <> "replace_text" Then Exit Function

    Dim argsObj As Object
    Set argsObj = callObj("args")
    Dim findText As String
    Dim replaceText As String
    Dim matchCase As Boolean
    findText = GetSuggestionText(argsObj, "find", "")
    replaceText = GetSuggestionText(argsObj, "replace", "")
    matchCase = False
    If HasDictionaryKey(argsObj, "match_case") Then matchCase = CBool(argsObj("match_case"))

    Dim foundRange As Range
    Set foundRange = FindLongString(NormalizeForDocument(findText), targetRange, matchCase)
    If foundRange Is Nothing Then Exit Function

    IsToolCallNoOp = IsFormattingAlreadyApplied(foundRange, replaceText)
    Exit Function

CleanFail:
    IsToolCallNoOp = False
End Function

Private Function DescribeToolFailure(ByVal suggestion As Object) As String
    Dim toolName As String
    Dim target As String
    Dim note As String

    toolName = GetSuggestionText(suggestion, "tool_name", "<tool>")
    target = GetSuggestionText(suggestion, "target", "<target>")
    note = GetSuggestionText(suggestion, "explanation", "")

    If Len(note) > 0 Then
        DescribeToolFailure = toolName & " @ " & target & " - " & note
    Else
        DescribeToolFailure = toolName & " @ " & target
    End If
End Function

Private Function GenerateRunId() As String
    Randomize
    GenerateRunId = Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
End Function

Private Function ElapsedMilliseconds(ByVal startTimer As Single) As Long
    Dim delta As Single
    delta = Timer - startTimer
    If delta < 0 Then delta = delta + 86400 ' Timer resets at midnight.
    ElapsedMilliseconds = CLng(delta * 1000)
End Function

Private Function BuildOutcomeJson(ByVal index As Long, ByVal toolName As String, ByVal targetRef As String, _
    ByVal status As String, ByVal errorCode As String, ByVal message As String) As String

    BuildOutcomeJson = "{""index"":" & index & ",""tool"":""" & JsonEscape(toolName) & """,""target"":""" & _
        JsonEscape(targetRef) & """,""status"":""" & JsonEscape(status) & """,""error_code"":""" & _
        JsonEscape(errorCode) & """,""message"":""" & JsonEscape(message) & """}"
End Function

Private Sub AddOutcome(ByRef outcomes As Collection, ByVal index As Long, ByVal suggestion As Object, _
    ByVal status As String, ByVal errorCode As String, ByVal message As String)

    Dim toolName As String
    Dim targetRef As String
    toolName = GetSuggestionText(suggestion, "tool_name", "<tool>")
    targetRef = GetSuggestionText(suggestion, "target", "<target>")
    outcomes.Add BuildOutcomeJson(index, toolName, targetRef, status, errorCode, message)
End Sub

Private Function BuildOutcomesArrayJson(ByVal outcomes As Collection) As String
    Dim i As Long
    Dim output As String

    output = "["
    For i = 1 To outcomes.Count
        If i > 1 Then output = output & ","
        output = output & CStr(outcomes(i))
    Next i
    output = output & "]"
    BuildOutcomesArrayJson = output
End Function

Private Function BuildRunResultJson(ByVal runId As String, ByVal ok As Boolean, ByVal mode As String, _
    ByVal total As Long, ByVal applied As Long, ByVal skipped As Long, ByVal failed As Long, _
    ByVal validationErrors As String, ByVal outcomes As Collection, ByVal startedAt As Date, _
    ByVal durationMs As Long, ByVal reportPath As String) As String

    BuildRunResultJson = "{""schema"":""v1"",""ok"":" & LCase$(CStr(ok)) & ",""run_id"":""" & JsonEscape(runId) & """," & _
        """mode"":""" & JsonEscape(mode) & """,""counts"":{""total"":" & total & ",""applied"":" & applied & _
        ",""skipped"":" & skipped & ",""failed"":" & failed & "},""validation_errors"":""" & _
        JsonEscape(validationErrors) & """,""outcomes"":" & BuildOutcomesArrayJson(outcomes) & _
        ",""timing"":{""started_at"":""" & Format(startedAt, "yyyy-mm-dd\Thh:nn:ss") & """,""duration_ms"":" & durationMs & "}," & _
        """artifacts"":{""run_report_path"":""" & JsonEscape(reportPath) & """}}"
End Function

Private Function WriteRunResultReport(ByVal runId As String, ByVal resultJson As String) As String
    On Error GoTo Fail

    Dim filePath As String
    Dim fileNum As Integer
    Dim stem As String

    If ActiveDocument Is Nothing Then
        stem = "document"
    Else
        stem = GetDocStem()
    End If

    filePath = GetClaudeReviewFolder() & stem & "_run_" & runId & ".json"

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, resultJson
    Close #fileNum

    WriteRunResultReport = filePath
    Exit Function

Fail:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    WriteRunResultReport = ""
End Function

Private Sub StoreLastRunResult(ByVal runId As String, ByVal resultJson As String, ByVal reportPath As String)
    g_LastRunId = runId
    g_LastRunResultJson = resultJson
    g_LastRunReportPath = reportPath
End Sub

Private Function RunToolCallsInternal(ByVal jsonString As String, ByVal interactive As Boolean, _
    Optional ByVal automationMode As Boolean = False, Optional ByVal runId As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim toolCalls As Collection
    Dim errors As String
    Dim s As Object
    Dim startedAt As Date
    Dim startTimer As Single
    Dim mode As String
    Dim outcomes As New Collection
    Dim totalCount As Long
    Dim appliedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim resultJson As String
    Dim reportPath As String
    Dim nonInteractiveReport As String
    Dim restoreScreenUpdating As Boolean
    Dim performanceMode As Boolean

    startedAt = Now
    startTimer = Timer
    mode = IIf(interactive, "interactive", "apply_all")
    If Len(Trim$(runId)) = 0 Then runId = GenerateRunId()

    If Not ParseV4ToolCalls(jsonString, toolCalls, errors) Then
        reportPath = ""
        resultJson = BuildRunResultJson(runId, False, mode, 0, 0, 0, 0, errors, outcomes, startedAt, ElapsedMilliseconds(startTimer), reportPath)
        reportPath = WriteRunResultReport(runId, resultJson)
        resultJson = BuildRunResultJson(runId, False, mode, 0, 0, 0, 0, errors, outcomes, startedAt, ElapsedMilliseconds(startTimer), reportPath)
        Call StoreLastRunResult(runId, resultJson, reportPath)
        If Not automationMode Then
            MsgBox "V4 validation failed:" & vbCrLf & errors, vbCritical, "Invalid V4 Tool Calls"
        End If
        RunToolCallsInternal = False
        Exit Function
    End If

    ClearDocumentStructureMap
    If Not interactive Then
        restoreScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        performanceMode = True
    End If

    Dim suggestions As New Collection
    Dim callObj As Object
    For Each callObj In toolCalls
        suggestions.Add ConvertToolCallToSuggestion(callObj)
    Next callObj
    totalCount = suggestions.Count

    If interactive Then
        Dim idx As Long
        Dim action As String
        idx = 1

        Do While idx <= suggestions.Count
            Set s = suggestions(idx)
            Dim tRange As Range
            Set tRange = ResolveTargetToRange(GetSuggestionText(s, "target", ""))
            If Not tRange Is Nothing Then
                If IsToolCallNoOp(toolCalls(idx), tRange) Then
                    skippedCount = skippedCount + 1
                    AddOutcome outcomes, idx, s, "skipped", "NO_OP", "No change required; formatting/text already matches."
                    idx = idx + 1
                    GoTo NextLoop
                End If
            End If

            action = ShowSuggestionPreviewV4(s, idx, suggestions.Count)
            Select Case UCase$(action)
                Case "ACCEPT"
                    If ProcessSuggestionV4(s) Then
                        appliedCount = appliedCount + 1
                        AddOutcome outcomes, idx, s, "applied", "", ""
                    Else
                        failedCount = failedCount + 1
                        AddOutcome outcomes, idx, s, "failed", g_LastActionErrorCode, g_LastActionErrorMessage
                    End If
                Case "REJECT", "SKIP"
                    skippedCount = skippedCount + 1
                    AddOutcome outcomes, idx, s, "skipped", "USER_SKIPPED", "Skipped during interactive review."
                Case "ACCEPT_ALL"
                    Dim j As Long
                    For j = idx To suggestions.Count
                        If ProcessSuggestionV4(suggestions(j)) Then
                            appliedCount = appliedCount + 1
                            AddOutcome outcomes, j, suggestions(j), "applied", "", ""
                        Else
                            failedCount = failedCount + 1
                            AddOutcome outcomes, j, suggestions(j), "failed", g_LastActionErrorCode, g_LastActionErrorMessage
                        End If
                    Next j
                    idx = suggestions.Count + 1
                    GoTo NextLoop
                Case "STOP"
                    Dim k As Long
                    For k = idx To suggestions.Count
                        skippedCount = skippedCount + 1
                        AddOutcome outcomes, k, suggestions(k), "skipped", "USER_STOPPED", "Stopped interactive review."
                    Next k
                    Exit Do
                Case Else
                    skippedCount = skippedCount + 1
                    AddOutcome outcomes, idx, s, "skipped", "USER_SKIPPED", "Skipped during interactive review."
            End Select
            idx = idx + 1
NextLoop:
        Loop
    Else
        Dim failureDetails As String
        Dim failureLimit As Long
        Dim loggedFailures As Long
        Dim callIndex As Long

        failureLimit = 10
        loggedFailures = 0
        failureDetails = ""
        callIndex = 0

        ' --- NEW: PRE-FLIGHT RESOLUTION PASS ---
        ' Resolve all targets to Word Range objects BEFORE making any edits.
        ' This prevents the "Shifting Sand" problem where early edits misalign later targets.
        Dim preflightRange As Range
        Dim targetRefStr As String
        Dim preflightIndex As Long
        
        TraceLog "--- STARTING PRE-FLIGHT TARGET RESOLUTION ---"
        preflightIndex = 0
        For Each s In suggestions
            preflightIndex = preflightIndex + 1
            targetRefStr = GetSuggestionText(s, "target", "")
            Set preflightRange = ResolveTargetToRange(targetRefStr)
            
            If Not preflightRange Is Nothing Then
                ' Store the successfully anchored Range object into the JSON dictionary
                ' so ProcessSuggestionV4 can use it directly later.
                s.Add "pre_resolved_range", preflightRange
                TraceLog "  -> Pre-resolved: " & targetRefStr
            Else
                TraceLog "  -> FAILED to pre-resolve: " & targetRefStr
            End If
            If preflightIndex Mod APPLY_PREFLIGHT_PROGRESS_INTERVAL = 0 Then DoEvents
        Next s
        TraceLog "--- PRE-FLIGHT COMPLETE ---"

        ' --- ORIGINAL APPLY PASS ---
        For Each s In suggestions
            callIndex = callIndex + 1
            If ProcessSuggestionV4(s) Then
                appliedCount = appliedCount + 1
                AddOutcome outcomes, callIndex, s, "applied", "", ""
            Else
                failedCount = failedCount + 1

                ' Insert a comment at the target so the failed edit is visible in-document
                If HasDictionaryKey(s, "pre_resolved_range") Then
                    Dim failRange As Range
                    Set failRange = s("pre_resolved_range")
                    If Not failRange Is Nothing Then
                        Dim failNote As String
                        failNote = "[AI Review - Edit Not Applied]" & vbCrLf & _
                                   "Reason: " & g_LastActionErrorMessage & vbCrLf & _
                                   "Intended: " & Left$(GetSuggestionText(s, "replace", "(no replacement text)"), 500)

                        Dim failExplanation As String
                        failExplanation = GetSuggestionText(s, "explanation", "")
                        If Len(failExplanation) > 0 Then
                            failNote = failNote & vbCrLf & "Note: " & failExplanation
                        End If

                        On Error Resume Next
                        ExecuteCommentActionV4 failRange, failNote
                        On Error GoTo ErrorHandler
                    End If
                End If

                If loggedFailures < failureLimit Then
                    failureDetails = failureDetails & "  - " & DescribeToolFailure(s) & vbCrLf
                End If
                loggedFailures = loggedFailures + 1
                AddOutcome outcomes, callIndex, s, "failed", g_LastActionErrorCode, g_LastActionErrorMessage
            End If
        Next s

        nonInteractiveReport = "V4 Apply Complete" & vbCrLf & vbCrLf & _
                               "Applied: " & appliedCount & vbCrLf & _
                               "Failed: " & failedCount
        If failedCount > 0 Then
            If Len(failureDetails) > 0 Then
                nonInteractiveReport = nonInteractiveReport & vbCrLf & vbCrLf & "Failed entries:" & vbCrLf & failureDetails
            End If
            If loggedFailures > failureLimit Then
                nonInteractiveReport = nonInteractiveReport & "  ... and " & (loggedFailures - failureLimit) & " more failures." & vbCrLf
            End If
            nonInteractiveReport = nonInteractiveReport & vbCrLf & vbCrLf & _
                "Manual follow-up required:" & vbCrLf & _
                "- Review the failed tool_calls in your JSON input." & vbCrLf & _
                "- Use the DSM target IDs plus the tooling.manual_fallback guidance to locate each paragraph/table." & vbCrLf & _
                "- Apply the fixes manually in Word or adjust the JSON and rerun V4_ApplyToolCalls."
        End If
    End If

    RunToolCallsInternal = (failedCount = 0)

    reportPath = ""
    resultJson = BuildRunResultJson(runId, RunToolCallsInternal, mode, totalCount, appliedCount, skippedCount, failedCount, "", outcomes, startedAt, ElapsedMilliseconds(startTimer), reportPath)
    reportPath = WriteRunResultReport(runId, resultJson)
    resultJson = BuildRunResultJson(runId, RunToolCallsInternal, mode, totalCount, appliedCount, skippedCount, failedCount, "", outcomes, startedAt, ElapsedMilliseconds(startTimer), reportPath)
    Call StoreLastRunResult(runId, resultJson, reportPath)

    If Not automationMode Then
        If interactive Then
            MsgBox "V4 Interactive Review Complete" & vbCrLf & vbCrLf & _
                   "Applied: " & appliedCount & vbCrLf & _
                   "Skipped: " & skippedCount & vbCrLf & _
                   "Failed: " & failedCount, vbInformation, "V4 Complete"
        Else
            MsgBox nonInteractiveReport, vbInformation, "V4 Complete"
        End If
    End If

    If performanceMode Then
        On Error Resume Next
        Application.ScreenUpdating = restoreScreenUpdating
        Err.Clear
        On Error GoTo 0
    End If

    Exit Function

ErrorHandler:
    Dim errSuggestion As Object
    Dim runtimeErr As String
    runtimeErr = Err.Description
    If performanceMode Then
        On Error Resume Next
        Application.ScreenUpdating = restoreScreenUpdating
        Err.Clear
    End If
    On Error Resume Next
    Set errSuggestion = CreateObject("Scripting.Dictionary")
    errSuggestion("tool_name") = "<runtime>"
    errSuggestion("target") = "<runtime>"
    AddOutcome outcomes, 0, errSuggestion, "failed", "EXECUTION_ERROR", runtimeErr
    reportPath = ""
    resultJson = BuildRunResultJson(runId, False, mode, totalCount, appliedCount, skippedCount, failedCount + 1, runtimeErr, outcomes, startedAt, ElapsedMilliseconds(startTimer), reportPath)
    reportPath = WriteRunResultReport(runId, resultJson)
    resultJson = BuildRunResultJson(runId, False, mode, totalCount, appliedCount, skippedCount, failedCount + 1, runtimeErr, outcomes, startedAt, ElapsedMilliseconds(startTimer), reportPath)
    Call StoreLastRunResult(runId, resultJson, reportPath)
    If Not automationMode Then
        MsgBox "Error in V4 processing: " & runtimeErr, vbCritical, "V4 Error"
    End If
    RunToolCallsInternal = False
End Function

' =========================================================================================
' === PUBLIC ENTRY POINTS (KEEP THIS SECTION LAST) =======================================
' =========================================================================================

Public Sub V4_GenerateDocumentMap()
    On Error GoTo ErrorHandler

    ClearDocumentStructureMap
    BuildDocumentStructureMap ActiveDocument
    If g_DocumentMapCount = 0 Then
        MsgBox "Document is empty or map generation failed.", vbExclamation, "V4 Map"
        Exit Sub
    End If

    Dim mapJson As String
    mapJson = ExportStructureMapAsJsonV42()

    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText mapJson
    dataObj.PutInClipboard

    MsgBox "V4 document map copied to clipboard." & vbCrLf & _
           "Elements: " & g_DocumentMapCount, vbInformation, "V4 Map Ready"
    Exit Sub

ErrorHandler:
    MsgBox "Error generating V4 document map: " & Err.Description, vbCritical, "V4 Error"
End Sub

Public Sub V4_CopyDocumentMapToClipboard()
    V4_GenerateDocumentMap
End Sub

Public Sub V4_ValidateToolCallsJson()
    Dim inputForm As New frmJsonInput
    inputForm.Show vbModal
    If Len(Trim$(inputForm.txtJson.Value)) = 0 Then Exit Sub
    Call V4_ValidateToolCallsJsonText(inputForm.txtJson.Value, True)
End Sub

Public Function V4_ValidateToolCallsJsonText(ByVal jsonString As String, Optional ByVal showMessages As Boolean = False) As Boolean
    Dim toolCalls As Collection
    Dim errors As String
    V4_ValidateToolCallsJsonText = ParseV4ToolCalls(jsonString, toolCalls, errors)
    If showMessages Then
        If V4_ValidateToolCallsJsonText Then
            MsgBox "Valid V4 tool-call JSON." & vbCrLf & "Calls: " & toolCalls.Count, vbInformation, "Validation Successful"
        Else
            MsgBox "Invalid V4 tool-call JSON:" & vbCrLf & errors, vbCritical, "Validation Failed"
        End If
    End If
End Function

Public Function V4_ProcessToolCallsJson(ByVal jsonString As String, ByVal interactive As Boolean) As Boolean
    V4_ProcessToolCallsJson = RunToolCallsInternal(jsonString, interactive, True)
End Function

Public Function V4_ProcessToolCallsJsonEx(ByVal jsonString As String, Optional ByVal interactive As Boolean = False, Optional ByVal runId As String = "") As String
    Call RunToolCallsInternal(jsonString, interactive, True, runId)
    V4_ProcessToolCallsJsonEx = g_LastRunResultJson
End Function

Public Function V4_GetLastRunResultJson() As String
    V4_GetLastRunResultJson = g_LastRunResultJson
End Function

Public Function V4_GetLastRunReportPath() As String
    V4_GetLastRunReportPath = g_LastRunReportPath
End Function

Public Sub V4_ApplyToolCalls()
    Dim inputForm As New frmJsonInput
    inputForm.Show vbModal
    If Len(Trim$(inputForm.txtJson.Value)) = 0 Then Exit Sub
    Call RunToolCallsInternal(inputForm.txtJson.Value, False)
End Sub

Public Sub V4_RunInteractiveReview()
    Dim inputForm As New frmJsonInput
    inputForm.Show vbModal
    If Len(Trim$(inputForm.txtJson.Value)) = 0 Then Exit Sub
    Call RunToolCallsInternal(inputForm.txtJson.Value, True)
End Sub

Public Sub V4_TestTargetResolution()
    TestTargetResolution
End Sub

Public Sub V4_FinalCheckForRemainingChanges()
    FinalCheckForRemainingChanges
End Sub

' =========================================================================================
' === FILE-BASED ENTRY POINTS (for Claude Code integration) ===============================
' =========================================================================================
'
' These subs read/write JSON to a temp folder instead of clipboard/form,
' enabling automated review via Claude Code's report-checking skill.
'
' Exchange folder: %TEMP%\claude_review\
' Export files:    {doc_stem}_dsm.md / {doc_stem}_dsm.json
' Import file:     {doc_stem}_toolcalls.json  (LLM-generated tool calls)
' Backup file:     {doc_stem}_backup_YYYYMMDD_HHMMSS.docx (pre-review safety copy)

Private Function GetClaudeReviewFolder() As String
    ' Returns the temp exchange folder path, creating it if needed.
    Dim folderPath As String
    folderPath = Environ("TEMP") & "\claude_review"
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    GetClaudeReviewFolder = folderPath & "\"
End Function

Private Function GetDocStem() As String
    ' Returns the active document filename without extension.
    ' E.g. "6246.260211.NIA" from "6246.260211.NIA.docx"
    Dim docName As String
    docName = ActiveDocument.Name
    If InStrRev(docName, ".") > 0 Then
        GetDocStem = Left(docName, InStrRev(docName, ".") - 1)
    Else
        GetDocStem = docName
    End If
End Function

Private Function GetTimestampSlug() As String
    GetTimestampSlug = Format(Now, "yyyymmdd_hhnnss")
End Function

Public Function V4_ExportDocumentMapToFileEx(Optional ByVal copyToClipboard As Boolean = True, Optional ByVal showMessages As Boolean = False) As String
    ' Exports DSM markdown + JSON to %TEMP%\claude_review\ and returns the markdown path.
    On Error GoTo ErrorHandler

    If ActiveDocument.Path = "" Then
        If showMessages Then MsgBox "Please save the document first.", vbExclamation, "V4 Export"
        V4_ExportDocumentMapToFileEx = ""
        Exit Function
    End If

    ClearDocumentStructureMap
    BuildDocumentStructureMap ActiveDocument
    If g_DocumentMapCount = 0 Then
        If showMessages Then MsgBox "Document is empty or map generation failed.", vbExclamation, "V4 Export"
        V4_ExportDocumentMapToFileEx = ""
        Exit Function
    End If

    Dim markdown As String
    Dim mapJson As String
    markdown = ExportStructureMapAsMarkdown()
    mapJson = ExportStructureMapAsJsonV42()

    Dim markdownPath As String
    Dim jsonPath As String
    markdownPath = GetClaudeReviewFolder() & GetDocStem() & "_dsm.md"
    jsonPath = GetClaudeReviewFolder() & GetDocStem() & "_dsm.json"

    Dim fileNum As Integer

    fileNum = FreeFile
    Open markdownPath For Output As #fileNum
    Print #fileNum, markdown
    Close #fileNum

    fileNum = FreeFile
    Open jsonPath For Output As #fileNum
    Print #fileNum, mapJson
    Close #fileNum

    If copyToClipboard Then
        Dim dataObj As Object
        Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dataObj.SetText markdown
        dataObj.PutInClipboard
    End If

    V4_ExportDocumentMapToFileEx = markdownPath

    If showMessages Then
        MsgBox "V4 document map exported." & vbCrLf & _
               "Markdown: " & markdownPath & vbCrLf & _
               "JSON: " & jsonPath & vbCrLf & _
               "Elements: " & g_DocumentMapCount & vbCrLf & _
               IIf(copyToClipboard, "(Also copied to clipboard)", ""), vbInformation, "V4 Export Ready"
    End If
    Exit Function

ErrorHandler:
    If showMessages Then MsgBox "Error exporting V4 document map: " & Err.Description, vbCritical, "V4 Export Error"
    V4_ExportDocumentMapToFileEx = ""
End Function

Public Sub V4_ExportDocumentMapToFile()
    Call V4_ExportDocumentMapToFileEx(True, True)
End Sub

Public Function V4_ImportAndApplyToolCallsEx(Optional ByVal interactive As Boolean = False, _
    Optional ByVal runId As String = "", Optional ByVal showMessages As Boolean = False) As String
    ' Reads tool_calls JSON from %TEMP%\claude_review\{stem}_toolcalls.json and returns run result JSON.
    On Error GoTo ErrorHandler

    If ActiveDocument.Path = "" Then
        If showMessages Then MsgBox "Please save the document first.", vbExclamation, "V4 Import"
        V4_ImportAndApplyToolCallsEx = ""
        Exit Function
    End If

    Dim toolCallsPath As String
    toolCallsPath = GetClaudeReviewFolder() & GetDocStem() & "_toolcalls.json"
    If Dir(toolCallsPath) = "" Then
        If showMessages Then MsgBox "Tool calls file not found:" & vbCrLf & toolCallsPath, vbExclamation, "V4 Import"
        V4_ImportAndApplyToolCallsEx = ""
        Exit Function
    End If

    Dim backupPath As String
    backupPath = GetClaudeReviewFolder() & GetDocStem() & "_backup_" & GetTimestampSlug() & ".docx"
    ActiveDocument.Save ' Save current state first
    
    ' Safely copy the file using FileSystemObject to avoid Word/Cloud sync locks (Error 70)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile ActiveDocument.FullName, backupPath, True
    If Err.Number <> 0 Then
        Debug.Print "Warning: Could not create backup file. " & Err.Description
        Err.Clear ' Clear the error so it doesn't stop the import
    End If
    On Error GoTo ErrorHandler ' Restore standard error handling

    Dim fileNum As Integer
    Dim jsonString As String
    Dim fileLine As String
    fileNum = FreeFile
    Open toolCallsPath For Input As #fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        jsonString = jsonString & fileLine & vbCrLf
    Loop
    Close #fileNum

    Dim previousTrackState As Boolean
    previousTrackState = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = True

    If showMessages Then
        Dim msg As String
        msg = "Ready to apply tool calls from:" & vbCrLf & toolCallsPath & vbCrLf & vbCrLf & _
              "Backup saved to:" & vbCrLf & backupPath & vbCrLf & vbCrLf & _
              "Tracked changes: ON" & vbCrLf & vbCrLf & _
              "Run interactive review (Yes) or apply all (No)?"
        Dim result As VbMsgBoxResult
        result = MsgBox(msg, vbYesNoCancel + vbQuestion, "V4 Import & Apply")

        If result = vbCancel Then
            ActiveDocument.TrackRevisions = previousTrackState
            V4_ImportAndApplyToolCallsEx = ""
            Exit Function
        End If
        interactive = (result = vbYes)
    End If

    Call RunToolCallsInternal(jsonString, interactive, (Not showMessages), runId)
    V4_ImportAndApplyToolCallsEx = g_LastRunResultJson

    If showMessages Then
        MsgBox "Review complete. Tracked changes are ON." & vbCrLf & _
               "Backup at: " & backupPath, vbInformation, "V4 Import Done"
    End If
    Exit Function

ErrorHandler:
    If showMessages Then MsgBox "Error importing tool calls: " & Err.Description, vbCritical, "V4 Import Error"
    V4_ImportAndApplyToolCallsEx = ""
End Function

Public Sub V4_ImportAndApplyToolCalls()
    Call V4_ImportAndApplyToolCallsEx(False, "", True)
End Sub

Public Function V4_ExportDocumentMapSilent() As String
    ' COM-safe export: no message boxes, no clipboard, returns markdown path or empty string.
    V4_ExportDocumentMapSilent = V4_ExportDocumentMapToFileEx(copyToClipboard:=False, showMessages:=False)
End Function

Public Function V4_ImportAndApplyToolCallsSilent(Optional ByVal runId As String = "") As String
    ' COM-safe import+apply: no message boxes, non-interactive, returns run result JSON or empty string.
    V4_ImportAndApplyToolCallsSilent = V4_ImportAndApplyToolCallsEx(interactive:=False, runId:=runId, showMessages:=False)
End Function
