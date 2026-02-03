    ' JSON parsing constants
    Private Const PARSE_SUCCESS As Long = 0
    Private Const PARSE_UNEXPECTED_END_OF_INPUT As Long = 1
    Private Const PARSE_INVALID_CHARACTER As Long = 2
    Private Const PARSE_INVALID_JSON_TYPE As Long = 3
    Private Const PARSE_INVALID_NUMBER As Long = 4
    Private Const PARSE_INVALID_KEY As Long = 5
    Private Const PARSE_INVALID_ESCAPE_CHARACTER As Long = 6

    ' Feature flags
    Private Const USE_GRANULAR_DIFF As Boolean = True

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

Option Explicit

' VBA Module to Apply LLM Review Suggestions (Version 3 - Production)
'
' REFINEMENTS:
' - Sets a distinct "AI Reviewer" identity for all changes.
' - Anchors comments to the full context for better visibility.
' - Pre-processes JSON to handle common errors (smart quotes, trailing commas).
' - UI now provides a progress indicator and validation.
' - Case sensitivity is now a user-configurable option.

' Always returns a new Scripting.Dictionary without requiring a reference

Private Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

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

        t = Trim$(NormalizeForDocument(p.Range.Text))

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

        t = Trim$(NormalizeForDocument(p.Range.Text))

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
    Dim row As row

    result = ""
    Set row = tbl.Rows(1)

    For Each c In row.Cells
        result = result & " " & Trim$(NormalizeForDocument(c.Range.Text))
    Next c

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
    Dim normalizedTitleStripped As String

    normalizedTitle = NormalizeForDocument(tableTitle)
    normalizedTitleStripped = NormalizeForDocument(StripTableNumberPrefix(tableTitle))

    ' First pass: exact match on title below
    For i = 1 To g_TableIndexCount
        If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).TitleBelow) Then
            matches.Add i
        ElseIf TextMatchesHeuristic(normalizedTitleStripped, StripTableNumberPrefix(g_TableIndex(i).TitleBelow)) Then
            matches.Add i
        End If
    Next i

    ' If no matches, try caption above
    If matches.Count = 0 Then
        For i = 1 To g_TableIndexCount
            If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).CaptionAbove) Then
                matches.Add i
            ElseIf TextMatchesHeuristic(normalizedTitleStripped, StripTableNumberPrefix(g_TableIndex(i).CaptionAbove)) Then
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
    Dim normalizedTitleStripped As String

    normalizedTitle = NormalizeForDocument(tableTitle)
    normalizedTitleStripped = NormalizeForDocument(StripTableNumberPrefix(tableTitle))

    ' Check title below
    For i = 1 To g_TableIndexCount
        If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).TitleBelow) Then
            matches.Add i
        ElseIf TextMatchesHeuristic(normalizedTitleStripped, StripTableNumberPrefix(g_TableIndex(i).TitleBelow)) Then
            matches.Add i
        End If
    Next i

    ' If no matches from title below, try caption above
    If matches.Count = 0 Then
        For i = 1 To g_TableIndexCount
            If TextMatchesHeuristic(normalizedTitle, g_TableIndex(i).CaptionAbove) Then
                matches.Add i
            ElseIf TextMatchesHeuristic(normalizedTitleStripped, StripTableNumberPrefix(g_TableIndex(i).CaptionAbove)) Then
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

' =========================================================================================
' === HOUSE STYLE INTEGRATION =============================================================
' =========================================================================================
' Customize this function to call your organization's style functions.
' Return True if the style was applied successfully, False to use fallback.

Private Function ApplyHouseStyle(ByVal rng As Range, ByVal styleName As String) As Boolean
    ' This function provides an extension point for your house style functions.
    ' Customize the Select Case below to call your actual style functions.
    '
    ' Example: If you have functions like ApplyH1(), ApplyH2(), etc., add them here.
    '
    ' To use:
    ' 1. Uncomment the relevant Case statements below
    ' 2. Replace the placeholder function calls with your actual function names
    ' 3. Ensure your style functions select the range before applying styles
    '
    ' Return True if the style was applied, False to fall back to default behavior.

    On Error GoTo Fail

    Dim normalizedStyle As String
    normalizedStyle = LCase$(Trim$(styleName))

    Select Case normalizedStyle
        ' ==========================================
        ' HEADING STYLES - Uncomment and customize:
        ' ==========================================

        'Case "heading 1", "h1"
        '    rng.Select
        '    Call ApplyH1  ' Replace with your actual function name
        '    ApplyHouseStyle = True
        '    Exit Function

        'Case "heading 2", "h2"
        '    rng.Select
        '    Call ApplyH2  ' Replace with your actual function name
        '    ApplyHouseStyle = True
        '    Exit Function

        'Case "heading 3", "h3"
        '    rng.Select
        '    Call ApplyH3  ' Replace with your actual function name
        '    ApplyHouseStyle = True
        '    Exit Function

        'Case "heading 4", "h4"
        '    rng.Select
        '    Call ApplyH4  ' Replace with your actual function name
        '    ApplyHouseStyle = True
        '    Exit Function

        ' ==========================================
        ' TABLE STYLES - Uncomment and customize:
        ' ==========================================

        'Case "table normal", "table"
        '    rng.Select
        '    Call ApplyTableStyle  ' Replace with your actual function name
        '    ApplyHouseStyle = True
        '    Exit Function

        ' ==========================================
        ' OTHER STYLES - Add more as needed:
        ' ==========================================

        'Case "body text", "normal"
        '    rng.Select
        '    Call ApplyBodyText  ' Replace with your actual function name
        '    ApplyHouseStyle = True
        '    Exit Function

        Case Else
            ' No matching house style function - use fallback
            ApplyHouseStyle = False
            Exit Function
    End Select

Fail:
    ' Error occurred - use fallback
    ApplyHouseStyle = False
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

    On Error Resume Next

    Dim field As field
    Dim fieldRange As Range

    ' Check all fields in the document for TOC fields
    For Each field In ActiveDocument.Fields
        If field.Type = wdFieldTOC Then
            Set fieldRange = field.Result

            ' Check if testRange overlaps with this TOC field
            If Not fieldRange Is Nothing Then
                If testRange.Start >= fieldRange.Start And testRange.Start < fieldRange.End Then
                    IsRangeInTOC = True
                    Debug.Print "    -> Skipping match: Range is within TOC at position " & fieldRange.Start
                    Exit Function
                End If
            End If
        End If
    Next field

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

Private Function ExtractAnchorTokens(ByVal segment As String) As Collection
    Dim normalized As String
    normalized = NormalizeForDocument(segment)

    Dim parts() As String
    Dim tokens As New Collection
    Dim p As Variant

    parts = Split(Trim$(normalized), " ")
    For Each p In parts
        Dim cleaned As String
        cleaned = CleanAnchorToken(CStr(p))

        If Len(cleaned) > 0 Then tokens.Add cleaned
    Next p

    Set ExtractAnchorTokens = tokens
End Function

Private Function CleanAnchorToken(ByVal token As String) As String
    Dim t As String
    t = Trim$(token)
    If Len(t) = 0 Then
        CleanAnchorToken = ""
        Exit Function
    End If

    Dim startIdx As Long
    Dim endIdx As Long
    startIdx = 1
    endIdx = Len(t)

    Do While startIdx <= endIdx
        Dim ch As String
        ch = Mid$(t, startIdx, 1)
        If ch Like "[A-Za-z0-9]" Then Exit Do
        startIdx = startIdx + 1
    Loop

    Do While endIdx >= startIdx
        ch = Mid$(t, endIdx, 1)
        If ch Like "[A-Za-z0-9]" Then Exit Do
        endIdx = endIdx - 1
    Loop

    If startIdx > endIdx Then
        CleanAnchorToken = ""
        Exit Function
    End If

    t = Mid$(t, startIdx, endIdx - startIdx + 1)

    Dim hasLetter As Boolean
    Dim hasDigit As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[A-Za-z]" Then hasLetter = True
        If ch Like "[0-9]" Then hasDigit = True
    Next i

    If Not hasLetter Then
        CleanAnchorToken = ""
        Exit Function
    End If

    CleanAnchorToken = t
End Function

Private Function SelectAnchorTokens(ByVal tokens As Collection, Optional ByVal maxTokens As Long = 8) As Collection
    Dim selected As New Collection
    If tokens Is Nothing Then
        Set SelectAnchorTokens = selected
        Exit Function
    End If

    If tokens.Count <= maxTokens Then
        Dim t As Variant
        For Each t In tokens
            selected.Add t
        Next t
        Set SelectAnchorTokens = selected
        Exit Function
    End If

    Dim stepSize As Double
    stepSize = (tokens.Count - 1) / (maxTokens - 1)

    Dim i As Long
    For i = 0 To maxTokens - 1
        Dim indexPos As Long
        indexPos = CLng(1 + (i * stepSize))
        If indexPos < 1 Then indexPos = 1
        If indexPos > tokens.Count Then indexPos = tokens.Count
        selected.Add tokens(indexPos)
    Next i

    Set SelectAnchorTokens = selected
End Function

Private Function FindTokenSequenceInRange(ByVal tokens As Collection, ByVal searchRange As Range, ByVal matchCase As Boolean, ByVal windowSize As Long) As Range
    On Error GoTo Fail
    If tokens Is Nothing Then
        Set FindTokenSequenceInRange = Nothing
        Exit Function
    End If
    If tokens.Count = 0 Then
        Set FindTokenSequenceInRange = Nothing
        Exit Function
    End If

    Dim firstToken As String
    firstToken = CStr(tokens(1))

    Dim searchCursor As Range
    Set searchCursor = searchRange.Duplicate

    Dim loopCounter As Long
    Dim lastPos As Long
    lastPos = -1

    Do While searchCursor.Start < searchCursor.End
        loopCounter = loopCounter + 1
        If loopCounter > LOOP_SAFETY_LIMIT Then Exit Do

        Dim firstFound As Range
        Set firstFound = FindLongString(firstToken, searchCursor, matchCase)
        If firstFound Is Nothing Then Exit Do

        If firstFound.Start = lastPos Then
            Exit Do
        End If
        lastPos = firstFound.Start

        Dim lastFound As Range
        Set lastFound = firstFound

        Dim success As Boolean
        success = True

        Dim i As Long
        For i = 2 To tokens.Count
            Dim look As Range
            Set look = searchRange.Duplicate
            look.Start = lastFound.End
            look.End = firstFound.Start + windowSize
            If look.End > searchRange.End Then look.End = searchRange.End
            If look.Start >= look.End Then
                success = False
                Exit For
            End If

            Dim found As Range
            Set found = FindLongString(CStr(tokens(i)), look, matchCase)
            If found Is Nothing Then
                success = False
                Exit For
            End If
            Set lastFound = found
        Next i

        If success Then
            Set FindTokenSequenceInRange = ActiveDocument.Range(firstFound.Start, lastFound.End)
            Exit Function
        End If

        searchCursor.Start = firstFound.End + 1
        If searchCursor.Start >= searchCursor.End Then Exit Do
    Loop

Fail:
    Set FindTokenSequenceInRange = Nothing
End Function

Private Function CalculateTokenWindowSize(ByVal searchString As String) As Long
    Dim size As Long
    size = CLng(Len(searchString) * 1.5)
    If size < 400 Then size = 400
    If size > 8000 Then size = 8000
    CalculateTokenWindowSize = size
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
Sub ApplyLlmReview_V3()
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

Sub RunReviewProcess(ByVal TheForm As frmJsonInput)
    ' *** WORKFLOW SELECTOR ***
    ' Set to True for INTERACTIVE mode (new), False for TRACKED CHANGES mode (old)
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

Sub RunInteractiveReview(ByVal suggestions As Object, ByVal matchCase As Boolean, ByVal startTime As Single)
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

Sub RunTrackedChangesReview(ByVal TheForm As frmJsonInput, ByVal suggestions As Object, ByVal startTime As Single)
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

Public Sub HandleError(ByVal procedureName As String, ByVal errSource As ErrObject)
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

' === CORE LOGIC AND HELPERS ==============================================================

' =========================================================================================

Private Function HasDictionaryKey(ByVal dict As Object, ByVal keyName As String) As Boolean
    On Error GoTo CleanFail
    HasDictionaryKey = False
    If dict Is Nothing Then Exit Function
    If TypeName(dict) = "Collection" Then Exit Function
    If TypeName(dict) = "Dictionary" Or TypeName(dict) = "Scripting.Dictionary" Then
        If dict.Exists(keyName) Then HasDictionaryKey = True
        Exit Function
    End If
    If dict.Exists(keyName) Then HasDictionaryKey = True
    Exit Function
CleanFail:
    HasDictionaryKey = False
End Function

Private Function GetSuggestionContextText(ByVal suggestion As Object) As String
    On Error GoTo CleanFail
    GetSuggestionContextText = "<missing context>"
    If suggestion Is Nothing Then Exit Function
    If TypeName(suggestion) = "Collection" Then Exit Function
    If HasDictionaryKey(suggestion, "context") Then
        Dim ctxValue As Variant
        ctxValue = suggestion("context")
        If IsNull(ctxValue) Then
            GetSuggestionContextText = "<missing context>"
        Else
            GetSuggestionContextText = CStr(ctxValue)
        End If
    End If
    Exit Function
CleanFail:
    GetSuggestionContextText = "<context error>"
End Function

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

Private Function GetSuggestionText(ByVal suggestion As Object, ByVal key As String, Optional ByVal defaultText As String = "") As String
    ' Safely extracts a text value from a dictionary suggestion.
    On Error GoTo CleanFail
    If HasDictionaryKey(suggestion, key) Then
        Dim val As Variant
        val = suggestion(key)
        If IsNull(val) Then
            GetSuggestionText = defaultText
        Else
            GetSuggestionText = CStr(val)
        End If
    Else
        GetSuggestionText = defaultText
    End If
    Exit Function
CleanFail:
    GetSuggestionText = defaultText
End Function

Private Function GetFirstSignificantWord(ByVal text As String) As String
    ' Extracts the first word longer than 4 chars, or the very first word if none are long enough.
    On Error Resume Next
    Dim words() As String
    Dim word As Variant
    Dim cleanText As String

    ' Simple cleanup to handle punctuation
    cleanText = Replace(text, ",", " ")
    cleanText = Replace(cleanText, ".", " ")
    cleanText = Replace(cleanText, ";", " ")
    cleanText = Replace(cleanText, ":", " ")
    cleanText = Replace(cleanText, "'", " ")
    cleanText = Replace(cleanText, """", " ")

    words = Split(Trim(cleanText), " ")

    ' Find first word > 4 chars
    For Each word In words
        If Len(word) > 4 Then
            GetFirstSignificantWord = word
            Exit Function
        End If
    Next word

    ' If no word > 4 chars was found, return the first non-empty word
    For Each word In words
        If Len(word) > 0 Then
            GetFirstSignificantWord = word
            Exit Function
        End If
    Next word

    GetFirstSignificantWord = "" ' Return empty if no words found
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
    Do While EndIndex > 0 And Mid$(p_Json, EndIndex - 1, 1) = "\"
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
    Debug.Print "  - New table created successfully with " & newTable.Rows.Count & " rows and " & newTable.Columns.Count & " columns."
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

    Dim cellRange As Range
    Set cellRange = c.Range

    ' Remove the cell end markers (Chr 13 & Chr 7)
    Dim txt As String
    txt = cellRange.Text

    ' Remove trailing cell marker
    If Right$(txt, 1) = Chr$(7) Then
        txt = Left$(txt, Len(txt) - 1)
    End If
    If Right$(txt, 1) = Chr$(13) Then
        txt = Left$(txt, Len(txt) - 1)
    End If

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

    On Error Resume Next

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

    If hasHintRow Then
        startRow = hintRow
        endRow = hintRow
    Else
        startRow = 1
        endRow = tbl.Rows.Count
    End If

    If hasHintCol Then
        startCol = hintCol
        endCol = hintCol
    Else
        startCol = 1
        endCol = tbl.Columns.Count
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
                On Error Resume Next
                Set currentCell = tbl.Cell(r, c)
                If Err.Number <> 0 Then
                    Err.Clear
                    GoTo NextCell
                End If
                On Error GoTo 0

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
                    On Error Resume Next
                    Dim aboveCell As Cell
                    Set aboveCell = tbl.Cell(r - 1, c)
                    If Err.Number = 0 Then
                        If Not TextMatchesHeuristic(aboveText, GetCellText(aboveCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If

                ' Check below
                If matches And Len(belowText) > 0 And r < tbl.Rows.Count Then
                    On Error Resume Next
                    Dim belowCell As Cell
                    Set belowCell = tbl.Cell(r + 1, c)
                    If Err.Number = 0 Then
                        If Not TextMatchesHeuristic(belowText, GetCellText(belowCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If

                ' Check left
                If matches And Len(leftText) > 0 And c > 1 Then
                    On Error Resume Next
                    Dim leftCell As Cell
                    Set leftCell = tbl.Cell(r, c - 1)
                    If Err.Number = 0 Then
                        If Not TextMatchesHeuristic(leftText, GetCellText(leftCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If

                ' Check right
                If matches And Len(rightText) > 0 And c < tbl.Columns.Count Then
                    On Error Resume Next
                    Dim rightCell As Cell
                    Set rightCell = tbl.Cell(r, c + 1)
                    If Err.Number = 0 Then
                        If Not TextMatchesHeuristic(rightText, GetCellText(rightCell)) Then
                            matches = False
                        End If
                    Else
                        matches = False
                    End If
                    Err.Clear
                    On Error GoTo 0
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
End Function

Private Function FindCellInTableRobust(ByVal tbl As Table, _
                                       ByVal rowHeader As String, _
                                       ByVal colHeader As String, _
                                       ByVal tableCellInfo As Object) As Range
    On Error Resume Next

    Dim foundColHeader As Range
    Dim finalCell As Range

    Dim targetRow As Row
    Dim hintRow As Long
    Dim hintCol As Long
    Dim hasHintRow As Boolean
    Dim hasHintCol As Boolean
    hasHintRow = False
    hasHintCol = False

    If Len(rowHeader) > 0 Then
        Dim r As Long
        For r = 1 To tbl.Rows.Count
            Dim rowFirstText As String
            rowFirstText = ""
            On Error Resume Next
            rowFirstText = GetCellText(tbl.Cell(r, 1))
            On Error GoTo 0
            If TextMatchesHeuristic(rowHeader, rowFirstText) Then
                Set targetRow = tbl.Rows(r)
                hintRow = r
                hasHintRow = True
                Exit For
            End If
        Next r
        If targetRow Is Nothing Then
            Debug.Print "    -> Row Header '" & rowHeader & "' not found in table."
            Set FindCellInTableRobust = Nothing
            Exit Function
        End If
    End If

    If Len(colHeader) > 0 Then
        Dim headerCell As Cell
        Dim headerRowScan As Long
        For headerRowScan = 1 To tbl.Rows.Count
            If headerRowScan > 3 Then Exit For
            Dim headerRow As Row
            Set headerRow = tbl.Rows(headerRowScan)
            For Each headerCell In headerRow.Cells
                Dim headerText As String
                headerText = GetCellText(headerCell)
                If TextMatchesHeuristic(colHeader, headerText) Then
                    Set foundColHeader = headerCell.Range
                    hintCol = headerCell.ColumnIndex
                    hasHintCol = True
                    Exit For
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
        Set adj = tableCellInfo("adjacentCells")
        Dim adjCell As Cell
        Set adjCell = FindCellByAdjacentContent(tbl, adj, hintRow, hintCol, hasHintRow, hasHintCol)
        If Not adjCell Is Nothing Then
            Set finalCell = adjCell.Range
        End If
    End If

    If finalCell Is Nothing And Not targetRow Is Nothing Then
        Dim cell As Cell
        Dim cellIdx As Long
        cellIdx = 0

        For Each cell In targetRow.Cells
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
        Next cell
    End If

    If finalCell Is Nothing And Not foundColHeader Is Nothing Then
        ' As a fallback, scan all rows for the aligned column header position
        Dim tblRow As Row
        For Each tblRow In tbl.Rows
            For Each cell In tblRow.Cells
                If PositionsRoughlyAlign(cell.Range, foundColHeader) Then
                    Set finalCell = cell.Range
                    Exit For
                End If
            Next cell
            If Not finalCell Is Nothing Then Exit For
        Next tblRow
    End If

    If Not finalCell Is Nothing Then
        If finalCell.End > finalCell.Start Then finalCell.End = finalCell.End - 1
        Set FindCellInTableRobust = finalCell
    Else
        Debug.Print "    -> Cell not found after row/column matching."
        Set FindCellInTableRobust = Nothing
    End If
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
    ' 0.5 Pipe-delimited table row sequence
    ' 1. Exact match (normalized)
    ' 2. Anchor token sequence (order-preserving, gap-tolerant)
    ' 3. Progressive shortening (90%, 75% of context from start)
    ' 4. Case-insensitive if case-sensitive was requested
    ' 5. Anchor word (disabled to avoid false positives)

    On Error GoTo ErrorHandler

    Dim result As Range
    Dim shortenedContext As String
    Dim searchLen As Long
    Dim cutoffPercent As Variant
    Dim cutoffPercentages As Variant
    Dim normalizedSearch As String
    normalizedSearch = NormalizeForDocument(searchString)

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
    Set result = FindLongString(normalizedSearch, searchRange, matchCase)
    If Not result Is Nothing Then
        Debug.Print "    -> SUCCESS: Found with exact match"
        Set FindWithProgressiveFallback = result
        Exit Function
    End If

    ' Strategy 2: Anchor token sequence (handles numbering, headings, and longer contexts)
    Dim anchorTokens As Collection
    Set anchorTokens = SelectAnchorTokens(ExtractAnchorTokens(normalizedSearch), 8)
    If Not anchorTokens Is Nothing Then
        If anchorTokens.Count >= 2 Then
            Dim windowSize As Long
            windowSize = CalculateTokenWindowSize(normalizedSearch)
            Debug.Print "  - Strategy 2: Anchor token sequence (" & anchorTokens.Count & " tokens, window " & windowSize & " chars)..."
            Set result = FindTokenSequenceInRange(anchorTokens, searchRange, matchCase, windowSize)
            If Not result Is Nothing Then
                Debug.Print "    -> SUCCESS: Found with anchor token sequence"
                Set FindWithProgressiveFallback = result
                Exit Function
            End If
        End If
    End If
    
    ' Strategy 3: Progressive context shortening (for overly specific contexts)
    ' Try 90%, 75% of the original context length from the START
    ' REDUCED from 80%, 60%, 40%, 25% to minimize false positives
    If Len(normalizedSearch) > 50 Then ' Only for longer contexts
        Debug.Print "  - Strategy 3: Progressive shortening..."
        cutoffPercentages = Array(0.9, 0.75)

        For Each cutoffPercent In cutoffPercentages
            searchLen = CLng(Len(normalizedSearch) * cutoffPercent)
            If searchLen < 40 Then Exit For ' Require at least 40 chars for reliability
            
            shortenedContext = Left$(normalizedSearch, searchLen)
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
    
    ' Strategy 4: Case-insensitive fallback (if original was case-sensitive)
    If matchCase Then
        Debug.Print "  - Strategy 4: Case-insensitive fallback..."
        Set result = FindLongString(normalizedSearch, searchRange, False)
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
        If Len(normalizedSearch) > 50 Then
            cutoffPercentages = Array(0.9)
            For Each cutoffPercent In cutoffPercentages
                searchLen = CLng(Len(normalizedSearch) * cutoffPercent)
                If searchLen < 40 Then Exit For

                shortenedContext = TrimToWordBoundary(Left$(normalizedSearch, searchLen))
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
    
    ' Strategy 5: Anchor word as last resort (DISABLED - high false positive risk)
    ' This strategy is disabled because matching a single word has extremely high
    ' false positive rates. It will almost always match the wrong location.
    ' If you need this fallback, the HandleNotFoundContext function will place
    ' a comment at a keyword location instead.
    Debug.Print "  - Strategy 5: Anchor word fallback SKIPPED (disabled to prevent false positives)"

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

    Debug.Print "    [ComputeDiff] oldText length: " & Len(oldText) & ", newText length: " & Len(newText)

    ' Handle edge cases
    If oldText = newText Then
        ' Identical - single EQUAL operation
        Set equalOp = NewDictionary()
        equalOp("Operation") = DIFF_EQUAL
        equalOp("Text") = oldText
        result.Add equalOp
        Debug.Print "    [ComputeDiff] Texts are identical."
        Set ComputeDiff = result
        Exit Function
    End If

    If Len(oldText) = 0 Then
        ' Only insertion
        Set insertOp = NewDictionary()
        insertOp("Operation") = DIFF_INSERT
        insertOp("Text") = newText
        result.Add insertOp
        Debug.Print "    [ComputeDiff] Old text empty, full insert."
        Set ComputeDiff = result
        Exit Function
    End If

    If Len(newText) = 0 Then
        ' Only deletion
        Set deleteOp = NewDictionary()
        deleteOp("Operation") = DIFF_DELETE
        deleteOp("Text") = oldText
        result.Add deleteOp
        Debug.Print "    [ComputeDiff] New text empty, full delete."
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
        Debug.Print "    [ComputeDiff] Common prefix length: " & prefixLen
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
        Debug.Print "    [ComputeDiff] Common suffix length: " & suffixLen
    End If

    ' Process the differing middle section
    If Len(oldMiddle) > 0 Then
        Set deleteOp = NewDictionary()
        deleteOp("Operation") = DIFF_DELETE
        deleteOp("Text") = oldMiddle
        result.Add deleteOp
        Debug.Print "    [ComputeDiff] Delete: '" & oldMiddle & "'"
    End If

    If Len(newMiddle) > 0 Then
        Set insertOp = NewDictionary()
        insertOp("Operation") = DIFF_INSERT
        insertOp("Text") = newMiddle
        result.Add insertOp
        Debug.Print "    [ComputeDiff] Insert: '" & newMiddle & "'"
    End If

    ' Add common suffix
    If suffixLen > 0 Then
        Set equalOp = NewDictionary()
        equalOp("Operation") = DIFF_EQUAL
        equalOp("Text") = suffixText
        result.Add equalOp
    End If

    Set ComputeDiff = result
    Debug.Print "    [ComputeDiff] Total operations: " & result.Count
End Function

' Determines whether to use granular diff or fall back to wholesale replacement
Private Function ShouldUseGranularDiff(ByVal oldText As String, _
                                       ByVal newText As String, _
                                       ByVal formatSegments As Collection) As Boolean
    ' Check global feature flag first
    If Not USE_GRANULAR_DIFF Then
        Debug.Print "    [ShouldUseGranularDiff] Feature disabled globally."
        ShouldUseGranularDiff = False
        Exit Function
    End If

    ' Rule 1: Text too long? Use fallback for performance
    If Len(oldText) > 1000 Or Len(newText) > 1000 Then
        Debug.Print "    [ShouldUseGranularDiff] Text too long, using fallback."
        ShouldUseGranularDiff = False
        Exit Function
    End If

    ' Rule 2: No formatting or simple formatting? Use granular
    If formatSegments Is Nothing Or formatSegments.Count = 0 Then
        Debug.Print "    [ShouldUseGranularDiff] No formatting, using granular diff."
        ShouldUseGranularDiff = True
        Exit Function
    End If

    ' Rule 3: Simple formatting (<=  3 segments)? Use granular
    If formatSegments.Count <= 3 Then
        Debug.Print "    [ShouldUseGranularDiff] Simple formatting (" & formatSegments.Count & " segments), using granular diff."
        ShouldUseGranularDiff = True
        Exit Function
    End If

    ' Rule 4: Complex formatting? Use fallback
    Debug.Print "    [ShouldUseGranularDiff] Complex formatting (" & formatSegments.Count & " segments), using fallback."
    ShouldUseGranularDiff = False
End Function

' Applies diff operations to a range with Track Changes enabled
Private Sub ApplyDiffOperations(ByVal targetRange As Range, _
                                ByVal diffOps As Collection, _
                                ByVal formatSegments As Collection)
    Debug.Print "    [ApplyDiffOperations] Applying " & diffOps.Count & " diff operations..."

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

        Debug.Print "    [ApplyDiffOperations] Op #" & opIndex & ": " & opType & " (length=" & opLen & ")"

        Select Case opType
            Case DIFF_EQUAL
                ' Skip over unchanged text
                currentPos = currentPos + opLen
                Debug.Print "    [ApplyDiffOperations] EQUAL: Skipping " & opLen & " chars, now at " & currentPos

            Case DIFF_DELETE
                ' Delete this text (will show as tracked deletion)
                If opLen > 0 Then
                    Dim delRange As Range
                    Set delRange = ActiveDocument.Range(currentPos, currentPos + opLen)
                    Debug.Print "    [ApplyDiffOperations] DELETE: Deleting range " & currentPos & " to " & (currentPos + opLen) & " ('" & Left$(opText, 20) & "')"
                    delRange.Delete
                    ' Note: currentPos stays same after deletion
                End If

            Case DIFF_INSERT
                ' Insert new text (will show as tracked insertion)
                If opLen > 0 Then
                    Dim insRange As Range
                    Set insRange = ActiveDocument.Range(currentPos, currentPos)
                    Debug.Print "    [ApplyDiffOperations] INSERT: Inserting at " & currentPos & " ('" & Left$(opText, 20) & "')"
                    insRange.Text = opText
                    currentPos = currentPos + opLen
                End If

            Case Else
                Debug.Print "    [ApplyDiffOperations] WARNING: Unknown operation type: " & opType
        End Select
    Next op

    ' Apply formatting after text changes are complete
    Debug.Print "    [ApplyDiffOperations] Text changes complete. Final position: " & currentPos

    ' Determine the new range after all modifications
    Dim finalRange As Range
    Set finalRange = ActiveDocument.Range(targetRange.Start, currentPos)

    ' Reset font and apply formatting segments
    finalRange.Font.Reset
    ApplyFormattingToSegments finalRange, formatSegments

    Debug.Print "    [ApplyDiffOperations] Diff operations complete."
End Sub

' Applies formatting segments to a range (extracted from old ApplyFormattedReplacement)
Private Sub ApplyFormattingToSegments(ByVal targetRange As Range, ByVal segments As Collection)
    If segments Is Nothing Then
        Debug.Print "    [ApplyFormattingToSegments] No segments to apply."
        Exit Sub
    End If

    If segments.Count = 0 Then
        Debug.Print "    [ApplyFormattingToSegments] Segments collection is empty."
        Exit Sub
    End If

    Dim baseStart As Long
    baseStart = targetRange.Start
    Dim segment As Object

    Debug.Print "    [ApplyFormattingToSegments] Applying " & segments.Count & " formatting segments..."

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

    Debug.Print "    [ApplyFormattingToSegments] Formatting applied."
End Sub

' Applies formatting when only formatting changes (no text changes)
Private Sub ApplyFormattingOnly(ByVal targetRange As Range, ByVal segments As Collection)
    Debug.Print "    [ApplyFormattingOnly] Text identical, applying formatting only..."

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

Public Sub UndoLlmReview()
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
Function PreProcessJson(ByVal jsonString As String) As String
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
Public Function LLM_ParseJson(ByVal jsonString As String) As Object
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
' === NormalizeForDocument (HELPER for line breaks) =======================================
' =========================================================================================

Private Function NormalizeForDocument(ByVal value As String) As String
    ' This function normalizes text for robust matching in Word documents.
    ' It handles:
    ' - Various newline characters (LF, CRLF) -> CR (Chr(13))
    ' - Special whitespace (non-breaking spaces, tabs) -> regular spaces
    ' - Multiple consecutive spaces -> single space
    ' - Smart quotes and special characters -> standard equivalents

    Dim result As String
    result = value

    If result = "" Then
        NormalizeForDocument = ""
        Exit Function
    End If

    ' 1. Handle line breaks - Replace the two-character CRLF first
    result = Replace(result, vbCrLf, Chr(13))
    result = Replace(result, vbLf, Chr(13))
    
    ' 2. Normalize special whitespace characters to regular spaces
    ' Non-breaking space (Chr(160))
    result = Replace(result, Chr(160), " ")
    ' Tab characters
    result = Replace(result, vbTab, " ")
    ' Vertical tab (Chr(11))
    result = Replace(result, Chr(11), " ")
    ' Form feed (Chr(12))
    result = Replace(result, Chr(12), " ")
    
    ' 3. Normalize smart quotes and special punctuation
    ' Left double quotation mark (Chr(8220))
    result = Replace(result, ChrW(8220), Chr(34))
    ' Right double quotation mark (Chr(8221))
    result = Replace(result, ChrW(8221), Chr(34))
    ' Left single quotation mark (Chr(8216))
    result = Replace(result, ChrW(8216), "'")
    ' Right single quotation mark (Chr(8217))
    result = Replace(result, ChrW(8217), "'")
    ' En dash (Chr(8211))
    result = Replace(result, ChrW(8211), "-")
    ' Em dash (Chr(8212))
    result = Replace(result, ChrW(8212), "-")
    
    ' 4. Collapse multiple consecutive spaces into a single space
    ' (but preserve paragraph breaks)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' 5. Trim leading/trailing spaces from each line (but not the whole string)
    result = NormalizeLinesWhitespace(result)
    
    NormalizeForDocument = result
End Function

Private Function NormalizeLinesWhitespace(ByVal text As String) As String
    ' Trims leading and trailing spaces from each line while preserving paragraph structure
    Dim lines() As String
    Dim i As Long
    Dim result As String
    
    If text = "" Then
        NormalizeLinesWhitespace = ""
        Exit Function
    End If
    
    lines = Split(text, Chr(13))
    result = ""
    
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then result = result & Chr(13)
        result = result & Trim(lines(i))
    Next i
    
    NormalizeLinesWhitespace = result
End Function

' =========================================================================================
' === UTILITY SUBROUTINE TO CHECK FOR REMAINING CHANGES ===================================
' =========================================================================================

Public Sub FinalCheckForRemainingChanges()
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
    If insertPosition = "before" Then
        Set newRow = tbl.Rows.Add(tbl.Rows(targetRowNum))
        ' newRow is now BEFORE targetRowNum (insertion shifts indices)
    Else ' "after"
        If targetRowNum < tbl.Rows.Count Then
            Set newRow = tbl.Rows.Add(tbl.Rows(targetRowNum + 1))
        Else
            ' Adding after last row
            Set newRow = tbl.Rows.Add
        End If
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
 
Sub StartAiReview()

    ' This sub launches the modeless review form.

    Dim reviewForm As New frmReviewer
    reviewForm.Show
End Sub
