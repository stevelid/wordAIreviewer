' =========================================================================================
' === DELETE SECTION FUNCTION (ADD TO wordAIreviewer.bas) ===============================
' =========================================================================================
'
' Add this function to wordAIreviewer.bas to support delete_section action
'

' Add this Case statement to ExecuteSingleAction (around line 1070, after "delete_table_row"):
'
'        Case "delete_section"
'            Debug.Print "Action 'delete_section': Removing section."
'            Call DeleteSection(actionObject, topLevelSuggestion, matchCase)


' =========================================================================================
' New Function: DeleteSection
' =========================================================================================
Private Sub DeleteSection(ByVal actionObject As Object, _
                         ByVal topLevelSuggestion As Object, _
                         ByVal matchCase As Boolean)
    ' Deletes an entire section from the document
    ' Finds heading by text (ignoring auto-numbering)
    ' Deletes from heading through end of section (next same/higher level heading)
    ' Supports disambiguation when multiple sections have same heading text

    On Error GoTo ErrorHandler

    ' Step 1: Get the heading text to search for
    Dim context As String
    context = GetSuggestionContextText(topLevelSuggestion)
    Dim contextForSearch As String
    contextForSearch = NormalizeForDocument(context)

    Debug.Print "  - Searching for section heading: '" & contextForSearch & "'"

    ' Step 2: Get optional heading level (if specified)
    Dim specifiedLevel As Long
    specifiedLevel = 0 ' 0 means auto-detect
    If HasDictionaryKey(actionObject, "headingLevel") Then
        On Error Resume Next
        specifiedLevel = CLng(actionObject("headingLevel"))
        On Error GoTo ErrorHandler
        If specifiedLevel < 1 Or specifiedLevel > 9 Then
            Err.Raise vbObjectError + 526, "DeleteSection", "headingLevel must be between 1 and 9"
        End If
        Debug.Print "  - Specified heading level: " & specifiedLevel
    End If

    ' Step 3: Get optional adjacent section info for disambiguation
    Dim hasAdjacentInfo As Boolean
    Dim beforeHeading As String
    Dim afterHeading As String
    hasAdjacentInfo = False

    If HasDictionaryKey(actionObject, "adjacentSection") Then
        Dim adjacentObj As Object
        Set adjacentObj = actionObject("adjacentSection")

        If HasDictionaryKey(adjacentObj, "before") Then
            beforeHeading = NormalizeForDocument(CStr(adjacentObj("before")))
            hasAdjacentInfo = True
        End If

        If HasDictionaryKey(adjacentObj, "after") Then
            afterHeading = NormalizeForDocument(CStr(adjacentObj("after")))
            hasAdjacentInfo = True
        End If

        If hasAdjacentInfo Then
            Debug.Print "  - Using adjacent section info for disambiguation"
            If Len(beforeHeading) > 0 Then Debug.Print "    - Before: '" & beforeHeading & "'"
            If Len(afterHeading) > 0 Then Debug.Print "    - After: '" & afterHeading & "'"
        End If
    End If

    ' Step 4: Find all paragraphs that match the heading text
    Dim doc As Document
    Set doc = ActiveDocument

    Dim matchingHeadings As New Collection
    Dim para As Paragraph

    For Each para In doc.Paragraphs
        ' Check if this paragraph is a heading
        Dim paraStyle As String
        paraStyle = para.Style.NameLocal

        ' Check if it's a heading style (Heading 1-9)
        Dim isHeading As Boolean
        Dim headingLevel As Long
        isHeading = False
        headingLevel = 0

        If InStr(1, paraStyle, "Heading", vbTextCompare) > 0 Then
            ' Extract level number (e.g., "Heading 2" -> 2)
            Dim i As Long
            For i = 1 To 9
                If paraStyle = "Heading " & i Or paraStyle = "Heading" & i Then
                    isHeading = True
                    headingLevel = i
                    Exit For
                End If
            Next i
        End If

        ' Also check OutlineLevel property as fallback
        If Not isHeading Then
            On Error Resume Next
            Dim outlineLevel As Long
            outlineLevel = para.OutlineLevel
            If Err.Number = 0 And outlineLevel >= wdOutlineLevel1 And outlineLevel <= wdOutlineLevel9 Then
                isHeading = True
                headingLevel = outlineLevel - wdOutlineLevel1 + 1
            End If
            On Error GoTo ErrorHandler
        End If

        If isHeading Then
            ' Check if heading level matches (if specified)
            If specifiedLevel > 0 And headingLevel <> specifiedLevel Then
                GoTo NextPara
            End If

            ' Get heading text without numbering
            Dim headingText As String
            headingText = StripHeadingNumber(para.Range.Text)
            headingText = NormalizeForDocument(headingText)

            ' Check if text matches
            Dim textMatches As Boolean
            If matchCase Then
                textMatches = (InStr(1, headingText, contextForSearch, vbBinaryCompare) > 0)
            Else
                textMatches = (InStr(1, headingText, contextForSearch, vbTextCompare) > 0)
            End If

            If textMatches Then
                ' Store this heading with its level
                Dim headingInfo As Object
                Set headingInfo = NewDictionary()
                headingInfo.Add "paragraph", para
                headingInfo.Add "level", headingLevel
                headingInfo.Add "text", headingText
                matchingHeadings.Add headingInfo

                Debug.Print "  - Found matching heading: '" & headingText & "' (Level " & headingLevel & ")"
            End If
        End If

NextPara:
    Next para

    ' Step 5: Check if we found any matches
    If matchingHeadings.Count = 0 Then
        Err.Raise vbObjectError + 527, "DeleteSection", "Could not find section heading: '" & context & "'"
    End If

    ' Step 6: Disambiguate if multiple matches
    Dim targetHeading As Object
    Set targetHeading = Nothing

    If matchingHeadings.Count = 1 Then
        Set targetHeading = matchingHeadings(1)
    Else
        ' Multiple matches - use adjacentSection info
        If Not hasAdjacentInfo Then
            Err.Raise vbObjectError + 528, "DeleteSection", _
                "Found " & matchingHeadings.Count & " sections with heading '" & context & "'. " & _
                "Use 'adjacentSection' field to specify which one."
        End If

        ' Find the one that matches adjacency criteria
        Dim idx As Long
        For idx = 1 To matchingHeadings.Count
            Dim candidate As Object
            Set candidate = matchingHeadings(idx)

            Dim isCorrectOne As Boolean
            isCorrectOne = True

            ' Check "before" heading
            If Len(beforeHeading) > 0 Then
                Dim prevHeading As String
                prevHeading = GetPreviousHeadingText(candidate("paragraph"), candidate("level"))
                If InStr(1, NormalizeForDocument(prevHeading), beforeHeading, vbTextCompare) = 0 Then
                    isCorrectOne = False
                End If
            End If

            ' Check "after" heading
            If Len(afterHeading) > 0 Then
                Dim nextHeading As String
                nextHeading = GetNextHeadingText(candidate("paragraph"), candidate("level"))
                If InStr(1, NormalizeForDocument(nextHeading), afterHeading, vbTextCompare) = 0 Then
                    isCorrectOne = False
                End If
            End If

            If isCorrectOne Then
                Set targetHeading = candidate
                Debug.Print "  - Disambiguated to correct section using adjacent headings"
                Exit For
            End If
        Next idx

        If targetHeading Is Nothing Then
            Err.Raise vbObjectError + 529, "DeleteSection", _
                "Could not disambiguate section - adjacentSection criteria did not match any occurrence."
        End If
    End If

    ' Step 7: Find the end of this section
    Dim startPara As Paragraph
    Set startPara = targetHeading("paragraph")
    Dim sectionLevel As Long
    sectionLevel = targetHeading("level")

    Debug.Print "  - Deleting section at level " & sectionLevel

    ' Find next paragraph with same or higher level heading
    Dim endPara As Paragraph
    Set endPara = Nothing

    Dim currentPara As Paragraph
    Dim foundStart As Boolean
    foundStart = False

    For Each currentPara In doc.Paragraphs
        If foundStart Then
            ' Check if this is a heading of same or higher level
            Dim currentStyle As String
            currentStyle = currentPara.Style.NameLocal

            Dim currentIsHeading As Boolean
            Dim currentLevel As Long
            currentIsHeading = False

            If InStr(1, currentStyle, "Heading", vbTextCompare) > 0 Then
                For i = 1 To 9
                    If currentStyle = "Heading " & i Or currentStyle = "Heading" & i Then
                        currentIsHeading = True
                        currentLevel = i
                        Exit For
                    End If
                Next i
            End If

            ' Check OutlineLevel as fallback
            If Not currentIsHeading Then
                On Error Resume Next
                outlineLevel = currentPara.OutlineLevel
                If Err.Number = 0 And outlineLevel >= wdOutlineLevel1 And outlineLevel <= wdOutlineLevel9 Then
                    currentIsHeading = True
                    currentLevel = outlineLevel - wdOutlineLevel1 + 1
                End If
                On Error GoTo ErrorHandler
            End If

            If currentIsHeading And currentLevel <= sectionLevel Then
                ' Found next section at same or higher level
                Set endPara = currentPara
                Debug.Print "  - Section ends before: '" & StripHeadingNumber(currentPara.Range.Text) & "'"
                Exit For
            End If
        End If

        If currentPara.Range.Start = startPara.Range.Start Then
            foundStart = True
        End If
    Next currentPara

    ' Step 8: Delete the range
    Dim rangeToDelete As Range
    Set rangeToDelete = doc.Range(startPara.Range.Start, _
        IIf(endPara Is Nothing, doc.Content.End, endPara.Range.Start))

    Debug.Print "  - Deleting " & rangeToDelete.Paragraphs.Count & " paragraphs"

    rangeToDelete.Delete

    Debug.Print "  -> Section deleted successfully"
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "DeleteSection: " & Err.Source, Err.Description
End Sub


' =========================================================================================
' Helper Function: StripHeadingNumber
' =========================================================================================
Private Function StripHeadingNumber(ByVal headingText As String) As String
    ' Removes auto-numbering from heading text
    ' Examples:
    '   "1. Introduction" -> "Introduction"
    '   "8.2 Methodology" -> "Methodology"
    '   "Introduction" -> "Introduction" (unchanged)

    Dim result As String
    result = Trim$(headingText)

    ' Remove paragraph mark at end
    If Right$(result, 1) = vbCr Or Right$(result, 1) = vbLf Then
        result = Left$(result, Len(result) - 1)
    End If
    If Right$(result, 1) = vbCr Or Right$(result, 1) = vbLf Then
        result = Left$(result, Len(result) - 1)
    End If

    ' Pattern: starts with numbers, dots, spaces, tabs
    ' Remove leading number patterns like "1.", "8.2", "A.1.2", etc.
    Dim i As Long
    Dim foundNonNumber As Boolean
    foundNonNumber = False

    For i = 1 To Len(result)
        Dim ch As String
        ch = Mid$(result, i, 1)

        If ch >= "0" And ch <= "9" Then
            ' Number - keep looking
        ElseIf ch = "." Or ch = " " Or ch = vbTab Or (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            ' Could be separator or letter in outline numbering - keep looking
            If ch <> "." And ch <> " " And ch <> vbTab Then
                foundNonNumber = True
            End If
        Else
            ' Found actual text content
            result = Trim$(Mid$(result, i))
            Exit For
        End If

        ' If we found a letter followed by more separators, start fresh from next letter
        If foundNonNumber And (ch = "." Or ch = " " Or ch = vbTab) Then
            foundNonNumber = False
        End If
    Next i

    ' Final cleanup
    result = Trim$(result)

    StripHeadingNumber = result
End Function


' =========================================================================================
' Helper Function: GetPreviousHeadingText
' =========================================================================================
Private Function GetPreviousHeadingText(ByVal fromPara As Paragraph, ByVal maxLevel As Long) As String
    ' Returns the text of the previous heading (at same or higher level)
    ' Used for disambiguation

    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = fromPara.Range.Document

    Dim para As Paragraph
    Dim foundCurrent As Boolean
    foundCurrent = False

    ' Iterate backwards through paragraphs
    Dim paraIndex As Long
    For paraIndex = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(paraIndex)

        If foundCurrent Then
            ' Check if this is a heading
            Dim paraStyle As String
            paraStyle = para.Style.NameLocal

            If InStr(1, paraStyle, "Heading", vbTextCompare) > 0 Then
                Dim i As Long
                For i = 1 To maxLevel
                    If paraStyle = "Heading " & i Or paraStyle = "Heading" & i Then
                        GetPreviousHeadingText = StripHeadingNumber(para.Range.Text)
                        Exit Function
                    End If
                Next i
            End If
        End If

        If para.Range.Start = fromPara.Range.Start Then
            foundCurrent = True
        End If
    Next paraIndex

    GetPreviousHeadingText = ""
    Exit Function

ErrorHandler:
    GetPreviousHeadingText = ""
End Function


' =========================================================================================
' Helper Function: GetNextHeadingText
' =========================================================================================
Private Function GetNextHeadingText(ByVal fromPara As Paragraph, ByVal maxLevel As Long) As String
    ' Returns the text of the next heading (at same or higher level)
    ' Used for disambiguation

    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = fromPara.Range.Document

    Dim para As Paragraph
    Dim foundCurrent As Boolean
    foundCurrent = False

    For Each para In doc.Paragraphs
        If foundCurrent Then
            ' Check if this is a heading
            Dim paraStyle As String
            paraStyle = para.Style.NameLocal

            If InStr(1, paraStyle, "Heading", vbTextCompare) > 0 Then
                Dim i As Long
                For i = 1 To maxLevel
                    If paraStyle = "Heading " & i Or paraStyle = "Heading" & i Then
                        GetNextHeadingText = StripHeadingNumber(para.Range.Text)
                        Exit Function
                    End If
                Next i
            End If
        End If

        If para.Range.Start = fromPara.Range.Start Then
            foundCurrent = True
        End If
    Next para

    GetNextHeadingText = ""
    Exit Function

ErrorHandler:
    GetNextHeadingText = ""
End Function


' =========================================================================================
' INTEGRATION NOTES:
' =========================================================================================
'
' 1. Add the Case statement to ExecuteSingleAction around line 1070 (after "delete_table_row")
'
' 2. Add these four new functions to the end of the module:
'    - DeleteSection
'    - StripHeadingNumber
'    - GetPreviousHeadingText
'    - GetNextHeadingText
'
' 3. Update the action sorting in ProcessSuggestion (around line 716) to include:
'    Case "delete_section"
'        cOther.Add actionObject  ' Or add early in processing if you want sections deleted first
'
' 4. This action works with Track Changes enabled - section deletion will be tracked
'
' 5. The heading text search is case-insensitive by default (uses the suggestion's matchCase setting)
'
' 6. Auto-numbering is automatically ignored - the function looks for heading text only
'
' 7. For documents with duplicate section headings, use the adjacentSection field:
'    {
'      "action": "delete_section",
'      "context": "Methodology",
'      "adjacentSection": {
'        "before": "Background",
'        "after": "Results"
'      }
'    }
'
' =========================================================================================
