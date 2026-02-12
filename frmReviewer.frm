VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReviewer 
   Caption         =   "AI Review Tool"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "frmReviewer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================================
' === CODE FOR frmReviewer (Accept / Skip / Ignore workflow) ============================
' =========================================================================================

Option Explicit

' This range tracks the currently selected suggestion group.
Private m_CurrentRange As Range
' This will store the name of the user who ran the script.
Private m_ReviewerName As String


Private Sub UserForm_Initialize()
    ' Capture the current user's name when the form is launched.
    ' All subsequent searches will look for changes made by this specific user.
    m_ReviewerName = Application.UserName

    ' Position to the right side
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + Application.Width - Me.Width - 25
    Me.Top = Application.Top + 50

    ' Rename buttons for review workflow: Accept / Skip / Ignore / Close
    Me.btnAccept.Caption = "&Accept"
    Me.btnAccept.Accelerator = "A"    ' Alt+A
    Me.btnReject.Caption = "&Skip"
    Me.btnReject.Accelerator = "S"    ' Alt+S  - reject but save suggestion as comment
    Me.btnFindNext.Caption = "&Ignore"
    Me.btnFindNext.Accelerator = "I"  ' Alt+I  - do nothing, advance to next
    Me.btnClose.Accelerator = "C"     ' Alt+C

    If m_ReviewerName = "" Then
        MsgBox "Warning: Your user name is not set in Word's options. The review tool may not find any changes.", vbExclamation
    End If

    ' Start the search from the beginning and auto-navigate to the first suggestion.
    Set m_CurrentRange = ActiveDocument.Content
    m_CurrentRange.Collapse wdCollapseStart
    FindNextSuggestion
End Sub

' --- BUTTON CLICK HANDLERS ---

Private Sub btnAccept_Click()
    ' Accept: apply tracked changes, delete AI comments, advance.
    ProcessCurrentSuggestion True
    FindNextSuggestion
End Sub

Private Sub btnReject_Click()
    ' Skip: reject tracked changes but save the suggestion as a comment for later.
    SkipCurrentSuggestion
    FindNextSuggestion
End Sub

Private Sub btnFindNext_Click()
    ' Ignore: leave tracked changes untouched and advance to the next group.
    FindNextSuggestion
End Sub

Private Sub btnClose_Click()
    ' This just triggers the form's closing process.
    ' The actual final check happens in the QueryClose event.
    Unload Me
End Sub

' --- NEW: Automatically run the final check when the form is closed ---
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' This event fires whenever the form is about to close, for any reason.
    ' We call our standalone checking macro from here.
    V4_FinalCheckForRemainingChanges
End Sub


' --- CORE LOGIC ---

Private Sub FindNextSuggestion()
    Dim nextGroup As Range
    ' Find the next group of changes starting after the end of the last group.
    Set nextGroup = FindNextGroup(m_CurrentRange.End)
    
    If nextGroup Is Nothing Then
        MsgBox "No more suggestions by '" & m_ReviewerName & "' were found.", vbInformation
        ' Reset to the beginning for the next session
        Set m_CurrentRange = ActiveDocument.Content
        m_CurrentRange.Collapse wdCollapseStart
        Exit Sub
    End If
    
    ' Update our current range to the new group and select it
    Set m_CurrentRange = nextGroup
    m_CurrentRange.Select
    
    ' Ensure the selection is visible to the user
    ActiveWindow.ScrollIntoView m_CurrentRange, True
End Sub

Private Sub ProcessCurrentSuggestion(ByVal Accept As Boolean)
    If m_CurrentRange Is Nothing Then Exit Sub

    Dim i As Long
    
    ' Disable screen updates for a flicker-free, faster experience
    Application.ScreenUpdating = False
    On Error Resume Next ' In case an item is deleted out of sequence

    ' --- Process Revisions that overlap with the group ---
    For i = ActiveDocument.Revisions.Count To 1 Step -1
        With ActiveDocument.Revisions(i)
            If .Author = m_ReviewerName And RangesOverlap(.Range, m_CurrentRange) Then
                If Accept Then .Accept Else .Reject
            End If
        End With
    Next i

    ' --- Process Comments that overlap with the group ---
    For i = ActiveDocument.Comments.Count To 1 Step -1
        With ActiveDocument.Comments(i)
            If .Author = m_ReviewerName And RangesOverlap(.Range, m_CurrentRange) Then
                .Delete ' Comments are deleted whether you accept or reject
            End If
        End With
    Next i

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub SkipCurrentSuggestion()
    ' Reject tracked changes but preserve the suggestion as a comment so
    ' the user can revisit it later.
    If m_CurrentRange Is Nothing Then Exit Sub

    Dim i As Long
    Application.ScreenUpdating = False
    On Error Resume Next

    ' --- 1. Collect a description of every revision in the group ---
    Dim suggText As String
    suggText = ""

    For i = 1 To ActiveDocument.Revisions.Count
        With ActiveDocument.Revisions(i)
            If .Author = m_ReviewerName And RangesOverlap(.Range, m_CurrentRange) Then
                If Len(suggText) > 0 Then suggText = suggText & "; "
                Select Case .Type
                    Case wdRevisionInsert
                        suggText = suggText & "insert """ & Left$(.Range.Text, 120) & """"
                    Case wdRevisionDelete
                        suggText = suggText & "delete """ & Left$(.Range.Text, 120) & """"
                    Case Else
                        suggText = suggText & "change """ & Left$(.Range.Text, 120) & """"
                End Select
            End If
        End With
    Next i

    ' --- 2. Capture any existing AI comment text (rationale / notes) ---
    Dim noteText As String
    noteText = ""

    For i = 1 To ActiveDocument.Comments.Count
        With ActiveDocument.Comments(i)
            If .Author = m_ReviewerName And RangesOverlap(.Range, m_CurrentRange) Then
                If Len(noteText) > 0 Then noteText = noteText & "; "
                noteText = noteText & Left$(.Range.Text, 200)
            End If
        End With
    Next i

    ' --- 3. Note the anchor position before modifying anything ---
    Dim anchorStart As Long
    anchorStart = m_CurrentRange.Start

    ' --- 4. Reject all tracked revisions in the group ---
    For i = ActiveDocument.Revisions.Count To 1 Step -1
        With ActiveDocument.Revisions(i)
            If .Author = m_ReviewerName And RangesOverlap(.Range, m_CurrentRange) Then
                .Reject
            End If
        End With
    Next i

    ' --- 5. Delete existing AI comments in the group ---
    For i = ActiveDocument.Comments.Count To 1 Step -1
        With ActiveDocument.Comments(i)
            If .Author = m_ReviewerName And RangesOverlap(.Range, m_CurrentRange) Then
                .Delete
            End If
        End With
    Next i

    ' --- 6. Build and add a skip comment at the anchor position ---
    Dim commentBody As String
    commentBody = ""
    If Len(suggText) > 0 Then commentBody = "Suggested: " & suggText
    If Len(noteText) > 0 Then
        If Len(commentBody) > 0 Then commentBody = commentBody & vbCrLf
        commentBody = commentBody & "Note: " & noteText
    End If

    If Len(commentBody) > 0 Then
        ' Anchor to the word at the original change position
        Dim safeEnd As Long
        safeEnd = anchorStart + 1
        If safeEnd > ActiveDocument.Content.End Then safeEnd = ActiveDocument.Content.End
        Dim commentRange As Range
        Set commentRange = ActiveDocument.Range(anchorStart, safeEnd)
        commentRange.MoveEnd wdWord, 1

        ActiveDocument.Comments.Add Range:=commentRange, _
            Text:="[Skipped] " & commentBody
    End If

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub


' --- HELPER FUNCTIONS for Grouping and Finding ---

Private Function FindNextGroup(ByVal startPos As Long) As Range
    ' Finds the EARLIEST next revision or comment, then expands it to a full group.
    Dim bestPos As Long: bestPos = -1
    Dim candidateRange As Range
    Dim rev As Revision, com As Comment

    ' Find the earliest revision after the start position
    For Each rev In ActiveDocument.Revisions
        If rev.Author = m_ReviewerName And rev.Range.Start >= startPos Then
            If bestPos = -1 Or rev.Range.Start < bestPos Then
                bestPos = rev.Range.Start
                Set candidateRange = rev.Range.Duplicate
            End If
        End If
    Next rev

    ' Find the earliest comment after the start position (and see if it's earlier than the revision)
    For Each com In ActiveDocument.Comments
        If com.Author = m_ReviewerName And com.Range.Start >= startPos Then
            If bestPos = -1 Or com.Range.Start < bestPos Then
                bestPos = com.Range.Start
                Set candidateRange = com.Range.Duplicate
            End If
        End If
    Next com

    ' If we found a candidate, expand it to include its neighbors.
    If Not candidateRange Is Nothing Then
        Set FindNextGroup = ExpandToGroup(candidateRange)
    End If
End Function

Private Function ExpandToGroup(ByVal seedRange As Range) As Range
    ' Expands a given range to include all adjacent/overlapping revisions and comments by the same author.
    Dim groupRange As Range: Set groupRange = seedRange.Duplicate
    Dim hasChanged As Boolean
    Dim rev As Revision, com As Comment
    
    Do
        hasChanged = False
        ' Check all revisions to see if they should be part of the group
        For Each rev In ActiveDocument.Revisions
            If rev.Author = m_ReviewerName Then
                ' If a revision is adjacent or overlapping...
                If rev.Range.Start <= groupRange.End + 1 And rev.Range.End >= groupRange.Start - 1 Then
                    ' ...then expand our group range to include it.
                    If rev.Range.Start < groupRange.Start Then groupRange.Start = rev.Range.Start: hasChanged = True
                    If rev.Range.End > groupRange.End Then groupRange.End = rev.Range.End: hasChanged = True
                End If
            End If
        Next rev
        
        ' Do the same for all comments
        For Each com In ActiveDocument.Comments
            If com.Author = m_ReviewerName Then
                If com.Range.Start <= groupRange.End + 1 And com.Range.End >= groupRange.Start - 1 Then
                    If com.Range.Start < groupRange.Start Then groupRange.Start = com.Range.Start: hasChanged = True
                    If com.Range.End > groupRange.End Then groupRange.End = com.Range.End: hasChanged = True
                End If
            End If
        Next com
    Loop While hasChanged ' Keep looping until the group stops growing
    
    Set ExpandToGroup = groupRange
End Function

Private Function RangesOverlap(ByVal rangeA As Range, ByVal rangeB As Range) As Boolean
    ' Returns True if two ranges touch or overlap.
    If rangeA Is Nothing Or rangeB Is Nothing Then Exit Function
    RangesOverlap = (rangeA.Start <= rangeB.End) And (rangeA.End >= rangeB.Start)
End Function
