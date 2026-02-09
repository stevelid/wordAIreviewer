VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuggestionPreview 
   Caption         =   "Review AI Suggestion"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   OleObjectBlob   =   "frmSuggestionPreview.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   0  'Manual
Attribute VB_Name = "frmSuggestionPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    ' =========================================================================================
    ' === frmSuggestionPreview - Interactive Suggestion Review Form ==========================
    ' =========================================================================================
    Option Explicit

    Public UserAction As String ' "ACCEPT", "REJECT", "SKIP", "ACCEPT_ALL", "STOP"
    Public CurrentSuggestion As Object
    Public CurrentIndex As Long
    Public TotalCount As Long

    ' Private variables for UI state
    Private m_ContextRange As Range
    Private m_ActionRange As Range

Private Sub UserForm_Initialize()
    UserAction = ""
    Me.Width = 460 ' points
    Me.Height = 400 ' points
    ' Position form to the bottom-left to avoid obscuring the document center
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + 25
    Me.Top = Application.Top + Application.Height - Me.Height - 25
    
    ' Set Shortcut Keys (Accelerators)
    Me.cmdAccept.Accelerator = "A"      ' Alt+A
    Me.cmdReject.Accelerator = "R"      ' Alt+R
    Me.cmdSkip.Accelerator = "S"        ' Alt+S
    Me.cmdAcceptAll.Accelerator = "L"   ' Alt+L
    Me.cmdStop.Accelerator = "X"        ' Alt+X

    ' Lock display textboxes so they don't capture keyboard events
    Me.txtContext.Locked = True
    Me.txtTarget.Locked = True
    Me.txtReplace.Locked = True
    Me.txtExplanation.Locked = True
    
    ' Set background color for locked fields (optional, for visual clarity)
    Me.txtContext.BackColor = &H8000000F ' Light gray
    Me.txtTarget.BackColor = &H8000000F
    Me.txtReplace.BackColor = &H8000000F
    Me.txtExplanation.BackColor = &H8000000F
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Treat closing with the window X as a STOP action
        UserAction = "STOP"
        Me.Hide
        Cancel = True ' prevent destruction; ShowSuggestionPreview loop will exit
    End If
End Sub

' --- Public method to populate the form ---
Public Sub LoadSuggestion(ByVal suggestion As Object, ByVal index As Long, ByVal total As Long, _
                          ByVal contextRange As Range, ByVal actionRange As Range)
    Set CurrentSuggestion = suggestion
    CurrentIndex = index
    TotalCount = total
    Set m_ContextRange = contextRange
    Set m_ActionRange = actionRange
    
    ' Update progress label
    Me.lblProgress.Caption = "Suggestion " & index & " of " & total
    
    ' Populate the form fields
    Call PopulateFormFields
End Sub

' --- V4 method to populate the form with target references ---
Public Sub LoadSuggestionV4(ByVal suggestion As Object, ByVal index As Long, ByVal total As Long, _
                            ByVal contextRange As Range, ByVal actionRange As Range, _
                            Optional ByVal errorMessage As String = "")
    Set CurrentSuggestion = suggestion
    CurrentIndex = index
    TotalCount = total
    Set m_ContextRange = contextRange
    Set m_ActionRange = actionRange
    
    ' Update progress label
    Me.lblProgress.Caption = "Suggestion " & index & " of " & total
    
    ' Populate the form fields for V4
    Call PopulateFormFieldsV4(errorMessage)
End Sub

Private Sub PopulateFormFieldsV4(Optional ByVal errorMessage As String = "")
    Dim actionType As String
    Dim targetRef As String
    Dim findText As String
    Dim replaceText As String
    Dim explanation As String
    Dim styleKey As String
    
    ' Extract V4 fields from suggestion
    targetRef = GetSuggestionText(CurrentSuggestion, "target", "")
    actionType = GetSuggestionText(CurrentSuggestion, "action", "unknown")
    explanation = GetSuggestionText(CurrentSuggestion, "explanation", "")
    
    ' Build context display showing target reference
    Dim contextDisplay As String
    If Len(errorMessage) > 0 Then
        contextDisplay = "ERROR: " & errorMessage
    ElseIf m_ContextRange Is Nothing Then
        contextDisplay = "Target: " & targetRef & vbCrLf & "NOT FOUND - Target could not be resolved"
    Else
        contextDisplay = "Target: " & targetRef & vbCrLf & vbCrLf & _
                        "Preview: " & Left$(m_ContextRange.Text, 200)
        If Len(m_ContextRange.Text) > 200 Then
            contextDisplay = contextDisplay & "..."
        End If
    End If
    
    ' Build action-specific display
    Dim actionDisplay As String
    Select Case LCase$(actionType)
        Case "replace"
            findText = GetSuggestionText(CurrentSuggestion, "find", "")
            replaceText = GetSuggestionText(CurrentSuggestion, "replace", "")
            actionDisplay = "Find: " & findText & vbCrLf & "Replace: " & replaceText
            
        Case "apply_style"
            styleKey = GetSuggestionText(CurrentSuggestion, "style", "")
            actionDisplay = "Apply style: " & styleKey
            
        Case "comment"
            actionDisplay = "Add comment"
            
        Case "delete"
            actionDisplay = "Delete content"
            
        Case "replace_table"
            replaceText = GetSuggestionText(CurrentSuggestion, "replace", "")
            actionDisplay = "Replace table with:" & vbCrLf & Left$(replaceText, 100)
            If Len(replaceText) > 100 Then actionDisplay = actionDisplay & "..."
            
        Case "insert_row"
            actionDisplay = "Insert table row"
            
        Case "delete_row"
            actionDisplay = "Delete table row"
            
        Case Else
            actionDisplay = "(Action details)"
    End Select
    
    ' Set action type label with color coding
    Me.lblActionType.Caption = "Action: " & actionType
    Select Case LCase$(actionType)
        Case "replace"
            Me.lblActionType.ForeColor = RGB(0, 102, 204)
        Case "apply_style"
            Me.lblActionType.ForeColor = RGB(0, 153, 0)
        Case "replace_table"
            Me.lblActionType.ForeColor = RGB(153, 0, 0)
        Case "comment"
            Me.lblActionType.ForeColor = RGB(210, 140, 0)
        Case "delete", "delete_row"
            Me.lblActionType.ForeColor = RGB(204, 0, 0)
        Case "insert_row"
            Me.lblActionType.ForeColor = RGB(0, 102, 0)
        Case Else
            Me.lblActionType.ForeColor = RGB(0, 0, 0)
    End Select
    
    ' Populate form fields
    Me.txtContext.Text = TruncateText(contextDisplay, 500)
    Me.txtTarget.Text = TruncateText(targetRef, 200)
    Me.txtReplace.Text = TruncateText(actionDisplay, 600)
    Me.txtExplanation.Text = explanation
    
    ' Highlight and scroll to target
    If Not m_ContextRange Is Nothing Then
        m_ContextRange.HighlightColorIndex = wdYellow
        m_ContextRange.Select
        ActiveWindow.ScrollIntoView m_ContextRange, True
    Else
        ' No range to show - collapse selection
        Selection.Collapse Direction:=wdCollapseEnd
    End If
    
    ' Update button states
    Me.cmdAcceptAll.Enabled = (CurrentIndex < TotalCount)
    Me.cmdAccept.Enabled = Not (m_ContextRange Is Nothing)
End Sub

Private Sub PopulateFormFields()
    Dim actionType As String
    Dim context As String
    Dim target As String
    Dim replaceText As String
    Dim explanation As String
    
    ' Extract fields from suggestion
    context = GetSuggestionText(CurrentSuggestion, "context", "")
    
    ' Prefix with NOT FOUND if context was not located in document
    If m_ContextRange Is Nothing Then
        context = "NOT FOUND: " & context
    End If
    
    explanation = GetSuggestionText(CurrentSuggestion, "explanation", "")
    
    ' Handle compound vs simple actions
    If CurrentSuggestion.Exists("actions") Then
        actionType = "Multiple Actions"
        target = "(See details below)"
        replaceText = "(Multiple changes)"
        
        ' Build detailed explanation
        Dim subActions As Object
        Set subActions = CurrentSuggestion("actions")
        Dim i As Long
        Dim actionDetail As String
        actionDetail = explanation & vbCrLf & vbCrLf & "Actions:" & vbCrLf
        
        For i = 1 To subActions.Count
            Dim subAction As Object
            Set subAction = subActions(i)
            actionDetail = actionDetail & "  " & i & ". " & GetSuggestionText(subAction, "action", "unknown")
            If subAction.Exists("target") And Len(GetSuggestionText(subAction, "target", "")) > 0 Then
                actionDetail = actionDetail & " '" & GetSuggestionText(subAction, "target", "") & "'"
            End If
            If subAction.Exists("replace") And Len(GetSuggestionText(subAction, "replace", "")) > 0 Then
                actionDetail = actionDetail & " → '" & GetSuggestionText(subAction, "replace", "") & "'"
            End If
            actionDetail = actionDetail & vbCrLf
        Next i
        
        explanation = actionDetail
    Else
        actionType = GetSuggestionText(CurrentSuggestion, "action", "unknown")
        target = GetSuggestionText(CurrentSuggestion, "target", "")
        replaceText = GetSuggestionText(CurrentSuggestion, "replace", "")
    End If
    
    Me.lblActionType.Caption = "Action: " & actionType
    Dim atLower As String
    atLower = LCase$(actionType)
    Select Case atLower
        Case "change", "replace"
            Me.lblActionType.ForeColor = RGB(0, 102, 204)
        Case "apply_heading_style"
            Me.lblActionType.ForeColor = RGB(0, 153, 0)
        Case "replace_with_table"
            Me.lblActionType.ForeColor = RGB(153, 0, 0)
        Case "comment"
            Me.lblActionType.ForeColor = RGB(210, 140, 0)
        Case "multiple actions"
            Me.lblActionType.ForeColor = RGB(128, 0, 128)
        Case Else
    End Select

    Me.txtContext.Text = TruncateText(context, 500)
    Me.txtTarget.Text = TruncateText(target, 200)

    Me.txtReplace.Text = TruncateText(replaceText, 600)
    Me.txtExplanation.Text = explanation
    
    If Not m_ContextRange Is Nothing Then
        m_ContextRange.HighlightColorIndex = wdYellow
        If Not m_ActionRange Is Nothing Then
            m_ActionRange.HighlightColorIndex = wdBrightGreen
        End If
        m_ContextRange.Select
        ActiveWindow.ScrollIntoView m_ContextRange, True
    Else
        ' Ensure no text is confusingly selected from a previous step
        Selection.Collapse Direction:=wdCollapseEnd
    End If
    
    ' Update button states
    Me.cmdAcceptAll.Enabled = (CurrentIndex < TotalCount)
    Me.cmdAccept.Enabled = Not (m_ContextRange Is Nothing)
End Sub

Private Function TruncateText(ByVal text As String, ByVal maxLen As Long) As String
    If Len(text) <= maxLen Then
        TruncateText = text
    Else
        TruncateText = Left(text, maxLen) & "..."
    End If
End Function

Private Sub UserForm_Terminate()
    On Error Resume Next
    ClearHighlights
End Sub

Private Sub ClearHighlights()
    On Error Resume Next
    If Not m_ActionRange Is Nothing Then
        m_ActionRange.HighlightColorIndex = wdNoHighlight
    End If
    If Not m_ContextRange Is Nothing Then
        m_ContextRange.HighlightColorIndex = wdNoHighlight
    End If
End Sub

' --- Button Click Handlers ---
Private Sub cmdAccept_Click()
    UserAction = "ACCEPT"
    Me.Hide
End Sub

Private Sub cmdReject_Click()
    UserAction = "REJECT"
    Me.Hide
End Sub

Private Sub cmdSkip_Click()
    UserAction = "SKIP"
    Me.Hide
End Sub

Private Sub cmdAcceptAll_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Accept this and all remaining suggestions?" & vbCrLf & vbCrLf & _
                      "This will apply " & (TotalCount - CurrentIndex + 1) & " suggestions without further review.", _
                      vbQuestion + vbYesNo, "Accept All Remaining")
    If response = vbYes Then
        UserAction = "ACCEPT_ALL"
        Me.Hide
    End If
End Sub

Private Sub cmdStop_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Stop processing and exit?" & vbCrLf & vbCrLf & _
                      "Remaining suggestions will not be applied.", _
                      vbQuestion + vbYesNo, "Stop Processing")
    If response = vbYes Then
        UserAction = "STOP"
        Me.Hide
    End If
End Sub

' --- Keyboard Shortcuts ---
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 65 ' A key
            If Shift = 0 Then cmdAccept_Click
        Case 82 ' R key
            If Shift = 0 Then cmdReject_Click
        Case 83 ' S key
            If Shift = 0 Then cmdSkip_Click
        Case 78 ' N key (Next/Skip)
            If Shift = 0 Then cmdSkip_Click
        Case 27 ' ESC key
            cmdStop_Click
    End Select
End Sub

' --- Keyboard event handlers for textboxes (delegate to form) ---
Private Sub txtContext_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UserForm_KeyDown KeyCode, Shift
End Sub

Private Sub txtTarget_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UserForm_KeyDown KeyCode, Shift
End Sub

Private Sub txtReplace_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UserForm_KeyDown KeyCode, Shift
End Sub

Private Sub txtExplanation_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UserForm_KeyDown KeyCode, Shift
End Sub

' --- Helper function (duplicated from main module for form independence) ---
Private Function GetSuggestionText(ByVal suggestion As Object, ByVal key As String, Optional ByVal defaultText As String = "") As String
    On Error GoTo CleanFail
    If suggestion Is Nothing Then
        GetSuggestionText = defaultText
        Exit Function
    End If
    
    ' Check if key exists
    Dim hasKey As Boolean
    hasKey = False
    On Error Resume Next
    hasKey = suggestion.Exists(key)
    On Error GoTo CleanFail
    
    If hasKey Then
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
