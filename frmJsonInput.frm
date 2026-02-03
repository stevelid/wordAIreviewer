VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJsonInput 
   Caption         =   "Paste LLM Review JSON"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "frmJsonInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJsonInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code for frmJsonInput UserForm (Version 3)
Option Explicit

' Public properties to be read by the main module
Public JsonText As String
Public IsCancelled As Boolean
Public UseCaseSensitive As Boolean

' --- Event Handlers for Controls ---

Private Sub cmdProcess_Click()
    Debug.Print "User clicked: Process"
    ' User wants to run the review.
    If Trim(Me.txtJson.value) = "" Then
        MsgBox "JSON text cannot be empty.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    ' --- REFINEMENT: Update UI to show progress ---
    Me.lblProgress.Caption = "Initializing..."
    Me.txtJson.Enabled = False
    Me.chkCaseSensitive.Enabled = False
    Me.cmdValidate.Enabled = False
    Me.cmdProcess.Enabled = False
    DoEvents ' Allow the UI to update

    ' --- Call the main processing logic from the module ---
    RunReviewProcess Me
    
    ' When finished, close the form. The final report will be a MsgBox.
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Debug.Print "User clicked: Cancel"
    Me.IsCancelled = True
    Me.Hide
End Sub

Private Sub cmdValidate_Click()
    Debug.Print "User clicked: Validate"
    ' --- NEW: Validate JSON without running the full process ---
    Dim jsonString As String
    Dim suggestions As Object
    
    jsonString = Me.txtJson.value
    If Trim(jsonString) = "" Then
        MsgBox "Nothing to validate. The textbox is empty.", vbInformation, "Validation"
        Exit Sub
    End If
    
    ' Pre-process the string to clean it up
    jsonString = PreProcessJson(jsonString)
    
    ' Attempt to parse
    On Error GoTo ErrorHandler
    Set suggestions = LLM_ParseJson(jsonString)
    On Error GoTo 0 ' Turn off error handling if parse is successful
    
    If suggestions Is Nothing Or Not TypeName(suggestions) = "Collection" Then
        MsgBox "JSON Validation Failed!" & vbCrLf & vbCrLf & "The text is not a valid JSON array. Check for missing brackets, commas, or quotes.", vbCritical, "Validation Failed"
    Else
        MsgBox "JSON is valid! Found " & suggestions.Count & " items to process.", vbInformation, "Validation Successful"
    End If
    
    Exit Sub

ErrorHandler:
    HandleError "cmdValidate_Click", Err
End Sub

Private Sub UserForm_Activate()
    Debug.Print "Form activated, setting initial state."
    ' Set initial state when form loads
    Me.JsonText = ""
    Me.IsCancelled = False
    Me.UseCaseSensitive = False
    Me.txtJson.value = ""
    Me.lblProgress.Caption = "Ready"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Debug.Print "User clicked form's 'X' button (CloseMode: " & CloseMode & ")"
    ' Handle the user clicking the "X" button
    If CloseMode = vbFormControlMenu Then
        If Not Me.IsCancelled Then
            cmdCancel_Click
        End If
        Cancel = 1 ' Prevent form destruction, just hide it
    End If
End Sub


