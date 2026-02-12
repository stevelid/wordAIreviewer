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

    ' Processing is initiated by the calling entry point.
    ' This button now only confirms input and closes the form.
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
    Dim isValid As Boolean
    
    jsonString = Me.txtJson.value
    If Trim(jsonString) = "" Then
        MsgBox "Nothing to validate. The textbox is empty.", vbInformation, "Validation"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    isValid = V4_ValidateToolCallsJsonText(jsonString, True)
    On Error GoTo 0
    
    Exit Sub

ErrorHandler:
    MsgBox "Validation failed: " & Err.Description, vbCritical, "Validation Failed"
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


