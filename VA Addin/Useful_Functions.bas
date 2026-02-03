Attribute VB_Name = "Useful_Functions"
Option Explicit
'Define Logginglevels that may be used. When a level is selected all higher levels are also enabled.
Enum LoggingLevels
    NoLogging = 0
    ErrorLogging = 1
    WarningLogging = 2
    Debuglogging = 3
End Enum

Const LOGGING_LEVEL As Long = LoggingLevels.Debuglogging
Const JOB_DIR_IDENTIFIER As String = "Shared drives\Venta\Jobs" 'JobName and JobNum Functions searches for a job name after this string in the directory string

Function AddinFolder() As String
'Return the parent folder of this addin file ready to concatenate with subfolders.
    AddinFolder = ThisDocument.Path
End Function

Sub FillBMs(ByRef oBMPassed As Bookmark, strTextPassed As String) 'Bookmark and ListBox column data passed as parameters
Dim oRng As Word.Range
Dim strName As String
  Set oRng = oBMPassed.Range
  strName = oBMPassed.Name 'Get bookmark name
  oRng.text = strTextPassed 'Write ListBox column data to bookmark range (Note: This destroys the bookmark)
  ActiveDocument.Bookmarks.Add strName, oRng 'Recreate the bookmark spanning the range text
lbl_Exit:
  Exit Sub
End Sub

Function fcnCellText(ByRef oCell As Word.Cell) As String
  'Strip end of cell marker.
  fcnCellText = Left(oCell.Range.text, Len(oCell.Range.text) - 2)
lbl_Exit:
  Exit Function
End Function

Function GetConsultantDetails() As String()
' Generate an array of singature information from the signature source document

    Const CONSULTANT_DETAILS_DOC As String = "\Consultant Details.docx" 'file of doc with consultant details relative to the addin parent folder
    
    Dim sourcedoc As Document
    Dim i As Long, j As Long, m As Long, n As Long
    Dim strColWidths As String

    Application.ScreenUpdating = False
    'Define an array to be loaded with the data
    Dim arrData() As String
    'Open the file containing the table with items to load
    Set sourcedoc = Documents.Open(fileName:=AddinFolder & CONSULTANT_DETAILS_DOC, visible:=False)
    'Get the number members = number of rows in the table of details less one for the header row
    i = sourcedoc.Tables(1).Rows.Count - 1
    'Get the number of columns in the table of details
    j = sourcedoc.Tables(1).Columns.Count
    'Dimension arrData
    ReDim arrData(i - 1, j - 1)
    'Load table data into arrData
    For n = 0 To j - 1
      For m = 0 To i - 1
        arrData(m, n) = fcnCellText(sourcedoc.Tables(1).Cell(m + 2, n + 1))
      Next m
    Next n
    'Close the file containing the individual details
    sourcedoc.Close SaveChanges:=wdDoNotSaveChanges
    
    GetConsultantDetails = arrData
    Application.ScreenUpdating = True
End Function

Function GetColumnWidths(ByRef dataArray() As String) As String
'Build ColumnWidths statement for the signiture box

Dim n As Long
Dim strColWidths As String

    strColWidths = "50"
    For n = 2 To UBound(dataArray) + 1
        strColWidths = strColWidths + ";0"
    Next n
    GetColumnWidths = strColWidths

End Function

Sub SaveAsType(ByVal jobNumber As String, ByVal fileType As String)
' Saves documents with correct name and dating protocol
    
    Dim dateString As String
    Dim Projectdir As String
    Dim fileName As String
    Dim savename As String

    dateString = Format(CStr(Now), "yymmdd")
    Projectdir = GetJobDir(jobNumber)
    fileName = jobNumber & "." & dateString & "." & fileType
    
    If fileType Like "Q#" Then Projectdir = Projectdir & "\" & jobNumber & " Admin"
    
    'set document properties
    ActiveDocument.BuiltInDocumentProperties("Title") = "VA" & fileName
    ActiveDocument.BuiltInDocumentProperties("Author") = "Venta Acoustics"
    
    'save file in correct directory
    If Projectdir = "" Then
           savename = fileName
    Else
           savename = Projectdir & "\" & fileName
    End If
    With Dialogs(wdDialogFileSaveAs)
        .Name = savename
        .Show
    End With
    
    'update all fields in document
    Call updateFieldsIncludeHeadersFooters
End Sub

Sub MakeSitePlan(ByVal jobNumber As String, ByVal jobTitle As String)
' Opens Site Plan and autofills
    
    Documents.Add Template:=AddinFolder & "\6. Attachments\Picture2 - autofill.dotm"
    ActiveDocument.ActiveWindow.View = wdNormalView
    With ActiveDocument
        .Bookmarks("Project").Range.text = jobNumber & "/SP1"
        .Bookmarks("Title").Range.text = "Indicative Site Plan"
        .Bookmarks("Projtitle").Range.text = jobTitle
'        .Bookmarks("Date").Range.InsertDateTime DateTimeFormat:="dd MMMM yyyy"
    End With
    ActiveDocument.ActiveWindow.View = wdPrintView
    Application.ScreenUpdating = True
End Sub

Sub MakeComplaintForm(ByVal jobNumber As String, ByVal jobTitle As String)
' Opens Site Plan and autofills
    
    Documents.Add Template:=AddinFolder & "\6. Attachments\Appendix B - Noise Complaint Form.docx"
'    ActiveDocument.ActiveWindow.View = wdNormalView
'    With ActiveDocument
'        .Bookmarks("Project").Range.text = jobNumber
'        .Bookmarks("Title").Range.text = "Indicative Site Plan"
'        .Bookmarks("Projtitle").Range.text = jobTitle
''        .Bookmarks("Date").Range.InsertDateTime DateTimeFormat:="dd MMMM yyyy"
'    End With
'    ActiveDocument.ActiveWindow.View = wdPrintView
    Application.ScreenUpdating = True
End Sub

Function GetJobName(Optional Ref As Variant) As String
'Return the job name. Ref may be a job number, job directory or omitted.
'If omitted the directory of the calling worksheet or the active worksheet is passed.
'Returns empty string if no job name found
    
    Dim numberlength As Integer
    Dim jobString As String
    Dim leftIndex As Integer
    Dim rightindex As Integer
    Dim directory As String
    
    GetJobName = ""
    
    'If no arguments passed then check if function was called from a worksheet
    If IsMissing(Ref) Then
        On Error Resume Next
        directory = Application.ActiveDocument.Path
        If Err.Number <> 0 Then Exit Function 'no directory to work off
        On Error GoTo 0
    
    'If a directory passed then use the directory
    ElseIf Not IsNumeric(Ref) And InStr(1, Ref, "\") <> 0 Then
        directory = Ref
        
    ' If a job number is passed then find the relevent directory
    ElseIf CInt(Ref) > 1000 Then
        directory = GetJobDir(Ref)
    End If
    
    'If still no valid Directory then exit
    If Len(directory) < 34 Then
        Exit Function
    ElseIf GetJobNum(directory) = 0 Then
        Exit Function
    End If
    
    'find the first forward slash that occures after the project name:
    leftIndex = InStr(Val(InStr(1, directory, JOB_DIR_IDENTIFIER)) + Len(JOB_DIR_IDENTIFIER), directory, "\") + 1
    rightindex = InStr(leftIndex, directory, "\")
    If Not rightindex > 0 Then rightindex = Len(directory) + 1
    
    jobString = Mid(directory, leftIndex, InStr(leftIndex, directory, " ") - leftIndex)
    numberlength = Len(jobString)
    If Abs(numberlength - 5) <= 1 And IsNumeric(jobString) Then
        GetJobName = Mid(directory, leftIndex + numberlength + 1, rightindex - leftIndex - numberlength - 1)
    End If

End Function

Function GetJobNum(Optional directory As String = "Null") As Integer
'Return the Job Number of the folder holding the active sheet.
'Returns 0 if no job number found

    Dim jobString As String
    Dim leftIndex As Integer
    Dim rightindex As Integer
    
    GetJobNum = 0
    
    'If no Directory was passed then check if function was called from a worksheet
    If directory = "Null" Then
        On Error Resume Next
        directory = Application.ActiveDocument.Path
        If Err.Number <> 0 Then Exit Function 'no directory to work off
        On Error GoTo 0
    End If
    
    If Len(directory) > 34 Then
        'find the first forward slash that occures after the project name:
        leftIndex = InStr(Val(InStr(1, directory, JOB_DIR_IDENTIFIER)) + Len(JOB_DIR_IDENTIFIER), directory, "\") + 1
        rightindex = InStr(leftIndex, directory, " ")
        
        If leftIndex <> 0 And rightindex <> 0 Then
            jobString = Mid(directory, leftIndex, rightindex - leftIndex)
            If IsNumeric(jobString) Then GetJobNum = CInt(jobString)
        End If
    End If
End Function

Private Function GetJobDir(ByVal jobNum As Variant) As String
'accepts a jobnumber and returns base directory of that job

Dim directory As String
Dim tmpStr As String
Dim iJobNum As Integer

If IsNumeric(jobNum) Then
    If TypeName(jobNum) = "String" Then
        jobNum = Trim(jobNum) 'Remove any leading or training white spaces
    Else
        On Error GoTo getout
        jobNum = CStr(jobNum) 'Convert to string
    End If
    If Val(jobNum) < 1500 Or Val(jobNum) > 99999 Then
        Call LogError(jobNum & " is not a valid job number", "GetJobDir")
        Exit Function
    End If
Else
    'Not a job number
    Call LogError(jobNum & " is not a valid job number", "GetJobDir")
    Exit Function
End If

'Define base Directory
directory = GetBaseDir()

'Find a folder in this directory that starts with our JobNum
tmpStr = directory & jobNum & "*"
tmpStr = Dir(tmpStr, vbDirectory)

'If folder found then return
If tmpStr <> "" Then
    directory = directory & tmpStr & "\"
    GetJobDir = directory
    'Call LogDebug("GetJobDir(): Directory found: " & Directory , "GetJobDir")
Else
    Call LogError("Directory not found for job number: " & jobNum, "GetJobDir")
End If
getout:
End Function

Private Function GetBaseDir() As String
'Get the base directory for Venta job files. It is expected to be on the same drive as this add-in and
'have the form of JOB_DIR_IDENTIFIER

Dim errStr As String
Dim directory As String
Dim baseDir As String

'try in folder above this addin
directory = ThisDocument.Path
baseDir = Split(JOB_DIR_IDENTIFIER, "\")(0) & "\"

If InStr(directory, baseDir) <> 0 Then
    directory = Left(directory, InStrRev(directory, baseDir) - 1) & JOB_DIR_IDENTIFIER
    
    If Len(Dir(directory, vbDirectory)) <> 0 Then
        GetBaseDir = directory & "\"
        Exit Function
    Else
        Call LogWarning("directory not found.", "GetBaseDir")
    End If
    
Else
    Call LogWarning("Job folder not found in the same drive as this addin.", "GetBaseDir")
    
''try in the \Shared drives\Venta\ folder
'    directory = Environ("USERPROFILE") & "\Shared drives\Venta\"
'
'    If Len(Dir(directory, vbDirectory)) = 0 Then
'        errStr = "ERROR: odrive directory expected at " & directory & ". Not found."
'        Err.Raise vbObjectError + 513, "GetJobDir", errStr
'        Call LogError(errStr, "GetBaseDir", True)
'    End If
'
'    GetBaseDir = FindFolder(directory, JOB_DIR_IDENTIFIER)
End If

End Function

Private Function FindFolder(srcDir As String, folderName As String, Optional maxDepth As Long = 3, Optional curDepth As Long = 0) As String

    'Searches srcDir recursively (inc subfolders) for a folder with a path containing folderName and returns the path.
    'foldername may be a single folder or in the form 'folder/subfolder'
    
    Dim FSO As New FileSystemObject
    Dim myFolder As folder
    Dim mySubFolder As folder
    Dim tmpPath As String
    Dim i As Long
   
    If Len(srcDir) = 0 Then Exit Function
    curDepth = curDepth + 1
    If curDepth > maxDepth Then Exit Function
    
    FindFolder = ""
    If Right(srcDir, 1) <> "\" Then
        srcDir = srcDir & "\"
    End If
    
    Set myFolder = FSO.GetFolder(srcDir)
    
    'check if our folder exists in this directory
    tmpPath = srcDir & folderName & "\"
    If Len(Dir(tmpPath, vbDirectory)) <> 0 Then
    
        FindFolder = tmpPath
        Exit Function
    
    Else
        'recursive search of subfolders
        On Error Resume Next
        For Each mySubFolder In myFolder.SubFolders
            'Debug.Print mySubFolder.path
            FindFolder = FindFolder(mySubFolder.Path, folderName, maxDepth, curDepth)
            If FindFolder <> "" Then Exit Function
            DoEvents
        Next
        
    End If
End Function

Private Function FindFile(srcDir As String, fileName As String, Optional maxDepth As Long = 1000, Optional curDepth As Long = 0) As String

    'Searches srcDir recursively (inc subfolders) for a file with a name like fileName (wildcards allowed) and returns the full directory of the first file found.
    
    Dim FSO As New FileSystemObject
    Dim myFolder As folder
    Dim mySubFolder As folder
    Dim myFile As file
    Dim tmpPath As String
    Dim i As Long
   
    If Len(srcDir) = 0 Then Exit Function
    curDepth = curDepth + 1
    If curDepth > maxDepth Then Exit Function
    
    FindFile = ""
    If Right(srcDir, 1) <> "\" Then
        srcDir = srcDir & "\"
    End If
    
    Set myFolder = FSO.GetFolder(srcDir)
    
    'look for our file in this directory
    For Each myFile In myFolder.Files
        'Debug.Print myFile.Name
        If myFile.Name Like fileName Then
            FindFile = myFile.Path
            Exit Function
        End If
        DoEvents
    
    Next
    'recursive search of subfolders
    On Error Resume Next
    For Each mySubFolder In myFolder.SubFolders
        'Debug.Print mySubFolder.Path
        FindFile = FindFile(mySubFolder.Path, fileName, maxDepth, curDepth)
        If FindFile <> "" Then Exit Function
        DoEvents
    Next
        
End Function
Sub updateFieldsIncludeHeadersFooters()
'update all fields, including those in the headers and footers.

    Dim sec As Section
    Dim hdrftr As HeaderFooter

    ActiveDocument.Fields.Update 'address the fields in the main text story

    'now go through headers/footers for each section, update fields per range
    For Each sec In ActiveDocument.Sections
        For Each hdrftr In sec.Headers
            hdrftr.Range.Fields.Update
        Next
        For Each hdrftr In sec.Footers
            hdrftr.Range.Fields.Update
        Next
    Next
End Sub


Public Sub LogDebug(Message As String, Optional From As String = "")
'Log message when logging level is 1 or higher
    If LOGGING_LEVEL >= LoggingLevels.Debuglogging Then
        Debug.Print Time & " DEBUG - " & From & ": " & Message
    End If
End Sub

Public Sub LogWarning(Message As String, Optional From As String = "")
'Log warning when logging level is 1 or higher
    If LOGGING_LEVEL >= LoggingLevels.WarningLogging Then
        Debug.Print Time & " WARNING - " & From & ": " & Message
    End If
End Sub

Public Sub LogError(Message As String, Optional From As String = "", Optional GenerateMsgBox As Boolean = False)
'Log warning when logging level is 1 or higher.
'The error message can be displayed to the user in a message box by setting GenerateMsgBox to TRUE
    If LOGGING_LEVEL >= LoggingLevels.ErrorLogging Then
        Debug.Print Time & " ERROR - " & From & ": " & Message
        If GenerateMsgBox Then
            MsgBox Message, vbCritical, "Error!"
        End If
    End If
    
End Sub

Public Function CountCharacters(ByVal text As String, ByVal ch As String) As Long
  Dim cnt As Long
  Dim i As Long
  cnt = 0
  
  i = InStr(1, text, ch)
  Do While i > 1
    cnt = cnt + 1
    i = InStr(i + Len(ch), text, ch)
  Loop
  
  CountCharacters = cnt
End Function

Function getClientAddressFrmQuote(ByVal jobNum As String) As String
    'Return the client name and address from the quote in the jobfolder.
    Dim quotefileFormat As String
    Dim quotePath As String
    Dim quoteAlreadyOpen As Boolean
    Dim quoteDoc As Document
    Dim addressStart As Long
    Dim addressEnd As Long
    Dim addressText As String
    
    'Find a quote
    quotefileFormat = jobNum & ".*.Q#.doc*"
    quotePath = FindFile(GetJobDir(jobNum), quotefileFormat)
    
    'Check if quote is already open
    On Error Resume Next
    Set quoteDoc = GetObject(quotePath)
    If Err.Number <> 0 Then
        'open quote
        quoteAlreadyOpen = False
        Set quoteDoc = Documents.Open(fileName:=quotePath, ReadOnly:=True, _
                                    AddToRecentFiles:=False, visible:=False)
    Else
        quoteAlreadyOpen = True
    End If
    On Error GoTo ErrorHandler
    
    'get address
    With quoteDoc
        addressStart = .Bookmarks("address").Start
        addressEnd = .Bookmarks("name").End - Len("Dear ") - 2 'remove trailing paragraph markers
        addressText = .Range(addressStart, addressEnd).text
    End With
    
    'Close the quote and release the object
    If Not quoteAlreadyOpen Then
        quoteDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    Set quoteDoc = Nothing
    
    getClientAddressFrmQuote = addressText
    Exit Function
    
ErrorHandler:
    ' Handle errors here
    getClientAddressFrmQuote = "Error: " & Err.description
    ' Ensure document is closed and object is released
    If Not quoteDoc Is Nothing Then
        If Not quoteAlreadyOpen Then
            quoteDoc.Close SaveChanges:=wdDoNotSaveChanges
        End If
        Set quoteDoc = Nothing
    End If
End Function


Sub OpenCMPCalc()
    
    Dim oExcel As Excel.Application
    Dim oWB As Workbook
    Set oExcel = New Excel.Application
    Set oWB = oExcel.Workbooks.Open(fileName:=AddinFolder & "\Excel Templates\Calculations\CMP Calcs.xlsx")
    oExcel.visible = True
    
    
'    Dim dateString As String
    Dim Projectdir As String
    Dim fileName As String
    Dim savename As String

'    dateString = Format(CStr(Now), "yymmdd")
'    Projectdir = GetJobDir(jobNumber)
'    fileName = jobNumber & "Appendix B"
   
    'save file in correct directory
    If Projectdir = "" Then
           savename = fileName
    Else
           savename = Projectdir & "\" & fileName
    End If
    With Dialogs(wdDialogFileSaveAs)
        .Name = savename
        .Show
    End With

End Sub

Public Sub FormatAddressBox(addressBox As MSForms.TextBox)
    Dim currentText As String, formattedText As String
    Dim cursorPosition As Long
    Dim selectionLength As Long
    Dim previousText As String
    
    currentText = addressBox.text
    cursorPosition = addressBox.SelStart
    selectionLength = addressBox.SelLength
    previousText = addressBox.Tag
    
    ' Only format if a large chunk of text was added (likely a paste operation)
    If Len(currentText) > Len(previousText) + 5 Then
        formattedText = FormatAddress(currentText)
        
        If formattedText <> currentText Then
            addressBox.text = formattedText
            ' Attempt to maintain cursor position and selection
            If cursorPosition <= Len(formattedText) Then
                addressBox.SelStart = cursorPosition
                If cursorPosition + selectionLength <= Len(formattedText) Then
                    addressBox.SelLength = selectionLength
                End If
            Else
                addressBox.SelStart = Len(formattedText)
            End If
        End If
    Else
        ' For manual entry, just ensure consistent line breaks
        formattedText = Replace(currentText, vbNewLine, vbCrLf)
        If formattedText <> currentText Then
            addressBox.text = formattedText
            addressBox.SelStart = cursorPosition
        End If
    End If
    
    ' Store current text for next comparison
    addressBox.Tag = addressBox.text
End Sub

Public Function FormatAddress(text As String) As String
    Dim lines() As String, i As Long
    Dim result As String
    Dim separators As String
    Dim formattedLine As String
    
    separators = "[,;/|]" ' Common separators: comma, semicolon, forward slash, pipe
    
    ' Split the text by newlines
    lines = Split(text, vbNewLine)
    
    For i = LBound(lines) To UBound(lines)
        ' Format lines that contain multiple separators
        If CountSeparators(lines(i), separators) > 1 Then
            formattedLine = RegExReplace(lines(i), separators, vbNewLine)
            ' Trim each line to remove leading/trailing spaces
            formattedLine = TrimLines(formattedLine)
            result = result & formattedLine & vbCrLf
        Else
            ' Keep lines with 0 or 1 separator as they are, but trim
            result = result & Trim(lines(i)) & vbCrLf
        End If
    Next i
    
    ' Remove the last extra line break
    If Len(result) > 2 Then
        result = Left(result, Len(result) - 2)
    End If
    
    FormatAddress = result
End Function

' New helper function to trim each line in a multi-line string
Private Function TrimLines(text As String) As String
    Dim lines() As String, i As Long
    Dim result As String
    
    lines = Split(text, vbNewLine)
    For i = LBound(lines) To UBound(lines)
        result = result & Trim(lines(i)) & vbNewLine
    Next i
    
    ' Remove the last extra newline
    If Len(result) > 2 Then
        result = Left(result, Len(result) - 2)
    End If
    
    TrimLines = result
End Function

Private Function CountSeparators(text As String, separators As String) As Long
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Global = True
    regex.pattern = separators
    
    CountSeparators = regex.Execute(text).Count
End Function

Private Function RegExReplace(text As String, pattern As String, replacement As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .pattern = pattern
        RegExReplace = .Replace(text, replacement)
    End With
End Function

Sub SaveAsPDF2()
    Dim sName As String
    Dim sPath As String
    
    ActiveWindow.View.ShowRevisionsAndComments = False

    With ActiveDocument
        .Save
        sName = Left(.Name, InStr(.Name, ".doc") - 1)
        sName = sName & ".pdf"
        sPath = .Path & "\"

        .ExportAsFixedFormat _
          OutputFileName:=sPath & sName, _
          ExportFormat:=wdExportFormatPDF
    End With
End Sub

Sub CheckandPrintAll()

    Dim doc As Document
    Dim file As String
    Dim folder As String

    folder = ActiveDocument.Path & "\"
    file = Dir(folder & "*.doc*", vbNormal)

    While file <> ""

        Set doc = Documents.Open(fileName:=folder & file, ReadOnly:=True)
        Set doc = ActiveDocument

        Call CheckforReferenceErrors

        With ActiveWindow.View
            .ShowRevisionsAndComments = False
            .RevisionsView = wdRevisionsViewFinal
        End With

        Application.PrintOut fileName:="", Range:=wdPrintAllDocument, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        Collate:=True, Background:=True, PrintToFile:=False, PrintZoomColumn:=0, _
        PrintZoomRow:=0, PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0

        'doc.Save
        doc.Close
        file = Dir()
    Wend


End Sub
Sub CheckforReferenceErrors()

ActiveDocument.Fields.Update
With Selection
    .HomeKey Unit:=wdStory
    With .Find
        .ClearFormatting
        .text = "Error!"
        If .Execute Then
            MsgBox "Broken cross reference."
        End If
    End With
End With


End Sub


Sub UpdateLanguage()
    Dim originalStart As Long
    Dim originalEnd As Long
    
    ' Store the original selection
    originalStart = Selection.Start
    originalEnd = Selection.End
    
    ' Perform language update
    Selection.WholeStory
    With Selection
        .LanguageID = wdEnglishUK
        .NoProofing = False
    End With
    Application.CheckLanguage = False
    
    ' Update fields and set document properties
    ActiveDocument.Fields.Update
    With ActiveDocument
        .ShowGrammaticalErrors = False
        .ShowSpellingErrors = True
    End With
    
    ' Check for broken cross-references
    If DocumentHasBrokenReference Then
        MsgBox "Broken cross reference found."
    End If
    
    ' Restore original selection
    Selection.SetRange originalStart, originalEnd
End Sub

Function DocumentHasBrokenReference() As Boolean
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .text = "Error!"
        .Forward = True
        .Wrap = wdFindStop
        DocumentHasBrokenReference = .Execute
    End With
End Function

