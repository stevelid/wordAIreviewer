Attribute VB_Name = "Module07Attachments"
Sub PNS()
'
' Plant Noise Schedule
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\PNS.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_PNS.Show
End Sub

Sub AVM()
'
' Anti-Vibration Mount Schedule
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\AVM.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_AVM.Show
End Sub

Sub FCU()
'
' Fan Coil Unit Schedule
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\FCU.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_FCU.Show
End Sub


Sub RSS()
'
' Roomside Silencer Schedule
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\RSS.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_RSS.Show
End Sub

Sub ASS()
'
' Atmospheric Silencer Schedule
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\ASS.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_AAS.Show
End Sub

Sub PRS()
'
' Plantroom Structural Schedule
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\PRS.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_PRS.Show
End Sub

'Sub Glazing()
''
'' Glazing Schedule
''
'    Documents.Add Template:= _
'        AddinFolder & "\6. Attachments\glazing v1.0.dotx" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub

Sub Lifts()
'
' Lift
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Lifts.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_Lifts.Show
End Sub

Sub NewWHO()
'
' WHO Tables
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\WHO.dotx" _
        , NewTemplate:=False, DocumentType:=0

End Sub
Sub Newbb93()
'
' BB93 Tables
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\bb93.dotx" _
        , NewTemplate:=False, DocumentType:=0

End Sub

'Sub AttachBS8233Table()
''
'' Insert autotext
'
'Selection.Collapse Direction:=wdCollapseEnd
'    Selection.InsertFile FileName:=AddinFolder & "\6. Attachments\Attachments.docx", Range:="T8233", Link:=False
'
'End Sub
'
'
'Sub NEC()
''
'' NEC Tables
''
'    Documents.Add Template:= _
'        AddinFolder & "\6. Attachments\NEC tables.dotm" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub

Sub A3landscape()
'
' A3 Picture Landscape
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\A3 Figure.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_LandscapeDrawing.Show
End Sub

Sub Landscape()
'
' Landscape Figure
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Picture2.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_LandscapeDrawing.Show
End Sub

Sub Portrait()
'
' Portrait Figure
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Picture1.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_LandscapeDrawing.Show
End Sub

Sub Drawingissue()
'
' Drawing Issue Sheet
'
    Dim oExcel As Excel.Application
    Dim oWB As Excel.Workbook
    Set oExcel = New Excel.Application
    Set oWB = oExcel.Workbooks.Open(AddinFolder & "\6. Attachments\Drawing Issue.xltx")
    oExcel.visible = True

End Sub

'Sub PR()
''
'' PR List
''
'    Documents.Add Template:= _
'        AddinFolder & "\6. Attachments\PR List.dotx" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub

Sub AppendixA()
'
' Appendix A
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Appendix A.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_AppendixA.Show
End Sub

Sub AppendixFacer()
'
' Appendix A
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Appendix Facer.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Form6_AppendixFacer.Show
End Sub

Sub NewAppendix()
'
' New Appendix
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Appendix.dotm" _
        , NewTemplate:=False, DocumentType:=0

End Sub

Sub Surveysheet()
'
' Site Survey Sheet
'
    Documents.Add Template:= _
        AddinFolder & "\6. Attachments\VA Manual Survey Sheet.dotx" _
        , NewTemplate:=False, DocumentType:=0

End Sub



