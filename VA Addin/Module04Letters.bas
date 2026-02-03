Attribute VB_Name = "Module04Letters"
Sub New1Page()
'
' New 1 Page Letter
'
'
    Documents.Add Template:= _
        AddinFolder & "\3. Letters\Letter 1 page.dotm" _
        , NewTemplate:=False, DocumentType:=0

    Form3_1PageLetter.Show
    
End Sub

Sub New2Page()
'
' New 2 Page Letter
'
'
    Documents.Add Template:= _
        AddinFolder & "\3. Letters\Letter 2 page.dotm" _
        , NewTemplate:=False, DocumentType:=0

    Form3_2PageLetter.Show

End Sub

'Sub Coverletter()
''
'' Cover Letter Macro
''
''
'    Documents.Add Template:= _
'        AddinFolder & "\3. Letters\cover letter.dotm" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub

Sub Instruction()
'
' Invoice Instruction Letter
'
    Documents.Add Template:= _
        AddinFolder & "\3. Letters\Invoice Instruction.dotm" _
        , NewTemplate:=False, DocumentType:=0

    Form3_InvoiceInstruction.Show

End Sub

'Sub Warranty()
''
'' Warranty Letter
''
''
'    Documents.Add Template:= _
'        AddinFolder & "\3. Letters\warranty.dotm" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub

Sub Checklist()
'
' PCT Site Checksheet Macro
'
     Documents.Add Template:= _
        AddinFolder & "\6. Attachments\Site Readiness.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form6_PCTChecklist.Show
End Sub
'
'Sub Memo()
''
'' New Memo Macro
''
''
'    Documents.Add Template:= _
'        AddinFolder & "\3. Letters\memo cover.dotm" _
'        , NewTemplate:=False, DocumentType:=0
'
'End Sub

Sub DebtLetter1()
'
' PCT Site Checksheet Macro
'
     Documents.Add Template:= _
        AddinFolder & "\3. Letters\Debt Letter 1.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form3_DebtLetter1.Show
End Sub

Sub DebtLetter2()
'
' PCT Site Checksheet Macro
'
     Documents.Add Template:= _
        AddinFolder & "\3. Letters\Debt Letter 2.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form3_DebtLetter2.Show
End Sub

Sub DebtLetter3()
'
' PCT Site Checksheet Macro
'
     Documents.Add Template:= _
        AddinFolder & "\3. Letters\Debt Letter 3.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form3_DebtLetter3.Show
End Sub

Sub DebtLetter4()
'
' PCT Site Checksheet Macro
'
     Documents.Add Template:= _
        AddinFolder & "\3. Letters\Debt Letter 4.dotm" _
        , NewTemplate:=False, DocumentType:=0
        
    Form3_DebtLetter4.Show
End Sub
