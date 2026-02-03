Attribute VB_Name = "Module10ProcsRAMS"
Sub NAWRAMS()
'
' Noise at Work RAMS
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\Noise at Work RAMS.dotm" _
        , NewTemplate:=False, DocumentType:=0

    Form10_RAMS.Show

End Sub

Sub SITRAMS()
'
' Noise at Work RAMS
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\SIT RAMS.dotm" _
        , NewTemplate:=False, DocumentType:=0

    Form10_RAMS.Show

End Sub

Sub ENSRAMS()
'
' Noise at Work RAMS
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\ENS RAMS.dotm" _
        , NewTemplate:=False, DocumentType:=0

    Form10_RAMS.Show

End Sub

Sub EnvPol()
'
' Environmental Policy
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\Environmental Policy.docx" _
        , NewTemplate:=False, DocumentType:=0


End Sub

Sub HASPol()
'
' Health & Safety Policy
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\Health and Safety Policy.docx" _
        , NewTemplate:=False, DocumentType:=0


End Sub

Sub QalPol()
'
' Quality Policy
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\Quality Policy.docx" _
        , NewTemplate:=False, DocumentType:=0


End Sub

Sub GDPR()
'
' GDPR
'
'
    Documents.Add Template:= _
        AddinFolder & "\8. Procedures and RAMS\GDPR.docx" _
        , NewTemplate:=False, DocumentType:=0


End Sub

