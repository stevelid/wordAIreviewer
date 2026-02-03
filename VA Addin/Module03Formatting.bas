Attribute VB_Name = "Module03Formatting"


Sub Qintro()
'
' Appies styles to text
   
    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Introduction")
    
End Sub

Sub QSectionheading()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Section Heading")
    
End Sub

Sub QSection()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Section Text")
    
End Sub
Sub QSectionsubheading()
'
' Appies styles to text
    
    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Section Subheading")
    
End Sub

Sub QSubheading()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Subheading")
    
End Sub

Sub QIndentedsection()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Indented Section Text")
    
End Sub

Sub Qtabletitle()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Table Number")
    
End Sub

Sub Qfigure()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Quote Figure")
    
End Sub

Sub RChapter()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Level 1")
    
End Sub


Sub RSectionheading()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Level 2")
    
End Sub

Sub RSubsection()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Level 3")
    
End Sub

Sub RHeadingL4()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Level 4")
    
End Sub

Sub RSection()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Text")
    
End Sub

Sub RBullet()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Bullet")
    
End Sub

Sub Tableheading()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Table Heading")
    
End Sub

Sub Tabletext()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Table Text")
    
End Sub


Sub RTabletitle()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Table Number")
    
End Sub

Sub Rfigure()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Report Figure")
    
End Sub

'Sub RChapter()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Report Chapter Title")
'
'End Sub
'
'Sub RSectionheading()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Report Section Heading")
'
'End Sub
'
'Sub RSection()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Report Section Text")
'
'End Sub
'
'Sub RSubsection()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Report Italic Subheading")
'
'End Sub
'
'Sub RIndentedsection()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Report Indented Section Text")
'
'End Sub
'
'Sub RTabletitle()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Table Title (Report)")
'
'End Sub
'
'Sub Tabletext()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Table Text")
'
'End Sub

Sub ExpertChapter()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Expert Chapter")
    
End Sub

Sub ExpertSection()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Expert Text")
    
End Sub

Sub ExpertSectionIndent()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Expert Indented Text")
    
End Sub

Sub Experttabletitle()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Expert Table Number")
    
End Sub

Sub Expertfigure()
'
' Appies styles to text

    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
    Selection.Style = ActiveDocument.Styles("Expert Figure")
    
End Sub

'Sub PChapter()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Proof Chapter Title")
'
'End Sub
'
'Sub PSection()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Proof Section Text")
'
'End Sub
'
'Sub PBold()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Proof Bold Subheading")
'
'End Sub
'
'Sub PItalic()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Proof Italic Subheading")
'
'End Sub
'
'Sub PIndentedsection()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Proof Indented Section Text")
'
'End Sub
'
'Sub PTabletitle()
''
'' Appies styles to text
'
'    ActiveDocument.CopyStylesFromTemplate (AddinFolder & "\VA Addin.dotm")
'    Selection.Style = ActiveDocument.Styles("Table Title (Proof)")
'
'End Sub


