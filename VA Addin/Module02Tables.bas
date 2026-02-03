Attribute VB_Name = "Module02Tables"

Sub CustomTable()
'
' Inserts Custom Table referring to InsertTable form
'
   InsertTableVA.Show
   
End Sub


Sub OctTable()
'
' Insert Octave table 63Hz - 8kHz

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\2. Tables\Tables Text.docx", Range:="T638k", Link:=False
    
        Selection.WholeStory
    Selection.LanguageID = wdEnglishUK
    Selection.NoProofing = False
    Application.CheckLanguage = False
    
End Sub

Sub ShortOctTable()
'
' Insert Octave table 125Hz - 4kHz

Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertFile fileName:=AddinFolder & "\2. Tables\Tables Text.docx", Range:="T1254k", Link:=False
    
        Selection.WholeStory
    Selection.LanguageID = wdEnglishUK
    Selection.NoProofing = False
    Application.CheckLanguage = False
    
End Sub
