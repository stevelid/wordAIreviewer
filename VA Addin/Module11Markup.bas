Attribute VB_Name = "Module11Markup"
Sub Allmarkup()
'
' Show all markup in tracked changes
'
'
    With ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupAll
        .View = wdRevisionsViewFinal
    End With
End Sub

Sub Nomarkup()
'
' Show no markup in tracked changes
'
'
    With ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
End Sub

