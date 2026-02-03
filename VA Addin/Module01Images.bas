Attribute VB_Name = "Module01Images"
Sub Letterhead()
'
' Inserts Letterhead

    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HeaderFooter.Shapes.AddPicture(fileName:= _
        AddinFolder & "\1. Images\Letter Background.jpg" _
        , LinkToFile:=False, SaveWithDocument:=True).Select
    Selection.ShapeRange.Name = "WordPictureWatermark20519607"
    Selection.ShapeRange.PictureFormat.Brightness = 0.5
    Selection.ShapeRange.PictureFormat.Contrast = 0.5
    Selection.ShapeRange.LockAspectRatio = True
    Selection.ShapeRange.Height = CentimetersToPoints(29.7)
    Selection.ShapeRange.Width = CentimetersToPoints(21)
    Selection.ShapeRange.WrapFormat.AllowOverlap = True
    Selection.ShapeRange.WrapFormat.Side = wdWrapNone
    Selection.ShapeRange.WrapFormat.Type = wdWrapBehind
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Left = wdShapeCenter
    Selection.ShapeRange.Top = wdShapeCenter
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

'    ActiveDocument.Sections(1).Range.Select
'    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
'    Selection.HeaderFooter.Shapes.AddPicture(FileName:= _
'        AddinFolder & "\1. Images\Letter Background.jpg" _
'        , LinkToFile:=False, SaveWithDocument:=True).Select
'    Selection.ShapeRange.Name = "WordPictureWatermark20519607"
'    Selection.ShapeRange.PictureFormat.Brightness = 0.5
'    Selection.ShapeRange.PictureFormat.Contrast = 0.5
'    Selection.ShapeRange.LockAspectRatio = True
'    Selection.ShapeRange.Height = CentimetersToPoints(29.76)
'    Selection.ShapeRange.Width = CentimetersToPoints(21.05)
'    Selection.ShapeRange.WrapFormat.AllowOverlap = True
'    Selection.ShapeRange.WrapFormat.Side = wdWrapNone
'    Selection.ShapeRange.WrapFormat.Type = 3
'    Selection.ShapeRange.RelativeHorizontalPosition = _
'        wdRelativeVerticalPositionMargin
'    Selection.ShapeRange.RelativeVerticalPosition = _
'        wdRelativeVerticalPositionMargin
'    Selection.ShapeRange.Left = wdShapeCenter
'    Selection.ShapeRange.Top = wdShapeCenter
'    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    
   '
End Sub

Sub Reportbackground()
'
'  Inserts Report Background
'
    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HeaderFooter.Shapes.AddPicture(fileName:= _
        AddinFolder & "\1. Images\Report Background.jpg" _
        , LinkToFile:=False, SaveWithDocument:=True).Select
    Selection.ShapeRange.Name = "WordPictureWatermark20519607"
    Selection.ShapeRange.PictureFormat.Brightness = 0.5
    Selection.ShapeRange.PictureFormat.Contrast = 0.5
    Selection.ShapeRange.LockAspectRatio = True
    Selection.ShapeRange.Height = CentimetersToPoints(29.76)
    Selection.ShapeRange.Width = CentimetersToPoints(21.05)
    Selection.ShapeRange.WrapFormat.AllowOverlap = True
    Selection.ShapeRange.WrapFormat.Side = wdWrapNone
    Selection.ShapeRange.WrapFormat.Type = 3
    Selection.ShapeRange.RelativeHorizontalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.RelativeVerticalPosition = _
        wdRelativeVerticalPositionMargin
    Selection.ShapeRange.Left = wdShapeCenter
    Selection.ShapeRange.Top = wdShapeCenter
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
End Sub

Sub Logo()
'
' Inserts Header Logo and resizes to correct size
'
    Selection.InlineShapes.AddPicture(fileName:= _
        AddinFolder & "\1. Images\Venta image - large.jpg" _
        , LinkToFile:=False, SaveWithDocument:=True).Select
        
    Dim PercentSize As Integer
    Dim i As Long

     PercentSize = 47
     For i = 1 To Selection.InlineShapes.Count

With Selection.InlineShapes(i)
         Selection.InlineShapes(i).ScaleHeight = PercentSize
         Selection.InlineShapes(i).ScaleWidth = PercentSize

'    Selection.InlineShapes.AddPicture fileName:= _
'        AddinFolder & "\1. Images\Header image.jpg" _
'        , LinkToFile:=False, SaveWithDocument:=True
        
End With
    Next i
        
End Sub

Sub Logolarge()
'
' Inserts Larger Logo Image
'
    Selection.InlineShapes.AddPicture fileName:= _
        AddinFolder & "\1. Images\Venta image - large.jpg" _
        , LinkToFile:=False, SaveWithDocument:=True
        
End Sub


