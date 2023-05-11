Public Sub PictureInlineWithText(ByVal control As Office.IRibbonControl)
    Dim i As Integer
    For i = ActiveDocument.Shapes.Count To 1 Step -1
        Select Case ActiveDocument.Shapes(i).Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoLinkedPicture, msoOLEControlObject, msoPicture, wdInlineShapePicture, wdInlineShapeLinkedPicture
            ActiveDocument.Shapes(i).ConvertToInlineShape
        End Select
    Next i
    Selection.HomeKey Unit:=wdStory
End Sub

Public Sub PicCenter(ByVal control As Office.IRibbonControl)
    Dim Pic As InlineShape
    Selection.HomeKey Unit:=wdStory
    For Each Pic In ActiveDocument.InlineShapes
        If Pic.Type = wdInlineShapePicture Then
            Pic.Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next
End Sub

Public Sub InlineAndCenterAllImages(ByVal control As Office.IRibbonControl)
    Dim i As Integer
    For i = ActiveDocument.Shapes.Count To 1 Step -1
        Select Case ActiveDocument.Shapes(i).Type
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoLinkedPicture, msoOLEControlObject, msoPicture, wdInlineShapePicture, wdInlineShapeLinkedPicture
            ActiveDocument.Shapes(i).ConvertToInlineShape
        End Select
    Next i
    Selection.HomeKey Unit:=wdStory
    Dim Pic As InlineShape
    Selection.HomeKey Unit:=wdStory
    For Each Pic In ActiveDocument.InlineShapes
        If Pic.Type = wdInlineShapePicture Then
            Pic.Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next
End Sub