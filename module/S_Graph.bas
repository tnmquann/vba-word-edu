Attribute VB_Name = "S_Graph"
Sub GraphDT(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long, ByRef x3 As Long, ByRef y3 As Long, _
            ByRef x4 As Long, ByRef y4 As Long, ByRef x5 As Long, ByRef y5 As Long, ByRef x6 As Long, ByRef y6 As Long, _
            ByRef x7 As Long, ByRef y7 As Long, ByRef x8 As Long, ByRef y8 As Long)

    Dim docNew As Document
    Dim shpCanvas As Shape
    Dim shpLine As Shape
    Dim shpTextBox As Shape
    Dim sngArray(1 To 7, 1 To 2) As Single
    
    Set docNew = ActiveDocument
    'Create a new drawing canvas
    Set shpCanvas = docNew.Shapes.AddCanvas(Left:=100, Top:=100, Width:=110, Height:=160)
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=x8, BeginY:=160, EndX:=x8, EndY:=0)
    With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    shpLine.Name = "s1"
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=0, BeginY:=y8, EndX:=110, EndY:=y8)
    With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    shpLine.Name = "s2"
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x8 - 20, Top:=0, Width:=20, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
    
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=110, Top:=y8 - 10, Width:=20, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
    
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
        
    'docNew.sShapeRange(Array(GroupItems(1), "s2")).Group
    
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x8 - 20, Top:=y8 - 5, Width:=20, Height:=27)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="O"
    
    
    sngArray(1, 1) = x1
    sngArray(1, 2) = y1
    sngArray(2, 1) = x2
    sngArray(2, 2) = y2
    sngArray(3, 1) = x3
    sngArray(3, 2) = y3
    sngArray(4, 1) = x4
    sngArray(4, 2) = y4
    sngArray(5, 1) = x5
    sngArray(5, 2) = y5
    sngArray(6, 1) = x6
    sngArray(6, 2) = y6
    sngArray(7, 1) = x7
    sngArray(7, 2) = y7
Dim gra3 As Shape
    Set gra3 = shpCanvas.CanvasItems.AddCurve(SafeArrayOfPoints:=sngArray)
    If S_bbtF.aam Then
        gra3.Flip msoFlipVertical
        gra3.Top = 100
    End If

End Sub
Sub GraphHT(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long, ByRef x3 As Long, ByRef y3 As Long, _
            ByRef x4 As Long, ByRef y4 As Long, ByRef x5 As Long, ByRef y5 As Long, ByRef x6 As Long, ByRef y6 As Long, _
            ByRef x7 As Long, ByRef y7 As Long, ByRef x8 As Long, ByRef y8 As Long)

    Dim docNew As Document
    Dim shpCanvas As Shape
    Dim shpLine As Shape
    Dim sngArray(1 To 7, 1 To 2) As Single
    
    Set docNew = ActiveDocument
    'Create a new drawing canvas
    Set shpCanvas = docNew.Shapes.AddCanvas(Left:=100, Top:=120, Width:=160, Height:=160)
    'He truc
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=x8 - Val(S_bbtF.xCD) * 15, BeginY:=160, EndX:=x8 - Val(S_bbtF.xCD) * 15, EndY:=0)
    With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=0, BeginY:=y8 + Val(S_bbtF.yCD) * 15, EndX:=160, EndY:=y8 + Val(S_bbtF.yCD) * 15)
    With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    'Tiem can
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=x8, BeginY:=160, EndX:=x8, EndY:=0)
    With shpLine.Line
         '.EndArrowheadStyle = msoArrowheadStealth
         '.BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=0, BeginY:=y8, EndX:=160, EndY:=y8)
    With shpLine.Line
         '.EndArrowheadStyle = msoArrowheadStealth
         '.BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=x8 - 5 - Val(S_bbtF.xCD) * 15, Top:=0, Width:=20, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=160, Top:=y8 - 10 + Val(S_bbtF.yCD) * 15, Width:=20, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=x8 - 15 - S_bbtF.xCD * 10, Top:=y8 - 2 + S_bbtF.yCD * 10, Width:=20, Height:=27)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="O"
        
    sngArray(1, 1) = x1
    sngArray(1, 2) = y1
    sngArray(2, 1) = x2
    sngArray(2, 2) = y2
    sngArray(3, 1) = x3
    sngArray(3, 2) = y3
    sngArray(4, 1) = x4
    sngArray(4, 2) = y4
    sngArray(5, 1) = x5
    sngArray(5, 2) = y5
    sngArray(6, 1) = x6
    sngArray(6, 2) = y6
    sngArray(7, 1) = x7
    sngArray(7, 2) = y7
    Dim gra1, gra2 As Shape
     'Add Bezier curve to drawing canvas
    Set gra1 = shpCanvas.CanvasItems.AddCurve(SafeArrayOfPoints:=sngArray)
    Set gra2 = gra1.Duplicate
    gra2.Top = 30
    gra2.Left = 30
    gra2.Rotation = 180
    gra2.Top = 0
    gra2.Left = 80
    If S_bbtF.aduong Then
        gra1.Flip msoFlipVertical
        gra2.Flip msoFlipVertical
        gra1.Top = 0
        gra2.Top = 30
    End If
End Sub
Sub GraphMu_Loga(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long, ByRef x3 As Long, ByRef y3 As Long, _
            ByRef x4 As Long, ByRef y4 As Long, ByRef x5 As Long, ByRef y5 As Long, ByRef x6 As Long, ByRef y6 As Long, _
            ByRef x7 As Long, ByRef y7 As Long, ByRef x8 As Long, ByRef y8 As Long)

    Dim docNew As Document
    Dim shpCanvas As Shape
    Dim shpLine As Shape
    Dim sngArray(1 To 7, 1 To 2) As Single
    
    Set docNew = ActiveDocument
    'Create a new drawing canvas
    Set shpCanvas = docNew.Shapes.AddCanvas(Left:=100, Top:=120, Width:=160, Height:=160)
    'He truc
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=x8 - Val(S_bbtF.xCD) * 15, BeginY:=160, EndX:=x8 - Val(S_bbtF.xCD) * 15, EndY:=0)
    With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=0, BeginY:=y8 + Val(S_bbtF.yCD) * 15, EndX:=160, EndY:=y8 + Val(S_bbtF.yCD) * 15)
    With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    'Tiem can
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=x8, BeginY:=160, EndX:=x8, EndY:=0)
    With shpLine.Line
         '.EndArrowheadStyle = msoArrowheadStealth
         '.BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=0, BeginY:=y8, EndX:=160, EndY:=y8)
    With shpLine.Line
         '.EndArrowheadStyle = msoArrowheadStealth
         '.BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
    End With
    
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=x8 - 5 - Val(S_bbtF.xCD) * 15, Top:=0, Width:=20, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=160, Top:=y8 - 10 + Val(S_bbtF.yCD) * 15, Width:=20, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
    Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=x8 - 15 - S_bbtF.xCD * 10, Top:=y8 - 2 + S_bbtF.yCD * 10, Width:=20, Height:=27)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="O"
        
    sngArray(1, 1) = x1
    sngArray(1, 2) = y1
    sngArray(2, 1) = x2
    sngArray(2, 2) = y2
    sngArray(3, 1) = x3
    sngArray(3, 2) = y3
    sngArray(4, 1) = x4
    sngArray(4, 2) = y4
    sngArray(5, 1) = x5
    sngArray(5, 2) = y5
    sngArray(6, 1) = x6
    sngArray(6, 2) = y6
    sngArray(7, 1) = x7
    sngArray(7, 2) = y7
    Dim gra1, gra2 As Shape
     'Add Bezier curve to drawing canvas
    Set gra1 = shpCanvas.CanvasItems.AddCurve(SafeArrayOfPoints:=sngArray)
    
End Sub


