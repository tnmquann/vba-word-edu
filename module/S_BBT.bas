Attribute VB_Name = "S_BBT"
Dim shpCanvas, shpLine, shpTextBox As Shape
Sub BBT_b2()
    If S_bbtF.cocuctri.Value Then
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=205, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=23, EndX:=195, EndY:=23)
     
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=43, EndX:=195, EndY:=43)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=93)
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " " 'ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=165, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=62, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=102, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=140, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
    
        'Dong 3
        If S_bbtF.aduong.Value = True Then
        y_dau = 38
        y_cuoi = 70
        Else
        y_dau = 70
        y_cuoi = 38
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=165, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aduong.Value = True Then
        y_dau = 58
        y_cuoi = 85
        Else
        y_dau = 85
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=60, BeginY:=y_dau, EndX:=105, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=120, BeginY:=y_cuoi, EndX:=165, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
    Else
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=205, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=24, EndX:=195, EndY:=24)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=72)
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=165, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
      
        'Dong 3
        If S_bbtF.aduong.Value = True Then
        y_dau = 21
        y_cuoi = 55
        Else
        y_dau = 55
        y_cuoi = 21
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=33, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=165, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aduong.Value = True Then
        y_dau = 35
        y_cuoi = 65
        Else
        y_dau = 65
        y_cuoi = 35
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=60, BeginY:=y_dau, EndX:=105, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=120, BeginY:=y_cuoi, EndX:=165, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
    End If
End Sub
Sub BBT_b3()

    Dim shpCanvas As Shape
    Dim shpLine, shpTextBox As Shape
    Dim y_dau, y_cuoi As Integer
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=260, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=23, EndX:=250, EndY:=23)
     
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=43, EndX:=250, EndY:=43)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=93)
    If S_bbtF.khongcuctri = False Then
    
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=151, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCT
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=62, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=102, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=130, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=161, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=195, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        'Dong 3
        If S_bbtF.aam.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=151, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCT
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=y_cuoi, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aam.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=60, BeginY:=y_dau, EndX:=105, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=120, BeginY:=y_cuoi, EndX:=165, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=180, BeginY:=y_dau, EndX:=225, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
    Else
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
             
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
     
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=130, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211)
        Else
        Selection.TypeText text:="+"
        End If
        
        
        'Dong 3
        If S_bbtF.aam.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=y_cuoi, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aam.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=60, BeginY:=y_dau, EndX:=225, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
    End If
End Sub
Sub BBT_b4()

    Dim shpCanvas As Shape
    Dim shpLine, shpTextBox As Shape
    Dim y_dau, y_cuoi As Integer
    
    If S_bbtF.khongcuctri = False Then
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=315, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=23, EndX:=305, EndY:=23)
     
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=43, EndX:=305, EndY:=43)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=93)
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=151, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCT
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=209, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=Mid(S_bbtF.xCD, 2)
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=280, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=62, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=102, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=130, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=161, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=195, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=220, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=255, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        'Dong 3
        If S_bbtF.aduong.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=151, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCT
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=210, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=280, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aduong.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=60, BeginY:=y_dau, EndX:=105, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=120, BeginY:=y_cuoi, EndX:=165, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=180, BeginY:=y_dau, EndX:=225, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=235, BeginY:=y_cuoi, EndX:=285, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
    Else
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=205, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=23, EndX:=195, EndY:=23)
     
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=43, EndX:=195, EndY:=43)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=93)
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=3, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=165, Top:=3, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=62, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=102, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=140, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
    
        'Dong 3
        If S_bbtF.aduong.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=92, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=165, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aduong.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aduong.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=60, BeginY:=y_dau, EndX:=105, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=120, BeginY:=y_cuoi, EndX:=165, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With

    End If
End Sub
Sub BBT_1_1()

    Dim shpCanvas As Shape
    Dim shpLine, shpTextBox As Shape
    Dim y_dau, y_cuoi As Integer
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=260, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=23, EndX:=250, EndY:=23)
     
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=43, EndX:=250, EndY:=43)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=93)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=140, BeginY:=26, EndX:=140, EndY:=93)
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=142, BeginY:=26, EndX:=142, EndY:=93)
    
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=122, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
          
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=75, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
                
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=185, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        'Dong 3
        If S_bbtF.aam.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=S_bbtF.yCD
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=110, Top:=y_cuoi, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=220, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.yCD
        
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=140, Top:=y_dau, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        If S_bbtF.aam.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=57, BeginY:=y_dau, EndX:=115, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=170, BeginY:=y_dau, EndX:=227, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
    
End Sub

Sub BBT_2_1()

    Dim shpCanvas As Shape
    Dim shpLine As Shape
    Dim shpTextBox As Shape
    Dim y_dau, y_cuoi As Integer
    Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=260, Height:=90)
    shpCanvas.WrapFormat.Type = wdWrapSquare
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=23, EndX:=250, EndY:=23)
     
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=5, BeginY:=43, EndX:=250, EndY:=43)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=30, BeginY:=8, EndX:=30, EndY:=93)
    
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=140, BeginY:=23, EndX:=140, EndY:=43)
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=142, BeginY:=23, EndX:=142, EndY:=43)
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=140, BeginY:=45, EndX:=140, EndY:=93)
    Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=142, BeginY:=45, EndX:=142, EndY:=93)
    If S_bbtF.cocuctri = False Then
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=122, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
          
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=72, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
                
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=190, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        'Dong 3
        If S_bbtF.aam.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=110, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=140, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=220, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        
        If S_bbtF.aam.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=57, BeginY:=y_dau, EndX:=115, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=170, BeginY:=y_dau, EndX:=227, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
    Else
        'Dong 1
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=1, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="x"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=ChrW(8211) & " "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=67, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=122, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=(Val(S_bbtF.xCD.Value) + Val(S_bbtF.xCT.Value)) / 2
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=172, Top:=4, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=S_bbtF.xCT
          
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=225, Top:=4, Width:=35, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="+ "
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        'Dong 2
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=18, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y'"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=52, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
                
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=75, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
                
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=105, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=152, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
                
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=182, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:="0"
                
                     
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=205, Top:=23, Width:=30, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        
        'Dong 3
        If S_bbtF.aam.Value = True Then
        y_dau = 41
        y_cuoi = 73
        Else
        y_dau = 73
        y_cuoi = 41
        End If
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=5, Top:=53, Width:=30, Height:=25)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.Font.Name = "Euclid"
        Selection.Font.Italic = True
        Selection.TypeText text:="y"
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=32, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
            
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=75, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=S_bbtF.yCD
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=110, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:="+ "
        Else
        Selection.TypeText text:=ChrW(8211) & " "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=140, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=185, Top:=y_dau, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        Selection.TypeText text:=S_bbtF.yCT
        
        Set shpTextBox = shpCanvas.CanvasItems.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=220, Top:=y_cuoi, Width:=40, Height:=20)
        shpTextBox.Line.Visible = msoFalse
        shpTextBox.Fill.Visible = msoFalse
        
        If S_bbtF.aam.Value = True Then
        Selection.TypeText text:=ChrW(8211) & " "
        Else
        Selection.TypeText text:="+ "
        End If
        Selection.InsertSymbol Font:="Times New Roman", CharacterNumber:=8734, Unicode:=True 'vo cuc
        
        
        If S_bbtF.aam.Value = True Then
        y_dau = 58
        y_cuoi = 88
        Else
        y_dau = 88
        y_cuoi = 58
        End If
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=55, BeginY:=y_dau, EndX:=80, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=90, BeginY:=y_cuoi, EndX:=115, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=165, BeginY:=y_cuoi, EndX:=190, EndY:=y_dau)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        Set shpLine = shpCanvas.CanvasItems.AddLine(BeginX:=200, BeginY:=y_dau, EndX:=225, EndY:=y_cuoi)
        With shpLine.Line
         .EndArrowheadStyle = msoArrowheadStealth
         .BeginArrowheadWidth = msoArrowheadWide
         .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255)
        End With
        
    End If
End Sub

