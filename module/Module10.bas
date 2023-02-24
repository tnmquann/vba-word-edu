Attribute VB_Name = "Module10"
Option Explicit
Public isCut_ As Boolean
Const SO_LOI_DUOC_DUYET = 2000
Dim i As Integer
Dim ab() As Integer
Dim dapanmoi() As Integer
Dim Title, msg As String
Dim InAns() As String
Public Function FExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
FExists = fs.fileexists(OrigFile)
End Function 'Returns a boolean - True if the file exists
Sub K_HightlightPictureMathType(ByVal control As Office.IRibbonControl)  '() '
'
' Macro5 Macro
'
    Dim iLoi As Integer
    Selection.HomeKey Unit:=wdStory
    'SoCongThuc = ActiveDocument.InlineShapes.Count
    iLoi = 0
    Dim shp As InlineShape
            For Each shp In ActiveDocument.InlineShapes
    'For iDem = 1 To SoCongThuc
'        With Selection.Find
'            .Text = "^g"
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .MatchWildcards = False
'        End With
'        Selection.Find.Execute
        'MsgBox wdInlineShapePicture 'ActiveDocument.InlineShapes.Item(iDem).Type
'        Select Case ActiveDocument.InlineShapes.Item(iDem).Type
'            Case wdInlineShapeChart
'                MsgBox "wdInlineShapeChart"
'            Case wdInlineShapeDiagram
'                MsgBox "wdInlineShapeDiagram"
'            ' Nhieu ............ Nhung cai nhu mathtype
'            Case wdInlineShapeEmbeddedOLEObject
'                MsgBox "wdInlineShapeEmbeddedOLEObject"
'            Case wdInlineShapeHorizontalLine
'                MsgBox "wdInlineShapeHorizontalLine"
'            Case wdInlineShapeLinkedOLEObject
'                MsgBox "wdInlineShapeLinkedOLEObject"
'            Case wdInlineShapeLinkedPicture
'                MsgBox "wdInlineShapeLinkedPicture"
'            Case wdInlineShapeLinkedPictureHorizontalLine
'                MsgBox "wdInlineShapeLinkedPictureHorizontalLine"
'            Case wdInlineShapeLockedCanvas
'                MsgBox "wdInlineShapeLockedCanvas"
'            Case wdInlineShapeOLEControlObject
'                MsgBox "wdInlineShapeOLEControlObject"
'            Case wdInlineShapeOWSAnchor
'                MsgBox "wdInlineShapeOWSAnchor"
'            Case wdInlineShapePicture
'                MsgBox "wdInlineShapePicture"
'            Case wdInlineShapePictureBullet
'                MsgBox "wdInlineShapePictureBullet"
'            Case wdInlineShapePictureHorizontalLine
'                MsgBox "wdInlineShapePictureHorizontalLine"
'            Case wdInlineShapeScriptAnchor
'                MsgBox "wdInlineShapeScriptAnchor"
'            Case wdInlineShapeSmartArt
'                MsgBox "wdInlineShapeSmartArt"
'        End Select
        'MsgBox ActiveDocument.InlineShapes.Item(iDem).Type
        
        'If ActiveDocument.InlineShapes.Item(iDem).Type = wdInlineShapePicture Then
        If shp.Type = wdInlineShapePicture Then
            'MsgBox ActiveDocument.InlineShapes.Item(iDem).Line.Style
            iLoi = iLoi + 1
            If (iLoi > SO_LOI_DUOC_DUYET) Then Exit For
            shp.Select
            'Selection.Range.HighlightColorIndex = wdPink 'wdBrightGreen
            Selection.Font.Underline = wdUnderlineThick
            Selection.Font.UnderlineColor = wdColorRed
        End If
    Next 'iDem
    If (iLoi <= SO_LOI_DUOC_DUYET) Then
'            For Each shp In ActiveDocument.InlineShapes
'            If shp.Type = wdInlineShapePicture Then
'                shp.Select
'                'Selection.Range.HighlightColorIndex = wdPink 'wdBrightGreen
'                Selection.Font.Underline = wdUnderlineThick
'                Selection.Font.UnderlineColor = wdColorRed
'            End If
'            Next
        MsgBox "So anh tren dong (khong phai cong thuc): " & iLoi & vbNewLine & _
            "Hay kiem tra lai cong thuc. " & vbNewLine & _
            "Neu chac chan khong co cong thuc hoa anh hay click Bo gach duoi anh"
    ElseIf iLoi = 0 Then
        MsgBox "KHONG TIM THAY LOI"
    Else
        MsgBox "CO NHIEU HON " & SO_LOI_DUOC_DUYET & " ANH TREN DONG (KHONG PHAI CONG THUC)! " & vbNewLine & _
        "CHI GACH DUOI " & SO_LOI_DUOC_DUYET & " ANH DAU!"
    End If
End Sub

Sub K_BoGachDuoiAnh(ByVal control As Office.IRibbonControl)  '() '
'
' Macro5 Macro
'
    Dim iLoi As Integer
    
    Selection.HomeKey Unit:=wdStory
    'SoCongThuc = ActiveDocument.InlineShapes.Count
    
    iLoi = 0 ' If (ActiveDocument.InlineShapes.Count > 1000
    Dim shp As InlineShape
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Then
                'MsgBox ActiveDocument.InlineShapes.Item(iDem).Line.Style
                iLoi = iLoi + 1
                If (iLoi > SO_LOI_DUOC_DUYET) Then Exit For
                shp.Select
                'Selection.Range.HighlightColorIndex = wdPink 'wdBrightGreen
                Selection.Font.Underline = wdUnderlineNone
                Selection.Font.UnderlineColor = wdColorAutomatic
        End If
    Next 'iDem
    If (iLoi <= SO_LOI_DUOC_DUYET) Then
'            For Each shp In ActiveDocument.InlineShapes
'            If shp.Type = wdInlineShapePicture Then
'                shp.Select
'                'Selection.Range.HighlightColorIndex = wdPink 'wdBrightGreen
'                Selection.Font.Underline = wdUnderlineThick
'                Selection.Font.UnderlineColor = wdColorRed
'            End If
'            Next
        MsgBox "So anh tren dong (khong phai cong thuc) da bo gach duoi: " & iLoi
    ElseIf iLoi = 0 Then
        MsgBox "KHONG TIM THAY !"
    Else
        MsgBox "CO NHIEU HON " & SO_LOI_DUOC_DUYET & " ANH TREN DONG (KHONG PHAI CONG THUC)! " & vbNewLine & _
        "CHI BO GACH DUOI " & SO_LOI_DUOC_DUYET & " ANH DAU!"
    End If
End Sub
Sub K_ThietLapHinhAnhInline_Temp(ByVal control As Office.IRibbonControl)  '() '
    Dim CountImage As Integer, iCount As Integer, Ret_type As Integer
    Dim strMsg As String, strTitle As String
    'Dim oShp As Shape
    Dim isContinue As Boolean
    isContinue = True
     CountImage = 0
     ' Dialog Message
        strMsg = "ÐaÞ ðýa môòt aÒnh vêÌ position In line with text! " & vbNewLine & _
                "Baòn muôìn thýòc hiêòn tâìt caÒ hay lâÌn lýõòt? " & vbNewLine & _
                "- Nhâìn Yes ðêÒ thýòc hiêòn tâìt caÒ!" & vbNewLine & _
                "- Nhâìn No ðêÒ thýòc hiêòn lâÌn lýõòt!" & vbNewLine & _
                "- Nhâìn Cancel ðêÒ thoaìt."
        ' Dialog's Title
        strTitle = "XýÒ liì ðiònh daòng aÒnh!"
        Dim totalShapes As Integer
        totalShapes = ActiveDocument.Shapes.Count
        
        'MsgBox totalShapes
        Dim shp As Shape
        For Each shp In ActiveDocument.Shapes
            'MsgBox shp.Type
            If (shp.Type = msoPicture Or shp.Type = msoGroup) Then
                shp.Select
                Selection.ShapeRange.WrapFormat.Type = wdWrapInline
'            ElseIf (shp.Type = msoGroup) Then
'                shp.Select
'                Selection.ShapeRange.WrapFormat.Type = wdWrapInline
'                'shp.ConvertToInlineShape
            End If
        Next
'    For iCount = 1 To ActiveDocument.Shapes.Count
'        'MsgBox ActiveDocument.Shapes(iCount).Type
'        'If (iCount <= ActiveDocument.Shapes.Count) Then
'
'            'ActiveDocument.Shapes(1).ConvertToInlineShape ' ham tren ko hay, thu ham nay
'            ActiveDocument.Shapes(1).Select
'            Selection.ShapeRange.WrapFormat.Type = wdWrapInline
'
'            CountImage = CountImage + 1
'            If isContinue Then
'                'Display MessageBox
'                Ret_type = MsgBox(strMsg, vbYesNoCancel + vbQuestion, strTitle)
'                ' Check pressed button
'                Select Case Ret_type
'                Case 6  ' Yes
'                    'MsgBox "You clicked 'YES' button."
'                    isContinue = False
'                Case 7  ' No
'                    'MsgBox "You clicked 'NO' button."
'                    If (ActiveDocument.Shapes.Count = 1) Then isContinue = False
'                Case 2  ' Cancel
'                    'MsgBox "You clicked 'CANCEL' button."
'                    isContinue = False
'                    Exit Sub
'                End Select
'            End If
'       'End If
'    Next iCount
    'MsgBox "DA " & CountImage
'    If (CountImage = 0) Then
'        MsgBox "Không thâìy hiÌnh aÒnh câÌn ðiònh daòng In line with text"
'    Else
'        MsgBox "ÐaÞ ðiònh daòng " & CountImage & "  hiÌnh aÒnh vêÌ daòng In line with text"
'    End If
    If (totalShapes > ActiveDocument.Shapes.Count) Then
        MsgBox "HAY CHAY LAI LAN NUA"
    Else
        MsgBox "XONG"
    End If
End Sub


