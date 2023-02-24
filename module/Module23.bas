Attribute VB_Name = "Module23"
Public Function CleanString(ByVal sText As String)
    CleanString = Replace(Replace(sText, Chr$(7), vbNullString), vbCr, vbNullString)
End Function
Sub Phieu_ZipGrade(ByVal control As Office.IRibbonControl)
socau = ActiveDocument.Tables(1).Rows.Count
If socau > 40 Then
 Call Phieu_ZipGrade_50
End If
If (socau > 30) And (socau <= 40) Then
 Call Phieu_ZipGrade_40
End If
If (socau > 20) And (socau <= 30) Then
 Call Phieu_ZipGrade_30
End If
If (socau > 10) And (socau <= 20) Then
 Call Phieu_ZipGrade_20
End If
End Sub
Sub Phieu_ZipGrade_20()
Dim ThisDoc, ThatDoc As Document
Dim socau As Byte
 Set ThisDoc = ActiveDocument
   socau = ThisDoc.Tables(1).Rows.Count - 1
'Mo file chua bang dap an mau
    Documents.Open FileName:="D:\ZIPGRADE\DAP AN ZIPGRADE.docx"
    Selection.WholeStory
    Selection.Copy
    Documents.add DocumentType:=wdNewBlankDocument
    Set indapan = ActiveDocument
    Selection.Paste
    Windows("DAP AN ZIPGRADE.docx").Close
'Quet dap an tung ma va to len phieu
For j = 1 To 4
  'To 10 cau dau tien
        indapan.Tables(j).Rows(3).Cells(1).Range = ThisDoc.Tables(1).Rows(1).Cells(j + 1).Range
        For i = 1 To 10
         Dapan = CleanString(ThisDoc.Tables(1).Rows(1 + i).Cells(j + 1).Range)
         Select Case Dapan
          Case "A"
            indapan.Tables(j).Rows(i + 16).Cells(6).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "B"
            indapan.Tables(j).Rows(i + 16).Cells(7).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "C"
          indapan.Tables(j).Rows(i + 16).Cells(8).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
        Case Else
          indapan.Tables(j).Rows(i + 16).Cells(9).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End Select
        Next i
     T = (socau - 1) Mod 10
    'To dap an cho 11 den 20
         For i = 1 To T + 1
            Dapan = CleanString(ThisDoc.Tables(1).Rows(11 + i).Cells(j + 1).Range)
            Select Case Dapan
             Case "A"
              indapan.Tables(j).Rows(i + 5).Cells(12).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "B"
              indapan.Tables(j).Rows(i + 5).Cells(13).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "C"
              indapan.Tables(j).Rows(i + 5).Cells(14).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case Else
              indapan.Tables(j).Rows(i + 5).Cells(15).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             End Select
        Next i
 Next j
 MsgBox "Da to xong cac phieu dap an"
 Exit Sub
End Sub
Sub Phieu_ZipGrade_30()
Dim ThisDoc, ThatDoc As Document
Dim socau As Byte
 Set ThisDoc = ActiveDocument
   socau = ThisDoc.Tables(1).Rows.Count - 1
'Mo file chua bang dap an mau
    Documents.Open FileName:="D:\ZIPGRADE\DAP AN ZIPGRADE.docx"
    Selection.WholeStory
    Selection.Copy
    Documents.add DocumentType:=wdNewBlankDocument
    Set indapan = ActiveDocument
    Selection.Paste
    Windows("DAP AN ZIPGRADE.docx").Close
'Quet dap an tung ma va to len phieu
For j = 1 To 4
  'To 10 cau dau tien
        indapan.Tables(j).Rows(3).Cells(1).Range = ThisDoc.Tables(1).Rows(1).Cells(j + 1).Range
        For i = 1 To 10
         Dapan = CleanString(ThisDoc.Tables(1).Rows(1 + i).Cells(j + 1).Range)
         Select Case Dapan
          Case "A"
            indapan.Tables(j).Rows(i + 16).Cells(6).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "B"
            indapan.Tables(j).Rows(i + 16).Cells(7).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "C"
          indapan.Tables(j).Rows(i + 16).Cells(8).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
        Case Else
          indapan.Tables(j).Rows(i + 16).Cells(9).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End Select
        Next i
    
    'To dap an cho 11 den 20
         For i = 1 To 10
            Dapan = CleanString(ThisDoc.Tables(1).Rows(11 + i).Cells(j + 1).Range)
            Select Case Dapan
             Case "A"
              indapan.Tables(j).Rows(i + 5).Cells(12).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "B"
              indapan.Tables(j).Rows(i + 5).Cells(13).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "C"
              indapan.Tables(j).Rows(i + 5).Cells(14).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case Else
              indapan.Tables(j).Rows(i + 5).Cells(15).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             End Select
        Next i
    'To dap an cho 21 den 30
        
         T = (socau - 1) Mod 10
         For i = 1 To T + 1
        If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "A" Then
          indapan.Tables(j).Rows(i + 16).Cells(12).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "B" Then
          indapan.Tables(j).Rows(i + 16).Cells(13).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "C" Then
          indapan.Tables(j).Rows(i + 16).Cells(14).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "D" Then
          indapan.Tables(j).Rows(i + 16).Cells(15).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
        Next i
     Next j
 MsgBox "Da to xong cac phieu dap an"
 Exit Sub
End Sub
Sub Phieu_ZipGrade_40()
Dim ThisDoc, ThatDoc As Document
Dim socau As Byte
 Set ThisDoc = ActiveDocument
   socau = ThisDoc.Tables(1).Rows.Count - 1
'Mo file chua bang dap an mau
    Documents.Open FileName:="D:\ZIPGRADE\DAP AN ZIPGRADE.docx"
    Selection.WholeStory
    Selection.Copy
    Documents.add DocumentType:=wdNewBlankDocument
    Set indapan = ActiveDocument
    Selection.Paste
    Windows("DAP AN ZIPGRADE.docx").Close
'Quet dap an tung ma va to len phieu
For j = 1 To 4
  'To 10 cau dau tien
        indapan.Tables(j).Rows(3).Cells(1).Range = ThisDoc.Tables(1).Rows(1).Cells(j + 1).Range
        For i = 1 To 10
         Dapan = CleanString(ThisDoc.Tables(1).Rows(1 + i).Cells(j + 1).Range)
         Select Case Dapan
          Case "A"
            indapan.Tables(j).Rows(i + 16).Cells(6).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "B"
            indapan.Tables(j).Rows(i + 16).Cells(7).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "C"
          indapan.Tables(j).Rows(i + 16).Cells(8).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
        Case Else
          indapan.Tables(j).Rows(i + 16).Cells(9).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End Select
        Next i
    
    'To dap an cho 11 den 20
         For i = 1 To 10
            Dapan = CleanString(ThisDoc.Tables(1).Rows(11 + i).Cells(j + 1).Range)
            Select Case Dapan
             Case "A"
              indapan.Tables(j).Rows(i + 5).Cells(12).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "B"
              indapan.Tables(j).Rows(i + 5).Cells(13).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "C"
              indapan.Tables(j).Rows(i + 5).Cells(14).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case Else
              indapan.Tables(j).Rows(i + 5).Cells(15).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             End Select
        Next i
    'To dap an cho 21 den 30
         For i = 1 To 10
        If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "A" Then
          indapan.Tables(j).Rows(i + 16).Cells(12).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "B" Then
          indapan.Tables(j).Rows(i + 16).Cells(13).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "C" Then
          indapan.Tables(j).Rows(i + 16).Cells(14).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "D" Then
          indapan.Tables(j).Rows(i + 16).Cells(15).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
        Next i
    'To dap an cho 31 den 40
               T = (socau - 1) Mod 10
        For i = 1 To T + 1
        If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "A" Then
          indapan.Tables(j).Rows(i + 5).Cells(18).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "B" Then
          indapan.Tables(j).Rows(i + 5).Cells(19).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "C" Then
          indapan.Tables(j).Rows(i + 5).Cells(20).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "D" Then
          indapan.Tables(j).Rows(i + 5).Cells(21).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
        Next i
    Next j
 MsgBox "Da to xong cac phieu dap an"
 Exit Sub
End Sub
Sub Phieu_ZipGrade_50()
Dim ThisDoc, ThatDoc As Document
Dim socau As Byte
 Set ThisDoc = ActiveDocument
   socau = ThisDoc.Tables(1).Rows.Count - 1
'Mo file chua bang dap an mau
    Documents.Open FileName:="D:\ZIPGRADE\DAP AN ZIPGRADE.docx"
    Selection.WholeStory
    Selection.Copy
    Documents.add DocumentType:=wdNewBlankDocument
    Set indapan = ActiveDocument
    Selection.Paste
    Windows("DAP AN ZIPGRADE.docx").Close
'Quet dap an tung ma va to len phieu
For j = 1 To 4
  'To 10 cau dau tien
        indapan.Tables(j).Rows(3).Cells(1).Range = ThisDoc.Tables(1).Rows(1).Cells(j + 1).Range
        For i = 1 To 10
         Dapan = CleanString(ThisDoc.Tables(1).Rows(1 + i).Cells(j + 1).Range)
         Select Case Dapan
          Case "A"
            indapan.Tables(j).Rows(i + 16).Cells(6).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "B"
            indapan.Tables(j).Rows(i + 16).Cells(7).Select
            Selection.TypeText ""
            Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                  Unicode:=True
        Case "C"
          indapan.Tables(j).Rows(i + 16).Cells(8).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
        Case Else
          indapan.Tables(j).Rows(i + 16).Cells(9).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End Select
        Next i
    
    'To dap an cho 11 den 20
         For i = 1 To 10
            Dapan = CleanString(ThisDoc.Tables(1).Rows(11 + i).Cells(j + 1).Range)
            Select Case Dapan
             Case "A"
              indapan.Tables(j).Rows(i + 5).Cells(12).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "B"
              indapan.Tables(j).Rows(i + 5).Cells(13).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case "C"
              indapan.Tables(j).Rows(i + 5).Cells(14).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             Case Else
              indapan.Tables(j).Rows(i + 5).Cells(15).Select
              Selection.TypeText ""
              Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                    Unicode:=True
             End Select
        Next i
    'To dap an cho 21 den 30
              For i = 1 To 10
        If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "A" Then
          indapan.Tables(j).Rows(i + 16).Cells(12).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "B" Then
          indapan.Tables(j).Rows(i + 16).Cells(13).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "C" Then
          indapan.Tables(j).Rows(i + 16).Cells(14).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(21 + i).Cells(j + 1).Range) = "D" Then
          indapan.Tables(j).Rows(i + 16).Cells(15).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
        Next i
    'To dap an cho 31 den 40
         For i = 1 To 10
        If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "A" Then
          indapan.Tables(j).Rows(i + 5).Cells(18).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "B" Then
          indapan.Tables(j).Rows(i + 5).Cells(19).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "C" Then
          indapan.Tables(j).Rows(i + 5).Cells(20).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(31 + i).Cells(j + 1).Range) = "D" Then
          indapan.Tables(j).Rows(i + 5).Cells(21).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
        Next i
     
            T = (socau - 1) Mod 10
         For i = 1 To T + 1
        If CleanString(ThisDoc.Tables(1).Rows(41 + i).Cells(j + 1).Range) = "A" Then
          indapan.Tables(j).Rows(i + 16).Cells(18).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(41 + i).Cells(j + 1).Range) = "B" Then
          indapan.Tables(j).Rows(i + 16).Cells(19).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(41 + i).Cells(j + 1).Range) = "C" Then
          indapan.Tables(j).Rows(i + 16).Cells(20).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
         If CleanString(ThisDoc.Tables(1).Rows(41 + i).Cells(j + 1).Range) = "D" Then
          indapan.Tables(j).Rows(i + 16).Cells(21).Select
          Selection.TypeText ""
          Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-3944, _
                Unicode:=True
         End If
        Next i
 Next j
 MsgBox "Da to xong cac phieu dap an"
 Exit Sub
End Sub

