VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_ReBank 
   Caption         =   "Repair"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   2760
   ClientWidth     =   3375
   OleObjectBlob   =   "S_ReBank.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "S_ReBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim addinBank As String
Dim i As Integer
Dim Title, msg As String
Private Sub Brfile()
Call CheckDrive
If Theo_CD Then
    Select Case ktlop
    Case 10
    addinBank = S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de\"
    Call ShowFolderList(addinBank)
    Case 11
    addinBank = S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de\"
    Call ShowFolderList(addinBank)
    Case 12
    addinBank = S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de\"
    Call ShowFolderList(addinBank)
    End Select
Else
    Select Case ktlop
    Case 10
    addinBank = S_Drive & "S_Bank&Test\S_Bank\Lop 10\"
    Call LayFile(addinBank)
    Case 11
    addinBank = S_Drive & "S_Bank&Test\S_Bank\Lop 11\"
    Call LayFile(addinBank)
    Case 12
    addinBank = S_Drive & "S_Bank&Test\S_Bank\Lop 12\"
    Call LayFile(addinBank)
    End Select
    
End If
End Sub
Private Sub LayFile(ByVal ThuMuc As String)
Dim f As String
If Right(ThuMuc, 1) <> "\" Then ThuMuc = ThuMuc & "\"
f = Dir$(ThuMuc & "*.dat")
S_List.Clear
While Len(f)
    S_List.AddItem f
    f = Dir$
Wend
End Sub
Private Sub ShowFolderList(ByVal ThuMuc As String)
Dim fs, f, f1, s, sf
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(ThuMuc)
Set sf = f.SubFolders
S_List.Clear
For Each f1 In sf
S_List.AddItem f1.Name
Next
'MsgBox s
End Sub
Private Sub boxBai_Change()
On Error GoTo Thoat
Dim tsc As Integer
If S_ReBank.Theo_Bai Then
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 1).Rows(1).Select
    If S_ReBank.OptionButton1.Value = True Then _
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 1).Rows(1).Select
    If S_ReBank.OptionButton2.Value = True Then _
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 2).Rows(1).Select
    If S_ReBank.OptionButton3.Value = True Then _
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 3).Rows(1).Select
    If S_ReBank.OptionButton4.Value = True Then _
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 4).Rows(1).Select
    S_ReBank.socau_M1_LT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 1).Rows.Count - 1) / 6
    S_ReBank.socau_M1_BT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 5).Rows.Count - 1) / 6
    S_ReBank.socau_M2_LT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 2).Rows.Count - 1) / 6
    S_ReBank.socau_M2_BT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 6).Rows.Count - 1) / 6
    S_ReBank.socau_M3_LT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 3).Rows.Count - 1) / 6
    S_ReBank.socau_M3_BT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 7).Rows.Count - 1) / 6
    S_ReBank.socau_M4_LT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 4).Rows.Count - 1) / 6
    S_ReBank.socau_M4_BT = (ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 8).Rows.Count - 1) / 6
    Tongsocaubai = Val(socau_M1_LT) + Val(socau_M1_BT) + Val(socau_M2_LT) + Val(socau_M2_BT) _
    + Val(socau_M3_LT) + Val(socau_M3_BT) + Val(socau_M4_LT) + Val(socau_M4_BT)
    For i = 1 To ActiveDocument.Tables.Count
     tsc = tsc + ActiveDocument.Tables(i).Rows.Count - 1
    Next i
    Tongsocauchuong = tsc / 6
Else
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 1).Rows(1).Select
    If S_ReBank.OptionButton1.Value = True Then _
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 1).Rows(1).Select
    If S_ReBank.OptionButton2.Value = True Then _
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 2).Rows(1).Select
    If S_ReBank.OptionButton3.Value = True Then _
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 3).Rows(1).Select
    If S_ReBank.OptionButton4.Value = True Then _
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 4).Rows(1).Select
    S_ReBank.socau_M1_LT = ""
    S_ReBank.socau_M1_BT = (ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 1).Rows.Count - 1) / 6
    S_ReBank.socau_M2_LT = ""
    S_ReBank.socau_M2_BT = (ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 2).Rows.Count - 1) / 6
    S_ReBank.socau_M3_LT = ""
    S_ReBank.socau_M3_BT = (ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 3).Rows.Count - 1) / 6
    S_ReBank.socau_M4_LT = ""
    S_ReBank.socau_M4_BT = (ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 2) - 1) + 4).Rows.Count - 1) / 6
    Tongsocaubai = Val(socau_M1_BT) + Val(socau_M2_BT) + Val(socau_M3_BT) + Val(socau_M4_BT)
    For i = 1 To ActiveDocument.Tables.Count
     tsc = tsc + ActiveDocument.Tables(i).Rows.Count - 1
    Next i
    Tongsocauchuong = tsc / 6
End If
Exit Sub
Thoat:
    'Dim Title2, msg As String
        'Title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        'msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a m" & ChrW(7903) & " ngân hàng câu h" & ChrW(7887) & "i"
        'Application.Assistant.DoAlert Title2, msg, 0, 4, 0, 0, 0
End Sub


Private Sub ChuyenCau_Click()
S_MoveQuestion.Show
Unload S_ReBank
End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub Danhthutu_Click()
For i = 1 To ActiveDocument.Tables.Count
        If ActiveDocument.Tables(i).Rows.Count > 1 Then
            For j = 1 To ActiveDocument.Tables(i).Rows.Count - 5 Step 6
                    ActiveDocument.Tables(i).Cell(j + 1, 1).Select
                    Selection.TypeText text:="Câu " & (j - 1) / 6 + 1 & ":"
            Next j
        End If
    Next i
End Sub

Private Sub MoFile_Click()
Dim ThisDoc As Document
On Error GoTo S_Quit
Dim add, fnguon As String
Dim selectedFilename As String
Dim item As Byte
Call CheckDrive
If O10.Value = False And O11.Value = False And O12.Value = False Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n l" & ChrW(7899) & "p"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
add = S_ReBank.S_List.Value
Unload S_ReBank
'If FExists(S_Drive & "S_Bank&Test\S_Bank\Lop " & ktlop & "\" & add) Then
    Dim ktMsg As Byte
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n mu" & ChrW(7889) & "n m" & ChrW(7903) & " file này?"
    ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 2, 0, 2, 0)
    If ktMsg = 6 Then
        If Theo_Bai Then
            fnguon = "S_Bank&Test\S_Bank\Lop " & ktlop & "\"
            Set ThisDoc = Documents.Open(S_Drive & fnguon & add, _
                ConfirmConversions:=False, ReadOnly:=False, _
                AddToRecentFiles:=False, PasswordDocument:="159", PasswordTemplate:="", _
                Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
                Format:=wdOpenFormatAuto, XMLTransform:="")
            ThisDoc.Activate
        Else
            ChangeFileOpenDirectory (S_Drive & "S_Bank&Test\S_Bank\Lop " & ktlop & "\Chuyen de\" & add & "\")
            
            With Application.FileDialog(msoFileDialogOpen)
                .AllowMultiSelect = False
                .Show
                item = .SelectedItems.Count
                If item = 0 Then
                    'ReDim selectedFilename(1) As String
                    selectedFilename = ""
                Else
                    'ReDim selectedFilename(item) As String
                    'For i = 1 To item
                        selectedFilename = .SelectedItems(1)
                    'Next
                End If
            End With
            fnguon = selectedFilename
            Set ThisDoc = Documents.Open(fnguon, _
                ConfirmConversions:=False, ReadOnly:=False, _
                AddToRecentFiles:=False, PasswordDocument:="159", PasswordTemplate:="", _
                Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
                Format:=wdOpenFormatAuto, XMLTransform:="")
            ThisDoc.Activate
            
        End If
        
        S_ReBank.Show
        If ktCD Then
        S_ReBank.Theo_CD = True
        Else
        S_ReBank.Theo_Bai = True
        End If
        Select Case ktlop
            Case 10
            S_ReBank.O10.Value = True
            Case 11
            S_ReBank.O11.Value = True
            Case 12
            S_ReBank.O12.Value = True
        End Select
        'For i = 1 To ActiveDocument.Tables.Count / 8
            'S_ReBank.boxBai.AddItem "Bài " & i
        'Next
        'S_ReBank.boxBai.Value = "Bài 1"
    End If
'End If
Exit Sub
S_Quit:
    'Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    'msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n file c" & ChrW(7847) & "n m" & ChrW(7903)
    'Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub

Private Sub OptionButton1_Click()
    On Error Resume Next
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 1).Rows(1).Select
End Sub
Private Sub OptionButton2_Click()
    On Error Resume Next
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 2).Rows(1).Select
End Sub
Private Sub OptionButton3_Click()
    On Error Resume Next
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 3).Rows(1).Select
End Sub
Private Sub OptionButton4_Click()
    On Error Resume Next
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 4).Rows(1).Select
End Sub
Private Sub OptionButton5_Click()
    On Error Resume Next
    If S_ReBank.Theo_Bai Then
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 5).Rows(1).Select
    Else
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 1) - 1) + 1).Rows(1).Select
    End If
End Sub
Private Sub OptionButton6_Click()
    On Error Resume Next
    If S_ReBank.Theo_Bai Then
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 6).Rows(1).Select
    Else
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 1) - 1) + 2).Rows(1).Select
    End If
End Sub
Private Sub OptionButton7_Click()
    On Error Resume Next
    If S_ReBank.Theo_Bai Then
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 7).Rows(1).Select
    Else
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 1) - 1) + 3).Rows(1).Select
    End If
End Sub
Private Sub OptionButton8_Click()
    On Error Resume Next
    If S_ReBank.Theo_Bai Then
    ActiveDocument.Tables(8 * (Right(S_ReBank.boxBai.text, 1) - 1) + 8).Rows(1).Select
    Else
    ActiveDocument.Tables(4 * (Right(S_ReBank.boxBai.text, 1) - 1) + 4).Rows(1).Select
    End If
End Sub
Private Sub boxBai_DropButtonClick()
On Error GoTo Thoat
    If boxBai.ListCount = 0 Then
        If S_ReBank.Theo_Bai Then
            For i = 1 To ActiveDocument.Tables.Count / 8
                S_ReBank.boxBai.AddItem "Bài " & i
            Next
            S_ReBank.boxBai.Value = "Bài 1"
        Else
            For i = 1 To ActiveDocument.Tables.Count / 4
                If i < 10 Then
                S_ReBank.boxBai.AddItem "Dang 0" & i
                Else
                S_ReBank.boxBai.AddItem "Dang " & i
                End If
            Next
            S_ReBank.boxBai.Value = "Dang 01"
        End If
    End If
Exit Sub
Thoat:
Dim title2, msg As String
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a m" & ChrW(7903) & " ngân hàng câu h" & ChrW(7887) & "i"
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
End Sub

Private Sub O10_Click()
ktlop = 10
Call Brfile
End Sub
Private Sub O11_Click()
ktlop = 11
Call Brfile
End Sub
Private Sub O12_Click()
ktlop = 12
Call Brfile
End Sub


Private Sub TheoBai_Click()
O10.Value = False
O11.Value = False
O12.Value = False
S_List.Clear
End Sub

Private Sub TheoChuyenDe_Click()
O10.Value = False
O11.Value = False
O12.Value = False
S_List.Clear
End Sub

Private Sub Theo_Bai_Click()
    boxBai.Clear
    ktCD = False
    S_ReBank.O10 = False
    S_ReBank.O11 = False
    S_ReBank.O12 = False
    S_ReBank.S_List.Clear
    OptionButton1.Enabled = True
    OptionButton2.Enabled = True
    OptionButton3.Enabled = True
    OptionButton4.Enabled = True
    FrBai.Caption = "Sô câu Bai"
    FrTong.Caption = "Sô câu Chuong"
End Sub

Private Sub Theo_CD_Click()
    boxBai.Clear
    ktCD = True
    S_ReBank.O10 = False
    S_ReBank.O11 = False
    S_ReBank.O12 = False
    S_ReBank.S_List.Clear
    OptionButton1.Enabled = False
    OptionButton2.Enabled = False
    OptionButton3.Enabled = False
    OptionButton4.Enabled = False
    FrBai.Caption = "Sô câu Dang"
    FrTong.Caption = "Sô câu CD"
End Sub

Private Sub TimCauTrong_Click()
Dim dem As Integer
    dem = 0
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "^p^p"
            .Replacement.text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        For i = 1 To ActiveDocument.Tables.Count
            For j = 1 To ActiveDocument.Tables(i).Rows.Count
                 If ActiveDocument.Tables(i).Cell(j, 2).Range.Characters.Count = 1 Then
                    ActiveDocument.Tables(i).Cell(j, 2).Select
                    Selection.Range.HighlightColorIndex = wdYellow
                    Selection.TypeText text:="KHONG CHUA DU LIEU"
                    dem = dem + 1
                 End If
            Next j
        Next i
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "S" & ChrW(7889) & " ô tr" & ChrW(7889) & "ng tìm th" & ChrW(7845) & "y là " & dem & "."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub

Private Sub TimCauTrung_Click()
Dim k, socautrung As Integer
Dim txt, md As String
Dim myRange As Range
Dim ktDel, ktTrung, ktTrung2 As Boolean
Dim ktMsg As Byte
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & ChrW(417) & "ng trình s" & _
        ChrW(7869) & " tìm nh" & ChrW(7919) & "ng câu trùng nhau." & Chr(13) & _
    "N" & ChrW(7871) & "u ch" & ChrW(7885) & _
        "n Yes ch" & ChrW(432) & ChrW(417) & "ng trình s" & ChrW(7869) & _
        " xóa các câu trùng nhau." & Chr(13) & _
    "N" & ChrW(7871) & "u ch" & ChrW(7885) & _
        "n No ch" & ChrW(432) & ChrW(417) & "ng trình s" & ChrW(7869) & " t" & _
        ChrW(7841) & "o ra file m" & ChrW(7899) & "i thông báo các câu trùng nhau."
    ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 4, 0, 2, 0)
    If ktMsg = 6 Then ktDel = True
    If ktMsg = 7 Then ktDel = False
    If ktMsg = 2 Then Exit Sub
txt = ""
socautrung = 0
S_Wait.Show
For k = 1 To ActiveDocument.Tables.Count
    For i = 1 To ActiveDocument.Tables(k).Rows.Count Step 6
        For j = i + 6 To ActiveDocument.Tables(k).Rows.Count Step 6
            ktTrung = False
            ktTrung2 = False
            If ActiveDocument.Tables(k).Cell(i + 3, 2).Range.InlineShapes.Count > 0 _
            And ActiveDocument.Tables(k).Cell(j + 3, 2).Range.InlineShapes.Count > 0 Then
                If ActiveDocument.Tables(k).Cell(i + 3, 2).Range.InlineShapes(1).Width = _
                ActiveDocument.Tables(k).Cell(j + 3, 2).Range.InlineShapes(1).Width _
                And ActiveDocument.Tables(k).Cell(i + 3, 2).Range.InlineShapes(1).Height = _
                ActiveDocument.Tables(k).Cell(j + 3, 2).Range.InlineShapes(1).Height Then
                    ktTrung = True
                End If
            End If
            
            If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.InlineShapes.Count > 0 _
            And ActiveDocument.Tables(k).Cell(j + 1, 2).Range.InlineShapes.Count > 0 Then
                If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.InlineShapes(1).Width = _
                ActiveDocument.Tables(k).Cell(j + 1, 2).Range.InlineShapes(1).Width _
                And ActiveDocument.Tables(k).Cell(i + 1, 2).Range.InlineShapes(1).Height = _
                ActiveDocument.Tables(k).Cell(j + 1, 2).Range.InlineShapes(1).Height Then
                    ktTrung2 = True
                End If
            End If
            
            If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.text = _
                ActiveDocument.Tables(k).Cell(j + 1, 2).Range.text _
                And ActiveDocument.Tables(k).Cell(i + 2, 2).Range.text = _
                ActiveDocument.Tables(k).Cell(j + 2, 2).Range.text _
                And ktTrung And ktTrung2 Then
                If Theo_Bai Then
                    Select Case k Mod 8
                    Case 1
                    md = ".MD1.LT"
                    Case 2
                    md = ".MD2.LT"
                    Case 3
                    md = ".MD3.LT"
                    Case 4
                    md = ".MD4.LT"
                    Case 5
                    md = ".MD1.BT"
                    Case 6
                    md = ".MD2.BT"
                    Case 7
                    md = ".MD3.BT"
                    Case 0
                    md = ".MD4.BT"
                    End Select
                Else
                    Select Case k Mod 4
                    Case 1
                    md = ".MD1"
                    Case 2
                    md = ".MD2"
                    Case 3
                    md = ".MD3"
                    Case 0
                    md = ".MD4"
                    End Select
                End If
                If ktDel Then
                    socautrung = socautrung + 1
                    Set myRange = ActiveDocument.Range( _
                    Start:=ActiveDocument.Tables(k).Cell(j + 1, 1).Range.Start, _
                    End:=ActiveDocument.Tables(k).Cell(j + 6, 2).Range.End)
                    myRange.Select
                    Selection.Cells.Delete ShiftCells:=wdDeleteCellsShiftLeft
                Else
                    If Theo_Bai Then
                        txt = txt & "[B" & Int(k / 8) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & (j - 1) / 6 + 1 & "]  "
                    Else
                        txt = txt & "[Dang " & Int(k / 4) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & (j - 1) / 6 + 1 & "]  "
                    End If
                End If
            End If
        Next j
    Next i
Next k
S_Wait.Hide
If ktDel = True Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = ChrW(272) & "ã xóa " & socautrung & " câu trùng nhau."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    End If
If ktDel = False And txt = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Không có câu trùng nào."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    End If
If ktDel = False And txt <> "" Then
    Documents.add
    Call S_PageSetup
    Selection.TypeText text:=txt
End If
End Sub

Private Sub TimCauTuongTu_Click()
Dim k, socautrung As Integer
Dim txt, md, md2 As String
Dim myRange As Range
Dim ktDel, ktTrung, ktTrung2 As Boolean
Dim ktMsg As Byte
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & ChrW(417) & "ng trình s" & _
        ChrW(7869) & " tìm nh" & ChrW(7919) & "ng câu t" & ChrW(432) & ChrW(417) & "ng t" & ChrW(7921) & " nhau." & Chr(13) & _
    "N" & ChrW(7871) & "u ch" & ChrW(7885) & _
        "n Yes ch" & ChrW(432) & ChrW(417) & "ng trình s" & ChrW(7869) & _
        " chuy" & ChrW(7875) & "n các câu t" & ChrW(432) & _
         ChrW(417) & "ng t" & ChrW(7921) & " v" & ChrW(7873) & " g" & _
        ChrW(7847) & "n nhau." & Chr(13) & _
    "N" & ChrW(7871) & "u ch" & ChrW(7885) & _
        "n No ch" & ChrW(432) & ChrW(417) & "ng trình s" & ChrW(7869) & " t" & _
        ChrW(7841) & "o ra file m" & ChrW(7899) & "i thông báo các câu t" & ChrW(432) & ChrW(417) & "ng t" & ChrW(7921) & " nhau."
    ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 4, 0, 2, 0)
    If ktMsg = 6 Then ktDel = True
    If ktMsg = 7 Then ktDel = False
    If ktMsg = 2 Then Exit Sub
txt = ""
socautrung = 0
S_ReBank.Hide
S_Wait.Show
For k = 1 To ActiveDocument.Tables.Count
    For i = 1 To ActiveDocument.Tables(k).Rows.Count Step 6
        For j = i + 6 To ActiveDocument.Tables(k).Rows.Count Step 6
            If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.text = _
                ActiveDocument.Tables(k).Cell(j + 1, 2).Range.text Then
                If Theo_Bai Then
                    Select Case k Mod 8
                    Case 1
                    md = ".MD1.LT"
                    Case 2
                    md = ".MD2.LT"
                    Case 3
                    md = ".MD3.LT"
                    Case 4
                    md = ".MD4.LT"
                    Case 5
                    md = ".MD1.BT"
                    Case 6
                    md = ".MD2.BT"
                    Case 7
                    md = ".MD3.BT"
                    Case 0
                    md = ".MD4.BT"
                    End Select
                Else
                    Select Case k Mod 4
                    Case 1
                    md = ".MD1"
                    Case 2
                    md = ".MD2"
                    Case 3
                    md = ".MD3"
                    Case 0
                    md = ".MD4"
                    End Select
                End If
                If ktDel And j > i + 1 Then
                    socautrung = socautrung + 1
                    Set myRange = ActiveDocument.Range( _
                    Start:=ActiveDocument.Tables(k).Cell(j + 1, 1).Range.Start, _
                    End:=ActiveDocument.Tables(k).Cell(j + 6, 2).Range.End)
                    myRange.Select
                    myRange.Copy
                    Selection.Cells.Delete ShiftCells:=wdDeleteCellsShiftLeft
                    ActiveDocument.Tables(k).Cell(i + 6, 1).Select
                    Selection.InsertRowsBelow 6
                    Selection.Paste
                Else
                    If Theo_Bai Then
                        txt = txt & "[B" & Int(k / 8) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                        "B" & Int(k / 8) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                    Else
                        txt = txt & "[D" & Int(k / 4) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                        "D" & Int(k / 4) + 1 & md & ":" & (j - 1) / 6 + 1 & "]  "
                    End If
                End If
            End If
        Next j
    Next i
    ''''''
    If S_ReBank.CungMD Then GoTo Tiep
    ''''''
    If k Mod 4 < 4 And k Mod 4 > 0 Then
    If ActiveDocument.Tables(k + 1).Rows.Count > 5 Then
        For i = 1 To ActiveDocument.Tables(k).Rows.Count Step 6
            For j = 1 To ActiveDocument.Tables(k + 1).Rows.Count Step 6
                If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.text = _
                    ActiveDocument.Tables(k + 1).Cell(j + 1, 2).Range.text Then
                    If Theo_Bai Then
                        Select Case k Mod 8
                        Case 1
                        md = ".MD1.LT"
                        md2 = ".MD2.LT"
                        Case 2
                        md = ".MD2.LT"
                        md2 = ".MD3.LT"
                        Case 3
                        md = ".MD3.LT"
                        md2 = ".MD4.LT"
                        Case 4
                        md = ".MD4.LT"
                        Case 5
                        md = ".MD1.BT"
                        md2 = ".MD2.BT"
                        Case 6
                        md = ".MD2.BT"
                        md2 = ".MD3.BT"
                        Case 7
                        md = ".MD3.BT"
                        md2 = ".MD4.BT"
                        Case 0
                        md = ".MD4.BT"
                        End Select
                    Else
                        Select Case k Mod 4
                        Case 1
                        md = ".MD1"
                        md2 = ".MD2"
                        Case 2
                        md = ".MD2"
                        md2 = ".MD3"
                        Case 3
                        md = ".MD3"
                        md2 = ".MD4"
                        Case 0
                        md = ".MD4"
                        End Select
                    End If
                    If ktDel And j > i + 1 Then
                        socautrung = socautrung + 1
                        Set myRange = ActiveDocument.Range( _
                        Start:=ActiveDocument.Tables(k + 1).Cell(j + 1, 1).Range.Start, _
                        End:=ActiveDocument.Tables(k + 1).Cell(j + 6, 2).Range.End)
                        myRange.Select
                        myRange.Copy
                        Selection.Cells.Delete ShiftCells:=wdDeleteCellsShiftLeft
                        ActiveDocument.Tables(k).Cell(i + 6, 1).Select
                        Selection.InsertRowsBelow 6
                        Selection.Paste
                    Else
                        If Theo_Bai Then
                            txt = txt & "[B" & Int(k / 8) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                            "B" & Int(k / 8) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                        Else
                            txt = txt & "[D" & Int(k / 4) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                            "D" & Int(k / 4) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                        End If
                    End If
                End If
            Next j
        Next i
    End If
    End If
    '''''
    '''''
    If k Mod 4 < 3 And k Mod 4 > 1 Then
    If ActiveDocument.Tables(k + 2).Rows.Count > 5 Then
        For i = 1 To ActiveDocument.Tables(k).Rows.Count Step 6
            For j = 1 To ActiveDocument.Tables(k + 2).Rows.Count Step 6
                If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.text = _
                    ActiveDocument.Tables(k + 2).Cell(j + 1, 2).Range.text Then
                    If Theo_Bai Then
                        Select Case k Mod 8
                        Case 1
                        md = ".MD1.LT"
                        md2 = ".MD3.LT"
                        Case 2
                        md2 = ".MD2.LT"
                        md = ".MD4.LT"
                        Case 3
                        md = ".MD3.LT"
                        Case 4
                        md = ".MD4.LT"
                        Case 5
                        md = ".MD1.BT"
                        md2 = ".MD3.BT"
                        Case 6
                        md = ".MD2.BT"
                        md2 = ".MD4.BT"
                        Case 7
                        md = ".MD3.BT"
                        Case 0
                        md = ".MD4.BT"
                        End Select
                    Else
                        Select Case k Mod 4
                        Case 1
                        md = ".MD1"
                        md2 = ".MD3"
                        Case 2
                        md = ".MD2"
                        md2 = ".MD4"
                        Case 3
                        md = ".MD3"
                        Case 0
                        md = ".MD4"
                        End Select
                    End If
                    If ktDel And j > i + 1 Then
                        socautrung = socautrung + 1
                        Set myRange = ActiveDocument.Range( _
                        Start:=ActiveDocument.Tables(k + 2).Cell(j + 1, 1).Range.Start, _
                        End:=ActiveDocument.Tables(k + 2).Cell(j + 6, 2).Range.End)
                        myRange.Select
                        myRange.Copy
                        Selection.Cells.Delete ShiftCells:=wdDeleteCellsShiftLeft
                        ActiveDocument.Tables(k).Cell(i + 6, 1).Select
                        Selection.InsertRowsBelow 6
                        Selection.Paste
                    Else
                        If Theo_Bai Then
                            txt = txt & "[B" & Int(k / 8) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                            "B" & Int(k / 8) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                        Else
                            txt = txt & "[D" & Int(k / 4) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                            "D" & Int(k / 4) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                        End If
                    End If
                End If
            Next j
        Next i
    End If
    End If
    '''''
    '''''
    If k Mod 4 < 2 And k Mod 4 > 0 Then
    If ActiveDocument.Tables(k + 3).Rows.Count > 5 Then
        For i = 1 To ActiveDocument.Tables(k).Rows.Count Step 6
            For j = 1 To ActiveDocument.Tables(k + 3).Rows.Count Step 6
                If ActiveDocument.Tables(k).Cell(i + 1, 2).Range.text = _
                    ActiveDocument.Tables(k + 3).Cell(j + 1, 2).Range.text Then
                    If Theo_Bai Then
                        Select Case k Mod 8
                        Case 1
                        md = ".MD1.LT"
                        md2 = ".MD4.LT"
                        Case 2
                        md = ".MD2.LT"
                        Case 3
                        md = ".MD3.LT"
                        Case 4
                        md = ".MD4.LT"
                        Case 5
                        md = ".MD1.BT"
                        md2 = ".MD4.BT"
                        Case 6
                        md = ".MD2.BT"
                        Case 7
                        md = ".MD3.BT"
                        Case 0
                        md = ".MD4.BT"
                        End Select
                    Else
                        Select Case k Mod 4
                        Case 1
                        md = ".MD1"
                        md2 = ".MD4"
                        Case 2
                        md = ".MD2"
                        Case 3
                        md = ".MD3"
                        Case 0
                        md = ".MD4"
                        End Select
                    End If
                    If ktDel And j > i + 1 Then
                        socautrung = socautrung + 1
                        Set myRange = ActiveDocument.Range( _
                        Start:=ActiveDocument.Tables(k + 3).Cell(j + 1, 1).Range.Start, _
                        End:=ActiveDocument.Tables(k + 3).Cell(j + 6, 2).Range.End)
                        myRange.Select
                        myRange.Copy
                        Selection.Cells.Delete ShiftCells:=wdDeleteCellsShiftLeft
                        ActiveDocument.Tables(k).Cell(i + 6, 1).Select
                        Selection.InsertRowsBelow 6
                        Selection.Paste
                    Else
                        If Theo_Bai Then
                            txt = txt & "[B" & Int(k / 8) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                            "B" & Int(k / 8) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                        Else
                            txt = txt & "[D" & Int(k / 4) + 1 & md & ":" & (i - 1) / 6 + 1 & "~" & _
                            "D" & Int(k / 4) + 1 & md2 & ":" & (j - 1) / 6 + 1 & "]  "
                        End If
                    End If
                End If
            Next j
        Next i
    End If
    End If
    '''''
Tiep:
Next k
S_Wait.Hide
S_ReBank.Show
If ktDel = True Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = ChrW(272) & "ã chuy" & ChrW(7875) & "n " & socautrung & " câu t" & ChrW(432) & ChrW(417) & "ng t" & ChrW(7921) & " nhau."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    End If
If ktDel = False And txt = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Không có câu t" & ChrW(432) & ChrW(417) & "ng t" & ChrW(7921) & "nào."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    End If
If ktDel = False And txt <> "" Then
    Documents.add
    Call S_PageSetup
    Selection.TypeText text:=txt
End If
End Sub

Private Sub UserForm_Initialize()
Call CheckDrive
CungMD = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If S_text = "" Then Exit Sub
Dim Tb As String
Tb = MsgBox("Ban muôn chuân hóa lai file <" & S_text & "> không?", vbYesNo, _
    "B&T Program Created by Le Hoai Son")
    If Tb = vbYes Then
        S_Wait.Label1.Visible = True
        S_Wait.Show
        'ChuanhoaBank (addinBank & S_text)
        S_Wait.Hide
    End If
If CloseMode = vbFormControlMenu Then Cancel = False
End Sub

Private Sub XoaFile_Click()
Dim ktMsg As Byte
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n mu" & ChrW(7889) & "n xóa file này?"
    ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 2, 0, 2, 0)
    If ktMsg = 6 Then
        Kill (S_Drive & "S_Bank&Test\S_Bank\Lop " & ktlop & "\" & S_ReBank.S_List.Value)
    End If
    Call Brfile
End Sub
