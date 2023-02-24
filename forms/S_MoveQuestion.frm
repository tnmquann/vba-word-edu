VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_MoveQuestion 
   Caption         =   "Move Question"
   ClientHeight    =   1155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745.001
   OleObjectBlob   =   "S_MoveQuestion.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "S_MoveQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    On Error GoTo S_Quit
    Dim md, LT_BT, sodong As Byte
    Dim title2, msg As String
    If Val(Selection.Rows.Count) Mod 6 <> 0 Or Selection.Columns.Count <> 2 Then
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a ch" & _
        ChrW(7885) & "n h" & ChrW(7871) & "t câu."
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
        Exit Sub
    End If
    sodong = Val(Selection.Rows.Count)
    If Left(Bai, 3) <> "Bài" And Theo_Bai Then
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a ch" & ChrW(7885) & _
        "n bài c" & ChrW(7847) & "n chuy" & ChrW(7875) & "n " & ChrW(273) & ChrW( _
        7871) & "n."
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
        Exit Sub
    End If
    If Left(Dang, 3) <> "Dan" And Theo_CD Then
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "B" & ChrW(7841) & "n chua ch" & ChrW(7885) & _
        "n d" & ChrW(7841) & "ng c" & ChrW(7847) & "n chuy" & ChrW(7875) & "n " & ChrW(273) & ChrW( _
        7871) & "n."
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
        Exit Sub
    End If
    Selection.Copy
    Selection.Cells.Delete ShiftCells:=wdDeleteCellsShiftLeft
    If ktCD = False Then
        Select Case Mucdo.text
            Case "Nhan biet"
            md = 0
            Case "Thong hieu"
            md = 1
            Case "Van dung"
            md = 2
            Case "Van dung cao"
            md = 3
        End Select
        Select Case Dang.text
            Case "Ly thuyet"
            LT_BT = 0
            Case "Bai tap"
            LT_BT = 4
        End Select
        ActiveDocument.Tables(8 * (Right(Bai, 1) - 1) + 1 + LT_BT + md).Select
        Selection.Rows(ActiveDocument.Tables(8 * (Right(Bai, 1) - 1) + 1 + LT_BT + md).Rows.Count).Select
        Selection.InsertRowsBelow Val(sodong)
        Selection.Paste
    Else
        Select Case Mucdo.text
            Case "Nhan biet"
            md = 0
            Case "Thong hieu"
            md = 1
            Case "Van dung"
            md = 2
            Case "Van dung cao"
            md = 3
        End Select
        ActiveDocument.Tables(4 * (Right(Dang, 2) - 1) + 1 + md).Select
        Selection.Rows(ActiveDocument.Tables(4 * (Right(Dang, 2) - 1) + 1 + md).Rows.Count).Select
        Selection.InsertRowsBelow Val(sodong)
        Selection.Paste
    End If
Exit Sub
S_Quit:
MsgBox "Thao tác sai!"
End Sub
Private Sub Bai_DropButtonClick()
On Error GoTo Thoat
    If Bai.ListCount = 0 Then
        For i = 1 To ActiveDocument.Tables.Count / 8
            Bai.AddItem "Bài " & i
        Next
    End If
Exit Sub
Thoat:
    Dim title2, msg As String
    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a m" & ChrW(7903) & " ngân hàng câu h" & ChrW(7887) & "i"
    Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
End Sub

Private Sub Theo_Bai_Click()
    'Mucdo.List = Array("Nhan biet", "Thong hieu", "Van dung", "Van dung cao")
    Dang.Clear
    Bai.Enabled = True
    Dang.list = Array("Ly thuyet", "Bai tap")
    Bai.text = "Chon bai"
    ktCD = False
End Sub

Private Sub Theo_CD_Click()
    Dim i As Integer
    'Mucdo.List = Array("Nhan biet", "Thong hieu", "Van dung", "Van dung cao")
    Dang.Clear
    For i = 1 To ActiveDocument.Tables.Count / 4
        If i < 10 Then
            Dang.AddItem "Dang 0" & i
        Else
            Dang.AddItem "Dang " & i
        End If
    Next i
    Bai.Enabled = False
    Dang.text = "Chon dang"
    ktCD = True
End Sub

Private Sub UserForm_Initialize()
'Call CheckDrive
Mucdo.list = Array("Nhan biet", "Thong hieu", "Van dung", "Van dung cao")
End Sub
