VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_mf 
   Caption         =   "B&T Program Created by Le Hoai Son"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "S_mf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "S_mf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim ktH As Boolean
Dim macl, ktMsg As Byte
Dim Title, msg As String

Private Sub CommandButton1_Click()
Dim mamon As String
Call S_SerialHDD
mamon = S_mf.mf_t1
If (mf_t3 > 3 Or mf_t4 > 20) And ktBanQuyen = False Then
    S_mf.Hide
    S_Free.Show
    Exit Sub
End If
If ktexist1 Or ktexist2 Or S_mf.mf_t1 = "" Or S_mf.mf_t2 = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Thông tin ch" & ChrW(432) & "a " & ChrW(273) & _
        ChrW(7847) & "y " & ChrW(273) & ChrW(7911) & " ho" & ChrW(7863) & _
        "c không chính xác."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
Else
    S_mf.L3.Visible = False
    S_mf.Hide
    Call S_Mix(S_mf.mf_t2, S_mf.mf_t1, S_mf.ComboBox1, S_mf.ComboBox2, S_mf.ComboBox3, S_mf.mf_t4, S_mf.mf_t3)
    If ktMix = True Then
        If ktlop = 13 Then
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "Các " & ChrW(273) & ChrW(7873) & " " & ChrW(273) & "ã l" & ChrW(432) & "u trong" & Chr(13) & _
            S_Drive & "S_Bank&Test\S_Test\Other\" & mamon & "\Made_xxx.docx"
            Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        Else
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "Các " & ChrW(273) & ChrW(7873) & " " & ChrW(273) & "ã l" & ChrW(432) & "u trong" & Chr(13) & _
            S_Drive & "S_Bank&Test\S_Test\Lop " & ktlop & "\" & mamon & "\Made_xxx.docx"
            Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        End If
    End If
End If
S_mf.mf_t1 = ""
If ktBanQuyen = False Then S_NoteRig.Show
End Sub

Private Sub OptionButtondriveC_Click()
S_Drive = "C:\"
Call Brfile
Call BrDir
End Sub
Private Sub OptionButtondriveD_Click()
S_Drive = "D:\"
Call Brfile
Call BrDir
End Sub
Private Sub Label15_Click()
If ktH Then
S_mf.Height = 300
ktH = False
Else
S_mf.Height = 384
ktH = True
End If
End Sub
Private Sub Label17_Click()
macl = 2
Call MakeMade(mf_t3)
Label17.Visible = False
Label18.Visible = True

End Sub

Private Sub Label18_Click()
macl = 3
Call MakeMade(mf_t3)
Label18.Visible = False
Label19.Visible = True
End Sub

Private Sub Label19_Click()
macl = 1
Call MakeMade(mf_t3)
Label19.Visible = False
Label17.Visible = True
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'On Error Resume Next
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n mu" & ChrW(7889) & "n xóa mã này?"
    ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 2, 0, 2, 0)
    If ktMsg = 6 Then
        If ktlop = 13 Then
            DeleteFolder (S_Drive & "S_Bank&Test\S_Test\Other\" & S_mf.ListBox2.Value)
        Else
            DeleteFolder (S_Drive & "S_Bank&Test\S_Test\Lop " & ktlop & "\" & S_mf.ListBox2.Value)
        End If
        Call BrDir
    End If
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n mu" & ChrW(7889) & "n xóa file này?"
    ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 2, 0, 2, 0)
    If ktMsg = 6 Then
        If ktlop = 13 Then
        Kill (S_Drive & "S_Bank&Test\S_Data\Other\" & _
        ListBox1.Value & "[" & ListBox1.list(ListBox1.listIndex, 1) & "]" _
            & ListBox1.list(ListBox1.listIndex, 2) & ListBox1.list(ListBox1.listIndex, 3) & ".dat")
        Else
        Kill (S_Drive & "S_Bank&Test\S_Data\Lop " & ktlop & "\" & _
        ListBox1.Value & "[" & ListBox1.list(ListBox1.listIndex, 1) & "]" _
            & ListBox1.list(ListBox1.listIndex, 2) & ListBox1.list(ListBox1.listIndex, 3) & ".dat")
        End If
        mf_t2 = ""
        Call Brfile
    End If
End Sub
Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim S_code As String
    Dim i As Byte
    
    S_code = InputBox("Nhâp mã dê moi (3 ky tu)", "Thông báo", ListBox3.Value)
    For i = 0 To ListBox3.ListCount - 1
    If ListBox3.list(i) = S_code Then
    ale = MsgBox("Mã dê bi trùng", vbOKOnly, "B&T Program")
    S_code = InputBox("Nhâp mã dê? ", "B&T Program Created by Le Hoai Son", ListBox3.Value)
    End If
    Next i
    
    If Len(S_code) = 3 Then
    ListBox3.list(ListBox3.listIndex) = S_code
    ElseIf Len(S_code) = 0 Then
    ListBox3.list(ListBox3.listIndex) = ListBox3.Value
    Else
    MsgBox "Ma de phai co 3 ky tu!"
    ListBox3.list(ListBox3.listIndex) = ListBox3.Value
    End If
End Sub
Private Sub OptionButton4_Click()
S_mf.mf_t2 = ""
ktlop = 13
Call Brfile
Call BrDir
Call mf_t1_Change
Call mf_t2_Change
S_mf.L3.Visible = False
End Sub
Private Sub OptionButton3_Click()
S_mf.mf_t2 = ""
ktlop = 12
Call Brfile
Call BrDir
Call mf_t1_Change
Call mf_t2_Change
S_mf.L3.Visible = False
End Sub
Private Sub OptionButton2_Click()
S_mf.mf_t2 = ""
ktlop = 11
Call Brfile
Call BrDir
Call mf_t1_Change
Call mf_t2_Change
S_mf.L3.Visible = False
End Sub
Private Sub OptionButton1_Click()
S_mf.mf_t2 = ""
ktlop = 10
Call Brfile
Call BrDir
Call mf_t1_Change
Call mf_t2_Change
S_mf.L3.Visible = False
End Sub
Private Sub mf_t1_Change()
Call BrDir
S_mf.L3.Visible = False

If (DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_mf.mf_t1 & "\") And S_mf.OptionButton1 = True) _
    Or (DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_mf.mf_t1 & "\") And S_mf.OptionButton2 = True) _
    Or (DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_mf.mf_t1 & "\") And S_mf.OptionButton3 = True) _
     Or (DirExists(S_Drive & "S_Bank&Test\S_Test\Other\" & S_mf.mf_t1 & "\") And S_mf.OptionButton4 = True) Then
S_mf.L1 = "(Mã dã tôn tai!)"
'mf_t1.BackColor = &HFFFFC0
ktexist1 = True
Else
S_mf.L1 = ""
ktexist1 = False
End If
If mf_t1.text = "" Then S_mf.L1 = ""
End Sub
Private Sub mf_t2_Change()
S_mf.L3.Visible = False
If (FExists(S_Drive & "S_Bank&Test\S_Data\Lop 10\" & S_mf.mf_t2 & ".dat") = False And S_mf.OptionButton1 = True) _
    Or (FExists(S_Drive & "S_Bank&Test\S_Data\Lop 11\" & S_mf.mf_t2 & ".dat") = False And S_mf.OptionButton2 = True) _
    Or (FExists(S_Drive & "S_Bank&Test\S_Data\Lop 12\" & S_mf.mf_t2 & ".dat") = False And S_mf.OptionButton3 = True) _
    Or (FExists(S_Drive & "S_Bank&Test\S_Data\Other\" & S_mf.mf_t2 & ".dat") = False And S_mf.OptionButton4 = True) Then
S_mf.L2 = "(không tôn tai!)"
ktexist2 = True
Else
S_mf.L2 = ""
ktexist2 = False
End If
If mf_t2.text = "" Then S_mf.L2 = ""
End Sub
Private Sub LayFile(ByVal ThuMuc As String)
Dim f As String
Dim idx As Byte
Dim tenF() As String
idx = 0
If Right(ThuMuc, 1) <> "\" Then ThuMuc = ThuMuc & "\"
f = Dir$(ThuMuc & "*.dat")
ListBox1.Clear
While Len(f)
    tenF = Split(f, "[")
    ListBox1.AddItem tenF(0)
    If UBound(tenF) = 1 Then
        ListBox1.list(idx, 1) = Left(tenF(1), Len(tenF(1)) - 5)
        ListBox1.list(idx, 2) = ""
        ListBox1.list(idx, 3) = ""
    End If
    If UBound(tenF) = 2 Then
        ListBox1.list(idx, 1) = Left(tenF(1), Len(tenF(1)) - 1)
        If tenF(2) = "TL]" Then
        ListBox1.list(idx, 3) = "[" & Left(tenF(2), Len(tenF(2)) - 4)
        ListBox1.list(idx, 2) = ""
        Else
        ListBox1.list(idx, 2) = "[" & Left(tenF(2), Len(tenF(2)) - 4)
        ListBox1.list(idx, 3) = ""
        End If
    End If
    If UBound(tenF) = 3 Then
        ListBox1.list(idx, 1) = Left(tenF(1), Len(tenF(1)) - 1)
        ListBox1.list(idx, 2) = "[" & tenF(2)
        ListBox1.list(idx, 3) = "[" & Left(tenF(3), Len(tenF(3)) - 4)
    End If
    f = Dir$
    idx = idx + 1
Wend
End Sub
Sub ShowFolderList(ByVal ThuMuc As String)
Dim fs, f, f1, s, sf
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(ThuMuc)
Set sf = f.SubFolders
ListBox2.Clear
For Each f1 In sf
ListBox2.AddItem f1.Name
Next
'MsgBox s
End Sub
Private Sub Brfile()
Select Case ktlop
Case 10
Call LayFile(S_Drive & "S_Bank&Test\S_Data\Lop 10")
Case 11
Call LayFile(S_Drive & "S_Bank&Test\S_Data\Lop 11")
Case 12
Call LayFile(S_Drive & "S_Bank&Test\S_Data\Lop 12")
Case 13
Call LayFile(S_Drive & "S_Bank&Test\S_Data\Other")
End Select
End Sub
Private Sub BrDir()
Select Case ktlop
Case 10
Call ShowFolderList(S_Drive & "S_Bank&Test\S_Test\Lop 10")
Case 11
Call ShowFolderList(S_Drive & "S_Bank&Test\S_Test\Lop 11")
Case 12
Call ShowFolderList(S_Drive & "S_Bank&Test\S_Test\Lop 12")
Case 13
Call ShowFolderList(S_Drive & "S_Bank&Test\S_Test\Other")
End Select
End Sub
Private Sub MakeMade(ByVal numMade As String)
Dim i As Byte
Dim made(50) As String
ListBox3.Clear
For i = 1 To Val(numMade)
Select Case macl
Case 1
made(i) = (i Mod 10) & Int(89 * Rnd() + 10)
Case 2
made(i) = (i Mod 5) * 2 & Int(89 * Rnd() + 10)
Case 3
made(1) = "1" & Int(89 * Rnd() + 10)
made(i + 1) = ((i Mod 5) * 2 + 1) & Int(89 * Rnd() + 10)
End Select
ListBox3.AddItem Trim(made(i))
Next
'MsgBox ListBox3.List(0)
End Sub
Private Sub mf_t3_Change()
If mf_t3 > 0 And mf_t3 <= 50 Then
ReDim made(Val(mf_t3)) As String
Call MakeMade(mf_t3)
End If
End Sub
Private Sub ListBox1_Click()
    mf_t2 = ListBox1.Value & "[" & ListBox1.list(ListBox1.listIndex, 1) & "]" _
            & ListBox1.list(ListBox1.listIndex, 2) & ListBox1.list(ListBox1.listIndex, 3)
End Sub


Private Sub UserForm_Initialize()
macl = 1
ComboBox1.list = Array("Header 1", "Header 2", "Header 3", "Header 4", "Header 5")
ComboBox2.list = Array("Default", "Footer 1", "Footer 2")
ComboBox3.list = Array("", "Before", "After")
ComboBox4.list = Array("Default", "In chung voi de")
ListBox1.ColumnWidths = "70,30,33,30"
mf_t4.list = Array("5", "6", "7", "8", "9", "10", "11", "12", _
"13", "14", "15", "20", "25", "30", "35", "40", "45", "50", "60", "70", "100")
mf_t3.list = Array("1", "2", "3", "4", "8", "10", "12", "24")
mf_t5.list = Array("11", "21", "31", "41", "81")
Label18.Visible = False
Label19.Visible = False
Call MakeMade(mf_t3)
Call CheckDrive
Select Case S_Drive
Case "C:\"
OptionButtonDriveC.Value = True
Case "D:\"
OptionButtonDriveD.Value = True
Case Else
S_Drive = "D:\"
End Select
OptionButtonDriveD.Enabled = False
OptionButtonDriveC.Enabled = False
        Select Case ktlop
            Case 10
                S_mf.OptionButton1 = True
            Case 11
                S_mf.OptionButton2 = True
            Case 12
                S_mf.OptionButton3 = True
            Case 13
                S_mf.OptionButton4 = True
        End Select
End Sub

