VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_inf 
   Caption         =   "B&T Program Created by Le Hoai Son"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "S_inf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "S_inf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim ktinput1 As Boolean
Dim ale As String
Dim ktCao As Boolean
Dim ktMsg As Byte
Dim Title, msg As String
Private Sub CommandButton1_Click()
If ktinput1 Or S_inf.t1 = "" Then
    S_inf.Lin3.Visible = True
Else
    S_inf.Hide
    Call S_Mark(S_inf.t1)
End If
End Sub

Private Sub Label15_Click()
If ktCao Then
S_inf.Height = 257
ktCao = False
Else
S_inf.Height = 393
ktCao = True
End If
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n mu" & ChrW(7889) & "n xóa file này?"
    ktMsg = Application.Assistant.DoAlert(Title, msg, 4, 2, 0, 0, 1)
    If ktMsg = 6 Then
        If ktlop = 13 Then
        Kill (S_Drive & "S_Bank&Test\S_Data\Other\" & ListBox1.Value & ".dat")
        Else
        Kill (S_Drive & "S_Bank&Test\S_Data\Lop " & ktlop & "\" & ListBox1.Value & ".dat")
        End If
         t1 = ""
        Call Brfile
    End If
End Sub
Private Sub OptionButton4_Click()
S_inf.Lin1.Visible = False
ktlop = 13
Call Brfile
End Sub
Private Sub OptionButton3_Click()
S_inf.Lin1.Visible = False
ktlop = 12
Call Brfile
End Sub
Private Sub OptionButton2_Click()
S_inf.Lin1.Visible = False
ktlop = 11
Call Brfile
End Sub
Private Sub OptionButton1_Click()
S_inf.Lin1.Visible = False
ktlop = 10
Call Brfile
End Sub

Private Sub LayFile(ByVal ThuMuc As String)
Dim f As String
If Right(ThuMuc, 1) <> "\" Then ThuMuc = ThuMuc & "\"
f = Dir$(ThuMuc & "*.dat")
ListBox1.Clear
While Len(f)
    ListBox1.AddItem Left(f, Len(f) - 4)
    f = Dir$
Wend
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
Private Sub ListBox1_Click()
Dim tam() As String
    tam = Split(ListBox1.Value, " [")
    t1 = tam(0)
End Sub

Private Sub OptionButtondriveC_Click()
S_Drive = "C:\"
Call Brfile
End Sub
Private Sub OptionButtondriveD_Click()
S_Drive = "D:\"
Call Brfile
End Sub

Private Sub t1_Change()
Dim tam() As String
If t1 <> "" And ListBox1.Value <> "" Then
tam = Split(ListBox1.Value, " [")
If (FExists(S_Drive & "S_Bank&Test\S_Data\Lop 10\" & t1 & " [" & tam(1) & ".dat") And S_inf.OptionButton1 = True) _
    Or (FExists(S_Drive & "S_Bank&Test\S_Data\Lop 11\" & t1 & " [" & tam(1) & ".dat") And S_inf.OptionButton2 = True) _
    Or (FExists(S_Drive & "S_Bank&Test\S_Data\Lop 12\" & t1 & " [" & tam(1) & ".dat") And S_inf.OptionButton3 = True) _
    Or (FExists(S_Drive & "S_Bank&Test\S_Data\Other\" & t1 & " [" & tam(1) & ".dat") And S_inf.OptionButton4 = True) Then
S_inf.Lin1.Visible = True
ktinput1 = False
Else
S_inf.Lin1.Visible = False
ktinput1 = False
End If
End If
End Sub
Private Sub UserForm_Initialize()
Call CheckDrive
Select Case S_Drive
Case "C:\"
OptionButtonDriveC.Value = True
Case "D:\"
OptionButtonDriveD.Value = True
Case Else
S_Drive = "D:\"
End Select
S_inf.Lin3.Visible = False
OptionButtonDriveD.Enabled = False
OptionButtonDriveC.Enabled = False
End Sub
