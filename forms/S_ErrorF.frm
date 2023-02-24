VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_ErrorF 
   Caption         =   "B&T Program Created by Le Hoai Son"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "S_ErrorF.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "S_ErrorF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ListBox1_Click()
Dim j As Byte
If ListBox1.Value <> "" Then
j = Mid(ListBox1.Value, 5, 2)
Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & "q"
End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim Tb As String
Tb = MsgBox("Ban muôn thoát không?", vbYesNo, "Thông báo")
    If Tb = vbYes Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

