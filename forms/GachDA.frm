VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GachDA 
   Caption         =   " "
   ClientHeight    =   2925
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5700
   OleObjectBlob   =   "GachDA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GachDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub CheckBox3_Click()

End Sub
Private Sub CommandButton1_Click()
    If CheckBox1 = False And CheckBox2 = False And CheckBox3 = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg1 = "Sao b" & ChrW(7841) & "n kh" & ChrW(244) & "ng cho bi" & ChrW(7871) & "t b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(225) & "nh d" & ChrW(7845) & "u " & ChrW(273) & "" & ChrW(225) & "p " & ChrW(225) & "n b" & ChrW(7857) & "ng c" & ChrW(225) & "ch n" & ChrW(224) & "o." & vbCrLf & "Vui l" & ChrW(242) & "ng ch" & ChrW(7885) & "n c" & ChrW(225) & "ch b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(225) & "nh d" & ChrW(7845) & "u ho" & ChrW(7863) & "c nh" & ChrW(7845) & "p ch" & ChrW(7885) & "n " & ChrW(8220) & "Hu" & ChrW(7927) & "" & ChrW(8221) & " l" & ChrW(7879) & "nh."
        Application.Assistant.DoAlert Title, msg1, 0, 1, 0, 0, 0
    Else
    GachDA.Hide
    NhanhCham.Show
    End If
End Sub
Private Sub CommandButton2_Click()
    GachDA.Hide
    CheckBox1 = False
    CheckBox2 = False
    CheckBox3 = False
End Sub
Private Sub Label1_Click()

End Sub


Private Sub UserForm_Click()

End Sub
