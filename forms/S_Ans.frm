VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_Ans 
   Caption         =   "Answers"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   3060
   ClientWidth     =   3705
   OleObjectBlob   =   "S_Ans.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "S_Ans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ktW As Boolean
Private Sub CommandButton1_Click()
Call S_MakeAns
End Sub
Private Sub CommandButton2_Click()
Call Luu_dap_an
End Sub

Private Sub CommandButton3_Click()
Call S_ExportAns
End Sub

Private Sub Label3_Click()
If ktW Then
S_Ans.Width = 196
ktW = False
Else
S_Ans.Width = 286
ktW = True
End If
End Sub
