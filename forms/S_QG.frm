VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_QG 
   Caption         =   "Questions Group"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   OleObjectBlob   =   "S_QG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "S_QG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
Call Taonhom(S_QG.ComboBox1.Value, S_QG.CheckBox1)
Unload S_QG
End Sub

Private Sub Label3_Click()
Call Taonhom(S_QG.ComboBox1.Value, S_QG.CheckBox1)
Unload S_QG
End Sub

Private Sub UserForm_Initialize()
ComboBox1.list = Array("1", "2", "3", "4", "5")
End Sub

