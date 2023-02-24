VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_Pages 
   Caption         =   "AnswerSheets"
   ClientHeight    =   690
   ClientLeft      =   5115
   ClientTop       =   10065
   ClientWidth     =   9735.001
   OleObjectBlob   =   "S_Pages.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "S_Pages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_A5.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub

Private Sub Label2_Click()
Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_NH.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub

Private Sub Label3_Click()
Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_50.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub

Private Sub Label4_Click()
Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_120.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub

Private Sub UserForm_Initialize()
Call CheckDrive
End Sub
