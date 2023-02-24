VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_matran 
   Caption         =   "B&T Program Created by Le Hoai Son"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16905
   OleObjectBlob   =   "S_matran.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "S_matran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim dong As Byte
Dim i As Integer
Dim sobai_C1, sobai_C2, sobai_C3, sobai_C4, sobai_C5, sobai_C6, sobai_C7, sobai_C8, sobai_C9 As Integer
Dim SoBai() As Byte
Dim kt2W As Boolean
Dim Title, msg As String
Dim ttcau As Integer
Private Sub Browers_Start()
Dim www As New Word.Application
Dim bank As New Word.Document
Dim tabNum() As Byte
Dim tt, j, Chuong As Byte
If ktlop = 0 Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n l" & ChrW(7899) & "p."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
Exit Sub
End If
If S_matran.ComboMon.Value = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n môn."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
Exit Sub
End If
Dim docOpener As Document
If docIsOpen("PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") Then
            Set docOpener = Application.Documents("PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")
            docOpener.Close
            Set docOpener = Nothing
End If
Call CheckDrive

    If FExists(S_Drive & "S_Bank&Test\S_Templates\PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "PPCT môn này ch" & ChrW(432) & "a có."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        M1_C1.Enabled = False
        M1_C2.Enabled = False
        M1_C3.Enabled = False
        M1_C4.Enabled = False
        M1_C5.Enabled = False
        M1_C6.Enabled = False
        M1_C7.Enabled = False
        M1_C8.Enabled = False
        M1_C9.Enabled = False
    Exit Sub
    End If
    Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")
    Chuong = 0
    Chuong = bank.Tables().Count
Select Case Chuong
Case 1
    M1_C1.Enabled = True
    M1_C2.Enabled = False
    M1_C3.Enabled = False
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 2
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = False
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 3
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 4
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 5
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 6
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 7
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = True
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 8
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = True
    M1_C8.Enabled = True
    M1_C9.Enabled = False
Case 9
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = True
    M1_C8.Enabled = True
    M1_C9.Enabled = True
Case Else
    M1_C1.Enabled = False
    M1_C2.Enabled = False
    M1_C3.Enabled = False
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
End Select
www.Quit
    M1_C1.Value = False
    M1_C2.Value = False
    M1_C3.Value = False
    M1_C4.Value = False
    M1_C5.Value = False
    M1_C6.Value = False
    M1_C7.Value = False
    M1_C8.Value = False
    M1_C9.Value = False
End Sub
Private Sub Browers_CD_Start()
Dim www As New Word.Application
Dim bank As New Word.Document
Dim tabNum() As Byte
'Dim sobai() As Byte
Dim tt, j, Chuong As Byte
If ktlop = 0 Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n l" & ChrW(7899) & "p."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
Exit Sub
End If
If S_matran.ComboMon.Value = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n môn."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
Exit Sub
End If
Dim docOpener As Document
If docIsOpen("ChDe" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") Then
            Set docOpener = Application.Documents("ChDe" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")
            docOpener.Close
            Set docOpener = Nothing
End If
Call CheckDrive

    If FExists(S_Drive & "S_Bank&Test\S_Templates\ChDe" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Chuyên dê môn này ch" & ChrW(432) & "a có."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        M1_C1.Enabled = False
        M1_C2.Enabled = False
        M1_C3.Enabled = False
        M1_C4.Enabled = False
        M1_C5.Enabled = False
        M1_C6.Enabled = False
        M1_C7.Enabled = False
        M1_C8.Enabled = False
        M1_C9.Enabled = False
    Exit Sub
    End If
    Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\ChDe" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")
    Chuong = 0
    Chuong = bank.Tables(1).Rows.Count - 1
Select Case Chuong
Case 1
    M1_C1.Enabled = True
    M1_C2.Enabled = False
    M1_C3.Enabled = False
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 2
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = False
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 3
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 4
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 5
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 6
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 7
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = True
    M1_C8.Enabled = False
    M1_C9.Enabled = False
Case 8
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = True
    M1_C8.Enabled = True
    M1_C9.Enabled = False
Case 9
    M1_C1.Enabled = True
    M1_C2.Enabled = True
    M1_C3.Enabled = True
    M1_C4.Enabled = True
    M1_C5.Enabled = True
    M1_C6.Enabled = True
    M1_C7.Enabled = True
    M1_C8.Enabled = True
    M1_C9.Enabled = True
Case Else
    M1_C1.Enabled = False
    M1_C2.Enabled = False
    M1_C3.Enabled = False
    M1_C4.Enabled = False
    M1_C5.Enabled = False
    M1_C6.Enabled = False
    M1_C7.Enabled = False
    M1_C8.Enabled = False
    M1_C9.Enabled = False
End Select
www.Quit
    M1_C1.Value = False
    M1_C2.Value = False
    M1_C3.Value = False
    M1_C4.Value = False
    M1_C5.Value = False
    M1_C6.Value = False
    M1_C7.Value = False
    M1_C8.Value = False
    M1_C9.Value = False
End Sub
Private Sub Browers()
Dim www As New Word.Application
Dim bank As New Word.Document
Dim tt As Byte
Dim docOpener As Document
If docIsOpen("PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") Then
            Set docOpener = Application.Documents("PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")
            docOpener.Close
            Set docOpener = Nothing
End If
Call CheckDrive

    If FExists(S_Drive & "S_Bank&Test\S_Templates\PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "PPCT môn này ch" & ChrW(432) & "a có."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
    End If
    Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")

    ListBox1.Clear
    tt = 0
    sobai_C1 = 0
    sobai_C2 = 0
    sobai_C3 = 0
    sobai_C4 = 0
    sobai_C5 = 0
    sobai_C6 = 0
    sobai_C7 = 0
    sobai_C8 = 0
    sobai_C9 = 0
If M1_C1.Value = True Then
    If www.ActiveDocument.Tables.Count > 0 Then
        For i = 2 To www.ActiveDocument.Tables(1).Rows.Count
            www.ActiveDocument.Tables(1).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        bank.Close
        M1_C1.Value = False
        www.Quit
        Exit Sub
    End If
sobai_C1 = i - 2
End If
'MsgBox sobai_C1
If M1_C2.Value = True Then
    If www.ActiveDocument.Tables.Count > 1 Then
        For i = 2 To www.ActiveDocument.Tables(2).Rows.Count
            www.ActiveDocument.Tables(2).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(2).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        bank.Close
        M1_C2.Value = False
        www.Quit
        Exit Sub
    End If
sobai_C2 = i - 2
End If
If M1_C3.Value = True Then
    If www.ActiveDocument.Tables.Count > 2 Then
        For i = 2 To www.ActiveDocument.Tables(3).Rows.Count
            www.ActiveDocument.Tables(3).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
           
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(3).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
            
        Next i
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        bank.Close
        M1_C3.Value = False
        www.Quit
        Exit Sub
    End If
sobai_C3 = i - 2
End If

If M1_C4.Value = True Then
    If www.ActiveDocument.Tables.Count > 3 Then
        For i = 2 To www.ActiveDocument.Tables(4).Rows.Count
            www.ActiveDocument.Tables(4).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(4).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C4.Value = False
    Exit Sub
    End If
sobai_C4 = i - 2
End If
If M1_C5.Value = True Then
    If www.ActiveDocument.Tables.Count > 4 Then
        For i = 2 To www.ActiveDocument.Tables(5).Rows.Count
            www.ActiveDocument.Tables(5).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(5).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C5.Value = False
    Exit Sub
    End If
sobai_C5 = i - 2
End If
If M1_C6.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For i = 2 To www.ActiveDocument.Tables(6).Rows.Count
            www.ActiveDocument.Tables(6).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
          
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(6).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C6.Value = False
    Exit Sub
    End If
sobai_C6 = i - 2
End If
If M1_C7.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For i = 2 To www.ActiveDocument.Tables(6).Rows.Count
            www.ActiveDocument.Tables(7).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
          
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(6).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C7.Value = False
    Exit Sub
    End If
sobai_C7 = i - 2
End If
If M1_C8.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For i = 2 To www.ActiveDocument.Tables(6).Rows.Count
            www.ActiveDocument.Tables(8).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
          
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(6).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C8.Value = False
    Exit Sub
    End If
sobai_C8 = i - 2
End If
If M1_C9.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For i = 2 To www.ActiveDocument.Tables(6).Rows.Count
            www.ActiveDocument.Tables(9).Cell(i, 2).Select
            www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
          
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".LT]"
                ListBox1.AddItem Left(www.Selection, Len(www.Selection) - 1) & ".BT]"
                www.ActiveDocument.Tables(6).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                ListBox1.list(tt + 1, 1) = www.Selection
                tt = tt + 2
        Next i
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C9.Value = False
    Exit Sub
    End If
sobai_C9 = i - 2
End If
'''''
bank.Close
www.Quit
End Sub

Private Sub Browers_CD()
If S_matran.ComboMon = "" Then Exit Sub
Dim www As New Word.Application
Dim bank As New Word.Document
Dim tabNum() As Byte
Dim tt, j, Chuong As Byte

Dim docOpener As Document
If docIsOpen("PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") Then
            Set docOpener = Application.Documents("PPCT" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")
            docOpener.Close
            Set docOpener = Nothing
End If
Call CheckDrive

    If FExists(S_Drive & "S_Bank&Test\S_Templates\ChDe" & ktlop & "_" & S_matran.ComboMon.Value & ".docx") = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Chuyên dê môn này ch" & ChrW(432) & "a có."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
    End If
    Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\ChDe" & ktlop & "_" & S_matran.ComboMon.Value & ".docx")

    Chuong = bank.Tables(1).Rows.Count - 1
    ReDim tabNum(Chuong) As Byte
    ReDim SoBai(Chuong) As Byte
    tabNum(1) = 2
    For i = 2 To Chuong
        tabNum(i) = tabNum(i - 1) + Val(bank.Tables(1).Cell(i, 4).Range.text)
    Next i
    For i = 1 To Chuong
        SoBai(i) = Val(bank.Tables(1).Cell(i + 1, 4).Range.text)
    Next i
    'MsgBox sobai(1) & sobai(2)
'Exit Sub
    ListBox1.Clear
    tt = 0
    sobai_C1 = 0
    sobai_C2 = 0
    sobai_C3 = 0
    sobai_C4 = 0
    sobai_C5 = 0
    sobai_C6 = 0
    sobai_C7 = 0
    sobai_C8 = 0
    sobai_C9 = 0
    'Dim i1 As Integer
If M1_C1.Value = True Then
    If www.ActiveDocument.Tables.Count > 0 Then
        For j = 1 To SoBai(1)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C1 = sobai_C1 + i - 2
        Next j
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        bank.Close
        M1_C1.Value = False
        www.Quit
        Exit Sub
    End If
    'MsgBox sobai_C1
End If

If M1_C2.Value = True Then
    If www.ActiveDocument.Tables.Count > SoBai(1) Then
        For j = SoBai(1) + 1 To SoBai(1) + SoBai(2)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C2 = sobai_C2 + i - 2
        Next j
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        bank.Close
        M1_C2.Value = False
        www.Quit
        Exit Sub
    End If
    'MsgBox sobai_C2
End If

If M1_C3.Value = True Then
    If www.ActiveDocument.Tables.Count > SoBai(1) + SoBai(2) Then
         For j = SoBai(1) + SoBai(2) + 1 To SoBai(1) + SoBai(2) + SoBai(3)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C3 = sobai_C3 + i - 2
        Next j
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        bank.Close
        M1_C3.Value = False
        www.Quit
        Exit Sub
    End If
    'MsgBox sobai_C3
End If
If M1_C4.Value = True Then
    If www.ActiveDocument.Tables.Count > 3 Then
        For j = SoBai(1) + SoBai(2) + SoBai(3) + 1 To SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C4 = sobai_C4 + i - 2
        Next j
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C4.Value = False
    Exit Sub
    End If
    'MsgBox sobai_C4
End If
If M1_C5.Value = True Then
    If www.ActiveDocument.Tables.Count > 4 Then
        For j = SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4) + 1 To SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4) + SoBai(5)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C5 = sobai_C5 + i - 2
        Next j
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C5.Value = False
    Exit Sub
    End If
End If
If M1_C6.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For j = SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4) + SoBai(5) + 1 To SoBai(1) + SoBai(2) _
        + SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C6 = sobai_C6 + i - 2
        Next j
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C6.Value = False
    Exit Sub
    End If
End If
If M1_C7.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For j = SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6) + 1 To SoBai(1) + SoBai(2) + _
        SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6) + SoBai(7)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C7 = sobai_C7 + i - 2
        Next j
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C7.Value = False
    Exit Sub
    End If
End If
If M1_C8.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For j = SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6) + SoBai(7) + 1 To SoBai(1) + SoBai(2) + _
        SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6) + SoBai(7) + SoBai(8)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C8 = sobai_C8 + i - 2
        Next j
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C8.Value = False
    Exit Sub
    End If
End If
If M1_C9.Value = True Then
    If www.ActiveDocument.Tables.Count > 5 Then
        For j = SoBai(1) + SoBai(2) + SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6) + SoBai(7) + SoBai(8) + 1 To SoBai(1) + SoBai(2) + _
        SoBai(3) + SoBai(4) + SoBai(5) + SoBai(6) + SoBai(7) + SoBai(8) + SoBai(9)
            For i = 2 To www.ActiveDocument.Tables(j + 1).Rows.Count
                www.ActiveDocument.Tables(j + 1).Cell(i, 2).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.AddItem www.Selection
                www.ActiveDocument.Tables(j + 1).Cell(i, 3).Select
                www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                ListBox1.list(tt, 1) = www.Selection
                tt = tt + 1
            Next i
            sobai_C9 = sobai_C9 + i - 2
        Next j
    Else
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & ChrW(417) & "ng này không t" & ChrW(7891) & "n t" & ChrW(7841) & "i."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    bank.Close
    www.Quit
    M1_C9.Value = False
    Exit Sub
    End If
End If
'''
bank.Close
www.Quit
End Sub
Private Sub socau(ByRef add As String)
    
    Dim www As New Word.Application
    Dim bank As New Word.Document
    Dim docOpener As Document
    Dim tendoc() As String
    tendoc = Split(add, "\")
    If docIsOpen(tendoc(4)) Then
            Set docOpener = Application.Documents(tendoc(4))
            docOpener.Close
            Set docOpener = Nothing
    End If
        Set bank = www.Documents.Open(add, PasswordDocument:="159")
       
            For i = 1 To www.ActiveDocument.Tables.Count - 3 Step 8
                ListBox1.list(ttcau, 2) = Int((www.ActiveDocument.Tables(i).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 4) = Int((www.ActiveDocument.Tables(i + 1).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 6) = Int((www.ActiveDocument.Tables(i + 2).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 8) = Int((www.ActiveDocument.Tables(i + 3).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 3) = "0"
                ListBox1.list(ttcau, 5) = "0"
                ListBox1.list(ttcau, 7) = "0"
                ListBox1.list(ttcau, 9) = "0"
                ttcau = ttcau + 1
                ListBox1.list(ttcau, 2) = Int((www.ActiveDocument.Tables(i + 4).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 4) = Int((www.ActiveDocument.Tables(i + 5).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 6) = Int((www.ActiveDocument.Tables(i + 6).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 8) = Int((www.ActiveDocument.Tables(i + 7).Rows.Count - 1) / 6)
                ListBox1.list(ttcau, 3) = "0"
                ListBox1.list(ttcau, 5) = "0"
                ListBox1.list(ttcau, 7) = "0"
                ListBox1.list(ttcau, 9) = "0"
                ttcau = ttcau + 1
            Next i
    bank.Close (False)
    www.Quit (False)
End Sub
Private Sub socauCD(ByRef listIndex As Byte)
    Dim tt As Byte
    Dim www As New Word.Application
    Dim bank As New Word.Document
    Dim docOpener As Document
    Dim add_txt As String
For tt = listIndex To listIndex + 4
    add_txt = "Chuyen de\" & Mid(S_matran.ListBox1.list(tt, 0), 2, 7) & "\" & Left(S_matran.ListBox1.list(tt, 0), 14)
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        
        If docIsOpen(Left(S_matran.ListBox1.list(tt, 0), 14) & "].dat") Then
                Set docOpener = Application.Documents(Left(S_matran.ListBox1.list(tt, 0), 14) & "].dat")
                docOpener.Close (False)
                Set docOpener = Nothing
        End If
        Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat", PasswordDocument:="159")
            ListBox1.list(tt, 2) = Int((www.ActiveDocument.Tables(1).Rows.Count - 1) / 6)
            ListBox1.list(tt, 4) = Int((www.ActiveDocument.Tables(2).Rows.Count - 1) / 6)
            ListBox1.list(tt, 6) = Int((www.ActiveDocument.Tables(3).Rows.Count - 1) / 6)
            ListBox1.list(tt, 8) = Int((www.ActiveDocument.Tables(4).Rows.Count - 1) / 6)
            ListBox1.list(tt, 3) = "0"
            ListBox1.list(tt, 5) = "0"
            ListBox1.list(tt, 7) = "0"
            ListBox1.list(tt, 9) = "0"
        bank.Close (False)
    Else
        ListBox1.list(tt, 2) = "0"
        ListBox1.list(tt, 4) = "0"
        ListBox1.list(tt, 6) = "0"
        ListBox1.list(tt, 8) = "0"
        ListBox1.list(tt, 3) = "0"
        ListBox1.list(tt, 5) = "0"
        ListBox1.list(tt, 7) = "0"
        ListBox1.list(tt, 9) = "0"
    End If
    If tt >= S_matran.ListBox1.ListCount - 1 Then GoTo S_Quit
Next tt
S_Quit:
    www.Quit (False)
End Sub
Private Sub listTrong()
        ListBox1.list(ttcau, 2) = "0"
        ListBox1.list(ttcau, 4) = "0"
        ListBox1.list(ttcau, 6) = "0"
        ListBox1.list(ttcau, 8) = "0"
        ListBox1.list(ttcau, 3) = "0"
        ListBox1.list(ttcau, 5) = "0"
        ListBox1.list(ttcau, 7) = "0"
        ListBox1.list(ttcau, 9) = "0"
        ttcau = ttcau + 1
    If S_matran.Theo_Bai Then
        ListBox1.list(ttcau, 2) = "0"
        ListBox1.list(ttcau, 4) = "0"
        ListBox1.list(ttcau, 6) = "0"
        ListBox1.list(ttcau, 8) = "0"
        ListBox1.list(ttcau, 3) = "0"
        ListBox1.list(ttcau, 5) = "0"
        ListBox1.list(ttcau, 7) = "0"
        ListBox1.list(ttcau, 9) = "0"
        ttcau = ttcau + 1
    End If
End Sub
Private Sub KTsocau()
On Error GoTo S_Quit
    Dim add_txt As String
    Dim daload As Byte
    daload = 0
S_Wait.Show
Dim S_Khode As String
If S_matran.Kho1 Then S_Khode = "S_Bank"
If S_matran.Kho2 Then S_Khode = "S_Bank 2"
If S_matran.Kho3 Then S_Khode = "S_Bank 3"
If M1_C1.Value = True Then
    
    add_txt = Left(S_matran.ListBox1.list(0, 0), 8)
    'MsgBox add_txt
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
        For i = 1 To sobai_C1
            Call listTrong
        Next i
    End If
    daload = daload + sobai_C1
End If

If M1_C2.Value = True Then

    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
        For i = 1 To sobai_C2
            Call listTrong
        Next i
    End If
    daload = daload + sobai_C2
End If
If M1_C3.Value = True Then
 
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)

    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C3
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C3
End If
If M1_C4.Value = True Then
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)

    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C4
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C4
End If
If M1_C5.Value = True Then
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)
    
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C5
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C5
End If
If M1_C6.Value = True Then
  
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)
    
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C6
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C6
End If
If M1_C7.Value = True Then
  
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)
    
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C7
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C7
End If
If M1_C8.Value = True Then
  
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C8
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C8
End If
If M1_C9.Value = True Then
  
    add_txt = Left(S_matran.ListBox1.list(2 * daload + 1, 0), 8)
    
    If FExists(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat") Then
        Call socau(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & ktlop & "\" & add_txt & "].dat")
    Else
    For i = 1 To sobai_C9
        Call listTrong
    Next i
    End If
    daload = daload + sobai_C9
End If
S_Wait.Hide
Dim tcmd1, tcmd2, tcmd3, tcmd4 As Integer
For i = 0 To ListBox1.ListCount - 1
tcmd1 = tcmd1 + Val(ListBox1.list(i, 2))
tcmd2 = tcmd2 + Val(ListBox1.list(i, 4))
tcmd3 = tcmd3 + Val(ListBox1.list(i, 6))
tcmd4 = tcmd4 + Val(ListBox1.list(i, 8))
Next i
S_matran.tongcaumd = "[" & tcmd1 & "][" & tcmd2 & "][" & tcmd3 & "][" & tcmd4 & "]"
S_matran.tongcau = tcmd1 + tcmd2 + tcmd3 + tcmd4
Exit Sub
S_Quit:
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Có lôi"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub
Private Sub KTsocau_CD()
Dim b As Byte
On Error GoTo S_Quit
Dim add_txt As String
Dim tcmd1, tcmd2, tcmd3, tcmd4 As Integer
S_Wait.Show
Dim S_Khode As String
If S_matran.Kho1 Then S_Khode = "S_Bank"
If S_matran.Kho2 Then S_Khode = "S_Bank 2"
If S_matran.Kho3 Then S_Khode = "S_Bank 3"

    For b = 0 To S_matran.ListBox1.ListCount - 2 Step 5
        Call socauCD(b)
    Next b

For i = 0 To ListBox1.ListCount - 1
tcmd1 = tcmd1 + Val(ListBox1.list(i, 2))
tcmd2 = tcmd2 + Val(ListBox1.list(i, 4))
tcmd3 = tcmd3 + Val(ListBox1.list(i, 6))
tcmd4 = tcmd4 + Val(ListBox1.list(i, 8))
Next i
S_matran.tongcaumd = "[" & tcmd1 & "][" & tcmd2 & "][" & tcmd3 & "][" & tcmd4 & "]"
S_matran.tongcau = tcmd1 + tcmd2 + tcmd3 + tcmd4
S_Wait.Hide
Exit Sub
S_Quit:
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Khong load duoc file"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub
Private Sub butKT_Click()
ttcau = 0
If Theo_Bai Then
    Call KTsocau
Else
    Call KTsocau_CD
End If
End Sub

Private Sub ComboMon_Change()
    Theo_Bai = False
    Theo_CD = False
    'Call Browers
End Sub

Private Sub CommandButton1_Click()
Call S_SerialHDD
If ktBanQuyen = False Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(7881) & " có b" & ChrW(7843) & _
        "n FULL m" & ChrW(7899) & "i s" & ChrW(7917) & " d" & ChrW(7909) & "ng " _
        & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" & ChrW(7913) & "c n" & ChrW _
        (259) & "ng này. M" & ChrW(7901) & "i b" & ChrW(7841) & _
        "n liên h" & ChrW(7879) & " tác gi" & ChrW(7843) & " " & ChrW(273) & ChrW _
        (7875) & " " & ChrW(273) & ChrW(259) & "ng ký."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
If S_matran.TextMon = "" Or Val(S_matran.Ltong) = 0 Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Thông tin ch" & ChrW(432) & "a " & ChrW(273) & _
        ChrW(7847) & "y " & ChrW(273) & ChrW(7911) & "."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
Else
    For i = 0 To S_matran.ListBox1.ListCount - 1
    If Val(ListBox1.list(i, 2)) < Val(ListBox1.list(i, 3)) Then GoTo Tiep
    If Val(ListBox1.list(i, 4)) < Val(ListBox1.list(i, 5)) Then GoTo Tiep
    If Val(ListBox1.list(i, 6)) < Val(ListBox1.list(i, 7)) Then GoTo Tiep
    If Val(ListBox1.list(i, 8)) < Val(ListBox1.list(i, 9)) Then GoTo Tiep
    Next i
    If ktInTheoBai = True And ktInTheoCD = False Then
        Call S_BankNew
    ElseIf ktInTheoBai = False And ktInTheoCD = True Then
        Call S_BankNew_CD
    Else
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & _
        "n in theo Ch" & ChrW(432) & ChrW(417) & "ng bài hay Chuyên " & ChrW(273) & ChrW(7873) & "."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    End If
End If
Exit Sub
Tiep:
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "S" & ChrW(7889) & " câu " & ChrW(273) & "ã ch" & _
         ChrW(7885) & "n v" & ChrW(432) & ChrW(7907) & "t quá s" & ChrW(7889) & _
        " câu " & ChrW(273) & "ã có trong ngân hàng " & ChrW(273) & ChrW(7873) & _
        "."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub
Private Sub kfCauSelect()
    For i = 0 To S_matran.ListBox1.ListCount - 1
    If ListMD1.list(i) < ListMD1_select.list(i) Then GoTo Tiep
    If ListMD2.list(i) < ListMD2_select.list(i) Then GoTo Tiep
    If ListMD3.list(i) < ListMD3_select.list(i) Then GoTo Tiep
    If ListMD4.list(i) < ListMD4_select.list(i) Then GoTo Tiep
    Next i
Tiep:
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "S" & ChrW(7889) & " câu " & ChrW(273) & "ã ch" & _
         ChrW(7885) & "n v" & ChrW(432) & ChrW(7907) & "t quá s" & ChrW(7889) & _
        " câu " & ChrW(273) & "ã có trong ngân hàng " & ChrW(273) & ChrW(7873) & _
        "."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub

Private Sub Kho1_Click()
 S_Khode = "S_Bank"
 TBKhode.Caption = "Kho 01"
End Sub
Private Sub Kho2_Click()
 S_Khode = "S_Bank 2"
 TBKhode.Caption = "Kho 02"
End Sub
Private Sub Kho3_Click()
 S_Khode = "S_Bank 3"
 TBKhode.Caption = "Kho 03"
End Sub

Private Sub Label27_Click()
    S_matran.Height = 487
End Sub

Private Sub Label33_Click()
On Error GoTo S_Quit
    If S_matran.ListBox1.ListCount = 0 Then
        Application.Assistant.DoAlert "", "Ch" & ChrW(432) & "a ki" & ChrW(7875) & _
        "m tra s" & ChrW(7889) & " câu " & ChrW(273) & "ã có trong ngân hàng." _
        , 0, 4, 0, 0, 0
    Exit Sub
    End If
    S_matran.Hide
    If Selection.Tables.Count > 0 Then
        If Selection.Tables(1).Columns.Count = 10 Then
            Selection.EndKey Unit:=wdStory
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            GoTo Tiep
        Else
            Documents.add
        End If
    End If
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(21)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
        
    Dim txt1, txt2 As String
    txt1 = "[DATA" & ktlop & "][" & ComboMon & "][" & IIf(ktCD, "CHUYENDE", "CHUONGBAI") _
    & "][KHO:" & IIf(Right(S_Khode, 1) = "k", "1", Right(S_Khode, 1)) & "]"
    txt2 = "L" & ChrW(7899) & "p " & ktlop & " - Môn : " & ComboMon & " - Qu" & ChrW(7843) & "n lý theo : " & IIf(ktCD, _
    "Chuyên " & ChrW(272) & ChrW(7873), "Ch" & ChrW(432) & ChrW(417) & "ng Bài") _
    & " - Kho " & ChrW(273) & ChrW(7873) & " s" & ChrW(7889) & " " & IIf(Right(S_Khode, 1) = "k", "1", Right(S_Khode, 1))
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText text:="D" & ChrW(7918) & " LI" & ChrW(7878) & _
        "U NGÂN HÀNG CÂU H" & ChrW(7886) & "I - B&T PRO" & Chr(13) & txt2
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.Size = 14
    With Selection.Find
        .text = "D" & ChrW(7918) & " LI" & ChrW(7878) & _
        "U NGÂN HÀNG CÂU H" & ChrW(7886) & "I - B&T PRO"
        .Replacement.text = "D" & ChrW(7918) & " LI" & ChrW(7878) & _
        "U NGÂN HÀNG CÂU H" & ChrW(7886) & "I - B&T PRO"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdPink
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.Size = 14
    With Selection.Find
        .text = txt2
        .Replacement.text = txt2
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    Selection.TypeParagraph
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=10, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Application.Keyboard (1033)
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=70, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=430, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(6).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(7).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(8).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(9).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(10).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustNone
    Selection.TypeText text:="[MÃ BÀI]"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText text:=txt1
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell, Count:=7
Tiep:
    For i = 0 To S_matran.ListBox1.ListCount - 1
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 0)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 1)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 2)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 3)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 4)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 5)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 6)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 7)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 8)
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:=S_matran.ListBox1.list(i, 9)
    Next i
    If Selection.Tables(1).Rows(1).Cells.Count = 6 Then GoTo Tiep2
    Selection.Tables(1).Columns(2).Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Tables(1).Columns(9).Select
    Selection.Shading.BackgroundPatternColor = wdColorLightGreen
    Selection.Tables(1).Columns(7).Select
    Selection.Shading.BackgroundPatternColor = wdColorLightGreen
    Selection.Tables(1).Columns(5).Select
    Selection.Shading.BackgroundPatternColor = wdColorLightGreen
    Selection.Tables(1).Columns(3).Select
    Selection.Shading.BackgroundPatternColor = wdColorLightGreen
    Selection.Tables(1).Rows(1).Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Shading.BackgroundPatternColor = wdColorTurquoise
    Selection.Font.Bold = True
    Selection.Tables(1).Cell(1, 9).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.TypeText text:="[VDC]"
    Selection.Tables(1).Cell(1, 7).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.TypeText text:="[VD]"
    Selection.Tables(1).Cell(1, 5).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.TypeText text:="[HI" & ChrW(7874) & "U]"
    Selection.Tables(1).Cell(1, 3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.TypeText text:="[BI" & ChrW(7870) & "T]"
    kt2W = True
Tiep2:
    Unload S_matran
Exit Sub
S_Quit:
    Application.Assistant.DoAlert "", "T" & ChrW(7841) & "o file m" & ChrW(7899) & _
        "i tr" & ChrW(432) & ChrW(7899) & "c khi xu" & ChrW(7845) & "t." _
        , 0, 4, 0, 0, 0
End Sub

Private Sub Label34_Click()
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "Khong tim thay Matran cac muc do"
        Exit Sub
    End If
    If Selection.Tables(1).Columns.Count <> 10 Then
        MsgBox "Ma tran  khong dung dang"
        Exit Sub
    End If
    ListBox1.Clear
    Dim tt As Byte
    Dim temp() As String
    tt = 0
    temp = Split(Selection.Tables(1).Cell(1, 2).Range.text, "]")
    
    Select Case Right(temp(0), 2)
    Case "10"
    S_matran.OptionButton1 = True
    Case "11"
    S_matran.OptionButton2 = True
    Case "12"
    S_matran.OptionButton3 = True
    End Select
    S_matran.ComboMon = Right(temp(1), 2)
    Select Case Right(temp(2), 2)
    Case "DE"
    S_matran.Theo_CD = True
    Case "AI"
    S_matran.Theo_Bai = True
    End Select
    Select Case Right(temp(3), 1)
    Case "1"
    S_matran.Kho1 = True
    Case "2"
    S_matran.Kho2 = True
    Case "3"
    S_matran.Kho3 = True
    End Select
    For i = 2 To Selection.Tables(1).Rows.Count
        Selection.Tables(1).Cell(i, 1).Select
        If Val(Selection.Tables(1).Cell(i, 4).Range) + Val(Selection.Tables(1).Cell(i, 6).Range) + _
        Val(Selection.Tables(1).Cell(i, 8).Range) + Val(Selection.Tables(1).Cell(i, 10).Range) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            ListBox1.AddItem Selection.text
            ListBox1.list(tt, 1) = Left(Selection.Tables(1).Cell(i, 2).Range, _
            Len(Selection.Tables(1).Cell(i, 2).Range) - 2)
            ListBox1.list(tt, 2) = Val(Selection.Tables(1).Cell(i, 3).Range)
            ListBox1.list(tt, 3) = Val(Selection.Tables(1).Cell(i, 4).Range)
            ListBox1.list(tt, 4) = Val(Selection.Tables(1).Cell(i, 5).Range)
            ListBox1.list(tt, 5) = Val(Selection.Tables(1).Cell(i, 6).Range)
            ListBox1.list(tt, 6) = Val(Selection.Tables(1).Cell(i, 7).Range)
            ListBox1.list(tt, 7) = Val(Selection.Tables(1).Cell(i, 8).Range)
            ListBox1.list(tt, 8) = Val(Selection.Tables(1).Cell(i, 9).Range)
            ListBox1.list(tt, 9) = Val(Selection.Tables(1).Cell(i, 10).Range)
            tt = tt + 1
        End If
    Next i
    If tt = 0 Then
    MsgBox "Chua chon cau hoi nào"
    Else
    Call Ltong_Click
    End If
    
    Dim tcmd1, tcmd2, tcmd3, tcmd4 As Integer
    For i = 0 To ListBox1.ListCount - 1
    tcmd1 = tcmd1 + Val(ListBox1.list(i, 2))
    tcmd2 = tcmd2 + Val(ListBox1.list(i, 4))
    tcmd3 = tcmd3 + Val(ListBox1.list(i, 6))
    tcmd4 = tcmd4 + Val(ListBox1.list(i, 8))
    Next i
    S_matran.tongcaumd = "[" & tcmd1 & "][" & tcmd2 & "][" & tcmd3 & "][" & tcmd4 & "]"
    S_matran.tongcau = tcmd1 + tcmd2 + tcmd3 + tcmd4
End Sub



Private Sub ListBox1_Click()
On Error GoTo S_Quit
S_matran.emd1 = ListBox1.list(ListBox1.listIndex, 2)
S_matran.emd2 = ListBox1.list(ListBox1.listIndex, 4)
S_matran.emd3 = ListBox1.list(ListBox1.listIndex, 6)
S_matran.emd4 = ListBox1.list(ListBox1.listIndex, 8)
S_matran.md1 = ListBox1.list(ListBox1.listIndex, 3)
S_matran.md2 = ListBox1.list(ListBox1.listIndex, 5)
S_matran.md3 = ListBox1.list(ListBox1.listIndex, 7)
S_matran.md4 = ListBox1.list(ListBox1.listIndex, 9)
S_Quit:
End Sub

Private Sub Ltong_Click()
Dim t1, t2, t3, t4, i As Integer
t1 = 0
t2 = 0
t3 = 0
t4 = 0
For i = 1 To ListBox1.ListCount
t1 = t1 + Val(ListBox1.list(i - 1, 3))
t2 = t2 + Val(ListBox1.list(i - 1, 5))
t3 = t3 + Val(ListBox1.list(i - 1, 7))
t4 = t4 + Val(ListBox1.list(i - 1, 9))
Next i
Ltong = t1 + t2 + t3 + t4
If Ltong = "0" Then Exit Sub
PhantramNB = Round((Val(t1) / Val(Ltong)), 3) * 100 & "%"
PhantramTH = Round((Val(t2) / Val(Ltong)), 3) * 100 & "%"
PhantramVD = Round((Val(t3) / Val(Ltong)), 3) * 100 & "%"
PhantramVDC = Round((Val(t4) / Val(Ltong)), 3) * 100 & "%"
PhantramNB.Visible = True
PhantramTH.Visible = True
PhantramVD.Visible = True
PhantramVDC.Visible = True
End Sub

Private Sub M1_C1_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C2_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub

Private Sub M1_C3_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C4_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C5_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C6_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C7_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C8_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub
Private Sub M1_C9_Click()
If Theo_Bai Then
    Call Browers
Else
    Call Browers_CD
End If
End Sub

Private Sub OptionButton1_Click()
ComboMon.text = ""
ktlop = 10

M1_C1.Value = False
M1_C2.Value = False
M1_C3.Value = False
M1_C4.Value = False
M1_C5.Value = False
M1_C6.Value = False
M1_C7.Value = False
M1_C8.Value = False
M1_C9.Value = False
M1_C1.Enabled = False
M1_C2.Enabled = False
M1_C3.Enabled = False
M1_C4.Enabled = False
M1_C5.Enabled = False
M1_C6.Enabled = False
M1_C7.Enabled = False
M1_C8.Enabled = False
M1_C9.Enabled = False

End Sub

Private Sub OptionButton2_Click()
ComboMon.text = ""
ktlop = 11

M1_C1.Value = False
M1_C2.Value = False
M1_C3.Value = False
M1_C4.Value = False
M1_C5.Value = False
M1_C6.Value = False
M1_C7.Value = False
M1_C8.Value = False
M1_C9.Value = False
M1_C1.Enabled = False
M1_C2.Enabled = False
M1_C3.Enabled = False
M1_C4.Enabled = False
M1_C5.Enabled = False
M1_C6.Enabled = False
M1_C7.Enabled = False
M1_C8.Enabled = False
M1_C9.Enabled = False

End Sub
Private Sub OptionButton3_Click()
ComboMon.text = ""
ktlop = 12

M1_C1.Value = False
M1_C2.Value = False
M1_C3.Value = False
M1_C4.Value = False
M1_C5.Value = False
M1_C6.Value = False
M1_C7.Value = False
M1_C8.Value = False
M1_C9.Value = False
M1_C1.Enabled = False
M1_C2.Enabled = False
M1_C3.Enabled = False
M1_C4.Enabled = False
M1_C5.Enabled = False
M1_C6.Enabled = False
M1_C7.Enabled = False
M1_C8.Enabled = False
M1_C9.Enabled = False

End Sub

Private Sub TextMon_Change()
datontai.Visible = False
If TextMon <> "" And ((DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 10\" & TextMon & "\") And OptionButton1 = True) _
    Or (DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 11\" & TextMon & "\") And OptionButton2 = True) _
    Or (DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 12\" & TextMon & "\") And OptionButton3 = True)) Then
datontai.Visible = True
End If
End Sub

Private Sub Theo_Bai_Click()
    Call Browers_Start
    ListBox1.Clear
    inHDG.Enabled = False
    InmaCH.Enabled = True
    ktInTheoBai = True
    ktInTheoCD = False
    ktCD = False
End Sub

Private Sub Theo_CD_Click()
    Call Browers_CD_Start
    ListBox1.Clear
    inHDG.Enabled = True
    InmaCH.Enabled = False
    ktInTheoCD = True
    ktInTheoBai = False
    ktCD = True
    'Call Browers_CD
End Sub

Private Sub Update_Click()
Dim msg As String
If ListBox1.listIndex = -1 Then
    msg = "Ch" & ChrW(7885) & "n m" & ChrW(7897) & "t ch" & _
         ChrW(7911) & " " & ChrW(273) & ChrW(7873) & " trong List."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
If Val(md1) > Val(emd1) Or Val(md2) > Val(emd2) Or Val(md3) > Val(emd3) Or Val(md4) > Val(emd4) Then
    msg = "S" & ChrW(7889) & " câu " & ChrW(273) & "ã ch" & _
         ChrW(7885) & "n v" & ChrW(432) & ChrW(7907) & "t quá s" & ChrW( _
        7889) & " câu " & ChrW(273) & "ã có trong ngân hàng câu h" & ChrW(7887) & "i."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
ListBox1.list(ListBox1.listIndex, 3) = S_matran.md1
ListBox1.list(ListBox1.listIndex, 5) = S_matran.md2
ListBox1.list(ListBox1.listIndex, 7) = S_matran.md3
ListBox1.list(ListBox1.listIndex, 9) = S_matran.md4
Call Ltong_Click
End Sub

Private Sub UserForm_Initialize()
Call CheckDrive
If ktlop = 0 Then ktlop = 12
ComboMon.list = Array("DS", "HH", "LY", "HO", "SI", "SU", "DI", "CD", "TI", "CN")
ComboHead.list = Array("Header 1", "Header 2", "Header 3", "Header 4", "Header 5")
ComboFooter.list = Array("Default", "Footer 1", "Footer 2")
ComboAns.list = Array("Default", "Before", "After")
ComboSode.list = Array("1", "2", "3", "4", "5", "6", "8", "24")
ComboMADE.list = Array("101", "201", "301", "401", "501", "601", "701", "801", "901")
ComboLevel.list = Array("Default", "(1,2)(3,4)", "(1,2)(3)(4)", "(1)(2)(3)(4)")
    M1_C1.Value = False
    M1_C2.Value = False
    M1_C3.Value = False
    M1_C4.Value = False
    M1_C5.Value = False
    M1_C6.Value = False
'Theo_CD.Enabled = False
ListBox1.ColumnWidths = "78,258,28,28,28,28,28,28,28,28"
TextMon.text = ""
Ltong = "0"
kt2W = False
Select Case ktlop
Case 10
OptionButton1.Value = True
Case 11
OptionButton2.Value = True
Case 12
OptionButton3.Value = True
End Select
datontai.Visible = False
'If ktCD Then
    'Theo_Bai.Enabled = False
    'Theo_CD.Enabled = True
    'Theo_CD.Value = True
'Else
    'Theo_CD.Enabled = False
    'Theo_Bai.Enabled = True
    'Theo_Bai.Value = True
'End If
If ktInTheoBai = True And ktInTheoCD = False Then inHDG.Enabled = False
If ktInTheoBai = False And ktInTheoCD = True Then S_matran.InmaCH.Enabled = False
PhantramNB.Visible = False
PhantramTH.Visible = False
PhantramVD.Visible = False
PhantramVDC.Visible = False
Select Case S_Khode
Case "S_Bank 2"
Kho2 = True
Case "S_Bank 3"
Kho3 = True
Case Else
Kho1 = True
End Select
'If DirExists(S_Drive & "S_Bank&Test\S_Bank 2") = False Then
    'Kho2.Enabled = False
'End If
'If DirExists(S_Drive & "S_Bank&Test\S_Bank 3") = False Then
    'Kho3.Enabled = False
'End If
BoxUpdate.Enabled = False
End Sub
