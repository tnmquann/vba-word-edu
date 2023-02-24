VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GanID5_ExTest 
   Caption         =   "Tien ich gan ID6 theo he thong dang toan cua Nhom Toan va Latex (tac gia: Duong Phuoc Sang)"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16995
   OleObjectBlob   =   "GanID5_ExTest.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "GanID5_ExTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub load_combo_tex(ByVal cb_name As String, ByVal cot1 As String, ByVal cot2 As String)
Dim cot As Long, list, bang As Integer
    Select Case cb_name
        Case "ComboBox1": cot = 1
        Case "ComboBox2": cot = 2
        Case "ComboBox3": cot = 1
        Case "ComboBox4": cot = 2
        Case "ComboBox5": cot = 3
    End Select
    If (cot1 = "" And cot2 = "") Or (cot1 = ComboBox1.Value And cot2 = "") Or (cot1 = ComboBox1.Value And cot2 = ComboBox2.Value) Then
        bang = 1
    Else
        bang = ComboBox5.Value + 25
    End If
    ' Tim Table chua PPCT cua Chuong hien hanh
    If cot > 0 Then
        list = prepare_list(bang, cot, cot1, cot2)
        If Not IsEmpty(list) Then Me.Controls(cb_name).list = list
    End If
End Sub

Private Sub ComboBox1_Change()
    If ComboBox1.Value <> "" Then
        ComboBox2.Enabled = True
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox1.BackColor = &HC0FFFF
        ComboBox2.BackColor = &HC0FFFF
        ComboBox3.BackColor = &H8000000F
        ComboBox4.BackColor = &H8000000F
    End If
    load_combo_tex "ComboBox2", ComboBox1.Value, ""
    On Error Resume Next
    ComboBox2.Value = ""
    ComboBox2.SetFocus
    SendKeys "%{Down}"
End Sub

Private Sub ComboBox2_Change()
    If ComboBox2.Value <> "" Then
        ComboBox3.Enabled = True
        ComboBox4.Enabled = False
        ComboBox3.BackColor = &HC0FFFF
        ComboBox4.BackColor = &H8000000F
    End If
    load_combo_tex "ComboBox5", ComboBox1.Value, ComboBox2.Value
    On Error Resume Next
        ComboBox5.listIndex = 0
End Sub

Private Sub ComboBox4_Change()
    Dim msg As String
    If Right(ComboBox4.Value, 8) = "--------" Then
        msg = "B" & ChrW(7841) & "n ch" & ChrW(7885) & "n d" & ChrW(7841) & "ng to" & ChrW(225) & "n l" & ChrW(7841) & "i gi" & ChrW(250) & "p nh" & ChrW(233) & "!"
        Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
        ComboBox4.Value = ""
    End If
End Sub

Private Sub ComboBox5_Change()
    load_combo_tex "ComboBox3", ComboBox2.Value, ""
    On Error Resume Next
    ComboBox3.Value = ""
    ComboBox3.SetFocus
    SendKeys "%{Down}"
End Sub

Private Sub ComboBox3_Change()
    If ComboBox3.Value <> "" Then
        ComboBox4.Enabled = True
        ComboBox4.BackColor = &HC0FFFF
    End If
    load_combo_tex "ComboBox4", ComboBox3.Value, ""
    On Error Resume Next
    ComboBox4.Value = ""
    ComboBox4.SetFocus
    SendKeys "%{Down}"
End Sub

Private Sub CommandButton1_Click()
' Gan ID6
    Dim msg As String, Lop As String, Mon As String, Chuong As String, Bai As String, CdDangtoan As String, Mucdo As String
    Dim i As Integer, TextID As String, NewTextID As String
    ActiveDocument.UndoClear
    If ComboBox1 = "" Or ComboBox2 = "" Or ComboBox3 = "" Or ComboBox4 = "" Or (OptionButton1 = False And OptionButton2 = False And OptionButton3 = False And OptionButton4 = False And OptionButton7 = False) Then
        msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i cung c" & ChrW(7845) & "p cho ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh " & ChrW(273) & "" & ChrW(7847) & "y " & ChrW(273) & "" & ChrW(7911) & " c" & ChrW(225) & "c th" & ChrW(244) & "ng tin" & vbCrLf & "L" & ChrW(7899) & "p - M" & ChrW(244) & "n - Ch" & ChrW(432) & "" & ChrW(417) & "ng - Chuy" & ChrW(234) & "n " & ChrW(273) & "" & ChrW(7873) & " - D" & ChrW(7841) & "ng to" & ChrW(225) & "n - M" & ChrW(7913) & "c " & ChrW(273) & "" & ChrW(7897)
        Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
    Else
        Application.ScreenUpdating = False
                CommandButton5.Top = 43
                CommandButton13.Top = 156
            On Error Resume Next
            ActiveDocument.UndoClear
            ActiveDocument.Range.ListFormat.ConvertNumbersToText
            Lop = Right(ComboBox1.Value, 1)
            If Left(ComboBox1.Value, 1) = "H" Then
                Mon = "H"
            Else
                Mon = "D"
            End If
            Chuong = Left(ComboBox2.Value, 1)
            Bai = Left(ComboBox3.Value, 1)
            CdDangtoan = Left(ComboBox4.Value, 1)
            If OptionButton1 = True Then Mucdo = OptionButton1.Caption
            If OptionButton2 = True Then Mucdo = OptionButton2.Caption
            If OptionButton3 = True Then Mucdo = OptionButton3.Caption
            If OptionButton4 = True Then Mucdo = OptionButton4.Caption
            If OptionButton7 = True Then Mucdo = OptionButton7.Caption
            
            TextID = "%[" & Lop & Mon & Chuong & Mucdo & Bai & "-" & CdDangtoan & "]"
            
            Selection.EndKey Unit:=wdLine
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "\begin{ex}"
                .Replacement.text = ""
                .Forward = False
                .Wrap = wdFindStop
                .MatchWildcards = False
            End With
            If Selection.Find.Execute = True Then
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.TypeText text:=TextID
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.MoveLeft Unit:=wdCharacter, Count:=10, Extend:=wdExtend
                Selection.Font.Bold = True
                Selection.Font.Color = wdColorPink
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .text = "\begin{ex}"
                    .Replacement.text = ""
                    .Forward = True
                    .Wrap = wdFindStop
                    .MatchWildcards = False
                End With
                If Selection.Find.Execute = True Then
                    Selection.MoveDown Unit:=wdLine, Count:=6
                    Selection.MoveUp Unit:=wdLine, Count:=6
                    Selection.EndKey Unit:=wdLine
                    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                End If
            End If
        ComboBox1.BackColor = &HC0FFFF
        ComboBox2.BackColor = &HC0FFFF
        ComboBox3.BackColor = &HC0FFFF
        ComboBox4.BackColor = &HC0FFFF
        Application.ScreenUpdating = True
    End If
    ActiveDocument.UndoClear
End Sub

Private Sub CommandButton13_Click()
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "\begin{ex}"
        .Forward = False
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    '==============
    Dim msg As String, ID As String, Cau As String, Mon As String, Chuong As Integer, Bai As String, Chuyende As Integer, Dangtoan As Integer, Mucdo As String
    Dim ThisDoc As Document, SourceDoc As Document
    Dim bang As Integer, dong As Integer
    On Error Resume Next
    ActiveDocument.UndoClear
    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .text = "(\[[0-2][DH][0-9][YBKGT][1-9]-[0-9]{1,2}\])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.EndKey Unit:=wdLine
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\begin{ex}"
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    If Selection.Find.Execute = True Then
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "(\[[0-2][DH][0-9][YBKGT][1-9]-[0-9]{1,2}\])"
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
        End With
        If Selection.Find.Execute = False Then
            msg = "T" & ChrW(224) & "i li" & ChrW(7879) & "u c" & ChrW(7911) & "a b" & ChrW(7841) & "n kh" & ChrW(244) & "ng c" & ChrW(243) & " c" & ChrW(226) & "u n" & ChrW(224) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c g" & ChrW(7855) & "n ID6 c" & ChrW(7843)
            Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
            Application.ScreenUpdating = True
            Exit Sub
        Else
            Application.ScreenUpdating = True
            Options.DefaultHighlightColorIndex = wdBrightGreen
            Selection.Range.HighlightColorIndex = wdBrightGreen
            If Mid(Selection.text, 2, 2) = "2D" Then
                Mon = "Gi" & ChrW(7843) & "i t" & ChrW(237) & "ch 12"
            Else
                If Mid(Selection.text, 3, 1) = "D" Then
                    Mon = ChrW(272) & ChrW(7841) & "i s" & ChrW(7889) & " 1" & Mid(Selection.text, 2, 1)
                Else
                    Mon = "H" & ChrW(236) & "nh h" & ChrW(7885) & "c" & " 1" & Mid(Selection.text, 2, 1)
                End If
            End If
            Chuong = Mid(Selection.text, 4, 1)
            Bai = Mid(Selection.text, 6, 1)
            Mucdo = Mid(Selection.text, 5, 1)
            If Mid(Selection.text, 9, 1) = "]" Then
                Dangtoan = Mid(Selection.text, 8, 1)
            Else
                Dangtoan = Mid(Selection.text, 8, 2)
            End If
            ComboBox1.Enabled = True
            ComboBox1.Value = Mon
            ComboBox2.listIndex = Chuong - 1
            ComboBox3.listIndex = Bai - 1
            ComboBox4.listIndex = Dangtoan - 1
            ComboBox1.BackColor = 8454016
            ComboBox2.BackColor = 8454016
            ComboBox3.BackColor = 8454016
            ComboBox4.BackColor = 8454016
            ComboBox6.SetFocus
            If Mucdo = "Y" Then OptionButton1.Value = True
            If Mucdo = "B" Then OptionButton2.Value = True
            If Mucdo = "K" Then OptionButton3.Value = True
            If Mucdo = "G" Then OptionButton4.Value = True
            If Mucdo = "T" Then OptionButton7.Value = True
            Selection.MoveDown Unit:=wdLine, Count:=6
            Selection.MoveUp Unit:=wdLine, Count:=6
            Selection.EndKey Unit:=wdLine
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
        End If
    Else
        Exit Sub
    End If
    ActiveDocument.UndoClear
End Sub

Private Sub CommandButton2_Click()
' Sua ID6
    Dim msg As String, Lop As String, Mon As String, Chuong As String, Bai As String, CdDangtoan As String, Mucdo As String
    Dim i As Integer, n As Integer, OldTextID As String, NewTextID As String
    ActiveDocument.UndoClear
    If ComboBox1 = "" Or ComboBox2 = "" Or ComboBox3 = "" Or ComboBox4 = "" Or (OptionButton1 = False And OptionButton2 = False And OptionButton3 = False And OptionButton4 = False And OptionButton7 = False) Then
        msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i cung c" & ChrW(7845) & "p cho ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh " & ChrW(273) & "" & ChrW(7847) & "y " & ChrW(273) & "" & ChrW(7911) & " c" & ChrW(225) & "c th" & ChrW(244) & "ng tin" & vbCrLf & "L" & ChrW(7899) & "p - M" & ChrW(244) & "n - Ch" & ChrW(432) & "" & ChrW(417) & "ng - Chuy" & ChrW(234) & "n " & ChrW(273) & "" & ChrW(7873) & " - D" & ChrW(7841) & "ng to" & ChrW(225) & "n - M" & ChrW(7913) & "c " & ChrW(273) & "" & ChrW(7897)
        Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
    Else
        Application.ScreenUpdating = False
                CommandButton5.Top = 43
                CommandButton13.Top = 156
            n = Len(Selection.text)
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.MoveRight Unit:=wdCharacter, Count:=(n + 2), Extend:=wdExtend
            OldTextID = Selection.text
            n = Len(OldTextID)
            If n < 7 Or (Left(OldTextID, 1) = "[" And Right(OldTextID, 1) = "]") Then
                msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i ch" & ChrW(7885) & "n v" & ChrW(249) & "ng v" & ChrW(259) & "n b" & ChrW(7843) & "n ch" & ChrW(7913) & "a ID c" & ChrW(7847) & "n ch" & ChrW(7881) & "nh s" & ChrW(7917) & "a"
                Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
                Exit Sub
            End If
            For i = 1 To n
                If Left(OldTextID, 1) <> "[" Then
                    If n - i < 6 Then
                        msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i ch" & ChrW(7885) & "n v" & ChrW(249) & "ng v" & ChrW(259) & "n b" & ChrW(7843) & "n ch" & ChrW(7913) & "a ID c" & ChrW(7847) & "n ch" & ChrW(7881) & "nh s" & ChrW(7917) & "a"
                        Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
                        Exit Sub
                    End If
                    OldTextID = Mid(OldTextID, 2, Len(OldTextID) - 1)
                Else
                    Exit For
                End If
            Next i
            n = Len(OldTextID)
            For i = 1 To n
                If Right(OldTextID, 1) <> "]" Then
                    If n - i < 6 Then
                        msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i ch" & ChrW(7885) & "n v" & ChrW(249) & "ng v" & ChrW(259) & "n b" & ChrW(7843) & "n ch" & ChrW(7913) & "a ID c" & ChrW(7847) & "n ch" & ChrW(7881) & "nh s" & ChrW(7917) & "a"
                        Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
                        Exit Sub
                    End If
                    OldTextID = Mid(OldTextID, 1, Len(OldTextID) - 1)
                Else
                    Exit For
                End If
            Next i
            If (Mid(OldTextID, 3, 1) <> "D" And Mid(OldTextID, 3, 1) <> "H") Or (Mid(OldTextID, 5, 1) <> "Y" And Mid(OldTextID, 5, 1) <> "B" And Mid(OldTextID, 5, 1) <> "K" And Mid(OldTextID, 5, 1) <> "G" And Mid(OldTextID, 5, 1) <> "T") Or Mid(OldTextID, 7, 1) <> "-" Then
                MsgBox OldTextID
                msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i ch" & ChrW(7885) & "n v" & ChrW(249) & "ng v" & ChrW(259) & "n b" & ChrW(7843) & "n ch" & ChrW(7913) & "a ID c" & ChrW(7847) & "n ch" & ChrW(7881) & "nh s" & ChrW(7917) & "a"
                Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
                Exit Sub
            End If
            On Error Resume Next
            ActiveDocument.UndoClear
            Lop = Right(ComboBox1.Value, 1)
            If Left(ComboBox1.Value, 1) = "H" Then
                Mon = "H"
            Else
                Mon = "D"
            End If
            Chuong = Left(ComboBox2.Value, 1)
            Bai = Left(ComboBox3.Value, 1)
            CdDangtoan = Left(ComboBox4.Value, 1)
            If OptionButton1 = True Then Mucdo = OptionButton1.Caption
            If OptionButton2 = True Then Mucdo = OptionButton2.Caption
            If OptionButton3 = True Then Mucdo = OptionButton3.Caption
            If OptionButton4 = True Then Mucdo = OptionButton4.Caption
            If OptionButton7 = True Then Mucdo = OptionButton7.Caption
            
            NewTextID = "[" & Lop & Mon & Chuong & Mucdo & Bai & "-" & CdDangtoan & "]"
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = OldTextID
                .Replacement.text = NewTextID
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = False
                .Execute Replace:=wdReplaceOne
            End With
            Selection.Font.Bold = True
            Selection.Font.Color = wdColorPink
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "\begin{ex}"
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
            End With
            If Selection.Find.Execute = True Then
                Selection.MoveDown Unit:=wdLine, Count:=6
                Selection.MoveUp Unit:=wdLine, Count:=6
                Selection.EndKey Unit:=wdLine
                Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            End If
        ComboBox1.BackColor = &HC0FFFF
        ComboBox2.BackColor = &HC0FFFF
        ComboBox3.BackColor = &HC0FFFF
        ComboBox4.BackColor = &HC0FFFF
    End If
    ActiveDocument.UndoClear
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton3_Click()
If OptionButton6 = False Then 'Xoa ID
    On Error Resume Next
    ActiveDocument.UndoClear
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(-[1-4]\])(^9)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "(-[1-4]\])(^32)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[[012][DH][0-9]{1,2}-[0-9]{1,2}.[0-9]{1,2}-[1-4]\])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "(\[[012][DH][0-9]{1,2}-[1-4]\])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = " "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Else
    Dim fd As Office.FileDialog, txtFileName As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Application.ScreenUpdating = False
    With fd
      .AllowMultiSelect = False
      ' Set the title of the dialog box.
      .Title = "Please select the file."
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.add "*.tex", "*.tex"
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox
        Selection.InsertFile FileName:=txtFileName, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
        TextBox1.Value = txtFileName
      End If
   End With
   If TextBox1.Value = "" Then Exit Sub
    CommandButton3.BackColor = &H8000000F
    CommandButton7.BackColor = &HC0FFC0
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^t"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(2)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 14
        .Bold = True
        .Color = wdColorRed
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With Selection.Find
        .text = "\begin{ex}"
        .Replacement.text = "\begin{ex}"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "\end{ex}"
        .Replacement.text = "\end{ex}"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 12
        .Bold = True
        .Color = wdColorBlue
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With Selection.Find
        .text = "\choice"
        .Replacement.text = "\choice"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Size = 12
        .Bold = True
        .Color = wdColorPink
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    With Selection.Find
        .text = "\loigiai{"
        .Replacement.text = "\loigiai{"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = 5287936
    With Selection.Find
        .text = "($*$)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "\begin{ex}"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End If
End Sub

Private Sub CommandButton4_Click()
' BT.Pro và ID6
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[[DH][SH]1[0-2].C[0-9]{1,2}.[0-9]{1,2}.D??.[abcd]\])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    If Selection.Find.Execute = True Then
        Call BTPro_DPS
        CommandButton4.Caption = "ID6-B&T"
    Else
        Call DPS_BTPro
        CommandButton4.Caption = "B&T-ID6"
    End If
    ActiveDocument.UndoClear
End Sub

Private Sub CommandButton5_Click()
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = wdColorPink
    End With
    With Selection.Find
        .text = "(%)(\[[0-2][DH][0-9][YBKGT][1-9]-[0-9]{1,2}\])"
        .Replacement.text = "\1^p\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Private Sub CommandButton6_Click()
' Doc ID6
    Dim msg As String, ID As String, Cau As String, Mon As String, Chuong As Integer, Bai As Integer, Dangtoan As Integer, Mucdo As String
    Dim ThisDoc As Document, SourceDoc As Document
    Dim bang As Integer, dong As Integer
    On Error Resume Next
        CommandButton13.Top = 43
        CommandButton5.Top = 156
    ActiveDocument.UndoClear
    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .text = "(\[[0-2][DH][0-9][YBKGT][1-9]-[0-9]{1,2}\])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.EndKey Unit:=wdLine
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\begin{ex}"
        .Forward = False
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    If Selection.Find.Execute = True Then
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "(\[[0-2][DH][0-9][YBKGT][1-9]-[0-9]{1,2}\])"
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = True
        End With
        If Selection.Find.Execute = False Then
            msg = "T" & ChrW(224) & "i li" & ChrW(7879) & "u c" & ChrW(7911) & "a b" & ChrW(7841) & "n kh" & ChrW(244) & "ng c" & ChrW(243) & " c" & ChrW(226) & "u n" & ChrW(224) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c g" & ChrW(7855) & "n ID6 c" & ChrW(7843)
            Application.Assistant.DoAlert "", msg, 0, 1, 0, 0, 0
            Application.ScreenUpdating = True
            Exit Sub
        Else
            Application.ScreenUpdating = True
            Options.DefaultHighlightColorIndex = wdBrightGreen
            Selection.Range.HighlightColorIndex = wdBrightGreen
            If Mid(Selection.text, 2, 2) = "2D" Then
                Mon = "Gi" & ChrW(7843) & "i t" & ChrW(237) & "ch 12"
            Else
                If Mid(Selection.text, 3, 1) = "D" Then
                    Mon = ChrW(272) & ChrW(7841) & "i s" & ChrW(7889) & " 1" & Mid(Selection.text, 2, 1)
                Else
                    Mon = "H" & ChrW(236) & "nh h" & ChrW(7885) & "c" & " 1" & Mid(Selection.text, 2, 1)
                End If
            End If
            Chuong = Mid(Selection.text, 4, 1)
            Bai = Mid(Selection.text, 6, 1)
            Mucdo = Mid(Selection.text, 5, 1)
            If Mid(Selection.text, 9, 1) = "]" Then
                Dangtoan = Mid(Selection.text, 8, 1)
            Else
                Dangtoan = Mid(Selection.text, 8, 2)
            End If
            ComboBox1.Enabled = True
            ComboBox1.Value = Mon
            ComboBox2.listIndex = Chuong - 1
            ComboBox3.listIndex = Bai - 1
            ComboBox4.listIndex = Dangtoan - 1
            ComboBox1.BackColor = 8454016
            ComboBox2.BackColor = 8454016
            ComboBox3.BackColor = 8454016
            ComboBox4.BackColor = 8454016
            ComboBox6.SetFocus
            If Mucdo = "Y" Then OptionButton1.Value = True
            If Mucdo = "B" Then OptionButton2.Value = True
            If Mucdo = "K" Then OptionButton3.Value = True
            If Mucdo = "G" Then OptionButton4.Value = True
            If Mucdo = "T" Then OptionButton7.Value = True
            Selection.MoveDown Unit:=wdLine, Count:=6
            Selection.MoveUp Unit:=wdLine, Count:=6
            Selection.EndKey Unit:=wdLine
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
        End If
    Else
        Exit Sub
    End If
    ActiveDocument.UndoClear
End Sub

Private Sub CommandButton7_Click()
If OptionButton6 = False Then ' Xoa ID thua
    On Error Resume Next
    ActiveDocument.UndoClear
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(-[1-4]\])(^9)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "(-[1-4]\])(^32)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "(-[1-4]\])(^9)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[[0-2][DH]*-[1-4]\])(\[[0-2][DH][0-9]*\])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(-[1-4]\])"
        .Replacement.text = "\1 "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = " "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Else
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^13"
        .Replacement.text = "^p^t"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    If GanID_ExTest.TextBox1.Value <> "" Then
        ActiveDocument.SaveAs2 GanID_ExTest.TextBox1.Value, FileFormat:= _
        wdFormatText, LockComments:=False, Password:="", AddToRecentFiles:=False, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
         SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False, Encoding:=65001, InsertLineBreaks:=False, AllowSubstitutions:= _
        False, LineEnding:=wdCRLF, CompatibilityMode:=0
    End If
End If
End Sub

Private Sub CommandButton8_Click()
' Mo chuc nang An, hien HDG
    If Me.CommandButton1.Top = 15 Then
        Me.CommandButton1.Top = 130
        Me.CommandButton2.Top = 130
        Me.CommandButton3.Top = 130
        Me.CommandButton4.Top = 130
        Me.CommandButton9.Top = 15
        Me.CommandButton10.Top = 15
        Me.CommandButton11.Top = 15
        Me.CommandButton12.Top = 15
        Me.CommandButton8.BackColor = &HC0FFC0
    Else
        Me.CommandButton1.Top = 15
        Me.CommandButton2.Top = 15
        Me.CommandButton3.Top = 15
        Me.CommandButton4.Top = 15
        Me.CommandButton9.Top = 130
        Me.CommandButton10.Top = 130
        Me.CommandButton11.Top = 130
        Me.CommandButton12.Top = 130
        Me.CommandButton8.BackColor = &H8000000F
    End If
End Sub


Private Sub CommandButton9_Click()
' An het HDG
    Application.ScreenUpdating = False
    On Error Resume Next
    ActiveWindow.ActivePane.View.ShowAll = True
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "([A-D].)"
        .Replacement.text = "\1 "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([A-D].)(  )"
        .Replacement.text = "\1 "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:]*[A-C].*[^9^13^32]D.*^13)"
        .Replacement.text = "!#^p\1#!"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "!#^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
    Selection.EndKey Unit:=wdStory
    Selection.TypeText text:="!#"
    With Selection.Find
        .text = "^13!#"
        .Replacement.text = " !#"
        .Replacement.Font.Size = 1
        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "#!"
        .Replacement.text = "#! "
        .Replacement.Font.Size = 1
        .Forward = False
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "#! #! "
        .Replacement.text = "#!"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "!# !#^13"
        .Replacement.text = "!#^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(#\!*\!#^13)"
        .Replacement.text = "\1"
        .Replacement.Font.Hidden = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    ActiveWindow.ActivePane.View.ShowAll = False
    ActiveDocument.UndoClear
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton10_Click()
' Hien het HDG
    Application.ScreenUpdating = False
    ActiveWindow.ActivePane.View.ShowAll = True
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(#\!*\!#^13)"
        .Replacement.text = "\1"
        .Replacement.Font.Hidden = False
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "#! "
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = " !#"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    ActiveWindow.ActivePane.View.ShowAll = False
    ActiveDocument.UndoClear
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton11_Click()
' An 1 HDG
    Application.ScreenUpdating = False
    On Error Resume Next
    ActiveWindow.ActivePane.View.ShowAll = True
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.EndKey Unit:=wdLine
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    If Selection.Find.Execute = False Then
        With Selection.Find
            .text = "(Câu [0-9]{1,4}[.:])"
            .Replacement.text = ""
            .Forward = False
            .Wrap = wdFindStop
            .MatchWildcards = True
        End With
        If Selection.Find.Execute = False Then
            ActiveWindow.ActivePane.View.ShowAll = False
            Exit Sub
        Else
            Selection.EndKey Unit:=wdStory
            Selection.TypeText text:="!#"
            Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
            Selection.Font.Size = 1
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
    Else
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeText text:="!#"
        Selection.TypeParagraph
        Selection.MoveLeft Unit:=wdCharacter, Count:=3
        Selection.TypeBackspace
        Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        Selection.Font.Size = 1
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
    With Selection.Find
        .text = "!#!#"
        .Replacement.text = "!#"
        .Forward = False
        .Wrap = wdFindStop
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = ""
        .Forward = False
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    With Selection.Find
        .text = "(A.*[B-C].*[^9^13^32]D.*^13)"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#!"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Font.Size = 1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    With Selection.Find
        .text = "#!#!"
        .Replacement.text = "#!"
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    With Selection.Find
        .text = "(#\!*\!#^13)"
        .Replacement.text = "\1"
        .Replacement.Font.Hidden = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceOne
    End With
    ActiveWindow.ActivePane.View.ShowAll = False
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton12_Click()
' Hien 1 HDG
    Application.ScreenUpdating = False
    ActiveWindow.ActivePane.View.ShowAll = True
    Selection.HomeKey Unit:=wdLine
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(#\!*\!#^13)"
        .Replacement.text = "\1"
        .Replacement.Font.Hidden = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceOne
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "#!"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
    With Selection.Find
        .text = "!#"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
    ActiveWindow.ActivePane.View.ShowAll = False
    ActiveDocument.UndoClear
    Application.ScreenUpdating = True
End Sub

Private Sub BTPro_DPS()
On Error Resume Next
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DH])([SH]1)([0-2].C[0-9]{1,2}.[0-9]{1,2}.D??.)"
        .Forward = True
        .Replacement.text = "[" & "\2\4"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.D??.)(a)(\])"
        .Forward = True
        .Replacement.text = "\1" & "1"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.D??.)(b)(\])"
        .Forward = True
        .Replacement.text = "\1" & "2"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.D??.)(c)(\])"
        .Forward = True
        .Replacement.text = "\1" & "3"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.D??.)(d)(\])"
        .Forward = True
        .Replacement.text = "\1" & "4"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DH])([012])(.C)([0-9]{1,2})(.)([0-9]{1,2}.D)([0-9]{2}.)([1234])"
        .Forward = True
        .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "\8" & "-" & "\9" & "]"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DH])([012])(.C)([0-9]{1,2})(.)([0-9]{1,2}.D0)([0-9]{1}.)([1234])"
        .Forward = True
        .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "\8" & "-" & "\9" & "]"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".D0"
        .Forward = True
        .Replacement.text = "."
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".D"
        .Forward = True
        .Replacement.text = "."
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".-"
        .Forward = True
        .Replacement.text = "-"
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
End Sub
Private Sub DPS_BTPro()
On Error Resume Next
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "\[([012])([DH])([0-9]{1,2})-([0-9]{1,2}).([0-9]{1,2})-([1-4])\]"
        .Replacement.text = "[\2H1\1.C\3.\4.D\5.\6]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = "\[DH1([012].C[0-9]{1,2}.)"
        .Replacement.text = "[DS1\1"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".D([0-9]{1}.[1-4]\])"
        .Replacement.text = ".D0\1"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".1]"
        .Replacement.text = ".a]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".2]"
        .Replacement.text = ".b]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".3]"
        .Replacement.text = ".c]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".4]"
        .Replacement.text = ".d]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineNone
        .Color = wdColorPink
    End With
    With Selection.Find
        .text = "(\[[DH]*.[abcd]\])"
        .Replacement.text = "\1"
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    ' Xu ly truong hop dac biet (dang toan so 0) (hinh nhu du thua - ben tren co roi)
    'Selection.HomeKey Unit:=wdStory
    'With Selection.Find
        '.ClearFormatting
        '.text = "(\[??1?.C?.[0-9]{1,2}.)(D0)(.?\])"
        '.Replacement.text = "\1D00\3"
        '.Replacement.ClearFormatting
        '.Wrap = wdFindContinue
        '.MatchWildcards = True
        '.Execute Replace:=wdReplaceAll, Forward:=True
    'End With
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub OptionButton5_Click()
    Me.Hide
    GanID_Word.Show
    GanID_Word.OptionButton5 = True
    'CommandButton3.Caption = "Xo" & ChrW(225) & " ID"
    'CommandButton3.BackColor = &H8000000F
    'CommandButton4.Width = 50
    'CommandButton5.Caption = ChrW(7848) & "n ID"
    'CommandButton7.Caption = "ID th" & ChrW(7915) & "a"
    'CommandButton7.BackColor = &H8000000F
    'CommandButton8.Width = 50
    'Label9.Width = 244
    'Label10.Top = 130
    'Frame1.Left = 719
    'Me.Width = 792
End Sub

Private Sub OptionButton6_Click()
    CommandButton1.Top = 15
    CommandButton2.Top = 15
    CommandButton3.Top = 15
    CommandButton4.Top = 15
    CommandButton9.Top = 130
    CommandButton10.Top = 130
    CommandButton11.Top = 130
    CommandButton12.Top = 130
    CommandButton8.BackColor = &H8000000F
    CommandButton3.Caption = "M" & ChrW(7903) & " .tex"
    CommandButton3.BackColor = &HC0FFC0
    CommandButton4.Width = 0
    CommandButton5.Caption = "M" & ChrW(224) & "u ID6"
    CommandButton7.Caption = "Save .tex"
    CommandButton7.BackColor = &H8000000F
    CommandButton8.Width = 0
    'Frame1.Left = 760
    'Label9.Width = 185
    'Label10.Top = 6
    
    'Me.Width = 864
End Sub

Private Sub OptionButton7_Click()

End Sub

Private Sub UserForm_Initialize()
    ComboBox2.Enabled = False
    ComboBox3.Enabled = False
    ComboBox4.Enabled = False
    OptionButton2 = True
    OptionButton5 = False
    OptionButton6 = True
    load_combo_tex "ComboBox1", "", ""
    ' Kiem tra ID dang la ID6 hay ID cua B&T.Pro
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[[DH][SH]1[0-2].C[1-9].[1-9].D??.[abcd]\])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    If Selection.Find.Execute = True Then
        CommandButton4.Caption = "B&T-ID6"
    Else
        CommandButton4.Caption = "ID6-B&T"
    End If
    ' kiem tra xem co ID an hay khong
    ActiveWindow.ActivePane.View.ShowAll = True
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[[0-2][DH][0-9]-)"
        .Font.Hidden = True
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    If Selection.Find.Execute = True Then
        CommandButton5.Caption = "Hi" & ChrW(7879) & "n ID"
    Else
        CommandButton5.Caption = ChrW(7848) & "n ID"
    End If
    ActiveWindow.ActivePane.View.ShowAll = False
    Application.ScreenUpdating = True
End Sub
