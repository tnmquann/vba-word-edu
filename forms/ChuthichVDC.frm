VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChuthichVDC 
   Caption         =   " "
   ClientHeight    =   2865
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5595
   OleObjectBlob   =   "ChuthichVDC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChuthichVDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox1_Click()
If ChuthichVDC.CheckBox1 = True Then
    ChuthichVDC.OptionButton1.Enabled = True
    ChuthichVDC.OptionButton2.Enabled = True
    ChuthichVDC.OptionButton3.Enabled = True
Else
    ChuthichVDC.OptionButton1.Enabled = False
    ChuthichVDC.OptionButton2.Enabled = False
    ChuthichVDC.OptionButton3.Enabled = False
End If
End Sub
Private Sub CommandButton2_Click()
    TextBox1 = ""
    ChuthichVDC.Hide
    End
End Sub
Private Sub CommandButton1_Click()
    If TextBox1 = "" Then
        msg = "B" & ChrW(7841) & "n vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p t" & ChrW(234) & "n ngu" & ChrW(7891) & "n c" & ChrW(7847) & "n ch" & ChrW(250) & " th" & ChrW(237) & "ch cho c" & ChrW(226) & "u h" & ChrW(7887) & "i." & vbCrLf & "N" & ChrW(7871) & "u b" & ChrW(7841) & "n kh" & ChrW(244) & "ng mu" & ChrW(7889) & "n ghi ch" & ChrW(250) & " th" & ChrW(237) & "ch n" & ChrW(7919) & "a h" & ChrW(227) & "y ch" & ChrW(7885) & "n " & ChrW(8220) & "Hu" & ChrW(7927) & "" & ChrW(8221)
        Application.Assistant.DoAlert "Hý" & ChrW(7899) & "ng d" & ChrW(7851) & "n", msg, 0, 4, 0, 0, 0
    Else
    ' Doan code yeu cau nhap ten nguon file có nhieu ky tu mot chut
    If Len(TextBox1) < 10 Then
    msg2 = "B" & ChrW(7841) & "n " & ChrW(417) & "i. B" & ChrW(7841) & "n vui l" & ChrW(242) & "ng vi" & ChrW(7871) & "t t" & ChrW(234) & "n c" & ChrW(7911) & "a ngu" & ChrW(7891) & "n " & ChrW(273) & "" & ChrW(7873) & " d" & ChrW(224) & "i ra th" & ChrW(234) & "m m" & ChrW(7897) & "t t" & ChrW(237) & " n" & ChrW(7919) & "a nh" & ChrW(233) & "."
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg2, 0, 4, 0, 0, 0
    Else
    ChuthichVDC.Hide
    Application.Run MacroName:="Ghi_CT"
    End If
    End If
    End
End Sub
Private Sub Ghi_CT()
    On Error Resume Next
    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    Tmp = ChuthichVDC.TextBox1
' Doan code nay dung de bo bot khoang trang 2 ben tu khoa
    For i = 1 To Len(Tmp)
    If Mid(Tmp, i, 1) = " " Then
    n = n + 1
    Else
    If n > 0 Then Tmp = Right(Tmp, Len(Tmp) - n)
    Exit For
    End If
    Next i
    n = 0
    For i = 1 To Len(Tmp)
    If Mid(Tmp, Len(Tmp) - i + 1, 1) = " " Then
    n = n + 1
    Else
    If n > 0 Then Tmp = Left(Tmp, Len(Tmp) - n)
    Exit For
    End If
    Next i
' Doan code dung de do cau hoi chua tu khoa va to mau highlight
    Selection.HomeKey Unit:=wdStory
    If ChuthichVDC.OptionButton1 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .ClearFormatting
            .text = "(Câu [0-9]{1,4}[.:])"
            .Replacement.text = "\1" & "^9[" & Tmp & "]^13"
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .text = "(^13)([^9^32])"
            .Replacement.text = "\1"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    Else
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .ClearFormatting
            .text = "(Câu [0-9]{1,4}[.:])"
            .Replacement.text = "\1" & "^9[" & Tmp & "] "
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .text = "( [^9^32])"
            .Replacement.text = " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End If
    Call To_mau_CT
    ActiveDocument.UndoClear
    ChuthichVDC.TextBox1 = ""
    Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
End Sub
Private Sub To_mau_CT()
On Error Resume Next
Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "\[" & "([0-2][D-H][1-6]-[1-4])" & "\]"
        .Replacement.text = "#" & "\1" & "~"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 192
    End With
    With Selection.Find
        .text = "\[" & "(*)" & "\]"
        .Replacement.text = "[" & "\1" & "]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "#" & "([0-2][D-H][1-6]-[1-4])" & "~"
        .Replacement.text = "[" & "\1" & "]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.UndoClear
End Sub
Private Sub Label3_Click()

End Sub
