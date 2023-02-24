VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} keysearch 
   Caption         =   " "
   ClientHeight    =   2520
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6870
   OleObjectBlob   =   "keysearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "keysearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
Public Sub TextBox1_Change()

End Sub
Private Sub CommandButton2_Click()
    TextBox1 = ""
    CheckBox1 = False
    keysearch.Hide
End Sub
Private Sub CommandButton1_Click()
    If TextBox1 = "" Then
        msg = "N" & ChrW(7871) & "u v" & ChrW(7851) & "n mu" & ChrW(7889) & "n d" & ChrW(242) & " t" & ChrW(236) & "m c" & ChrW(226) & "u h" & ChrW(7887) & "i theo c" & ChrW(7909) & "m t" & ChrW(7915) & ", b" & ChrW(7841) & "n ph" & ChrW(7843) & "i " & ChrW(273) & "i" & ChrW(7873) & "n c" & ChrW(7909) & "m t" & ChrW(7915) & " c" & ChrW(7847) & "n t" & ChrW(236) & "m." & vbCrLf & "N" & ChrW(7871) & "u kh" & ChrW(244) & "ng mu" & ChrW(7889) & "n d" & ChrW(242) & " t" & ChrW(236) & "m n" & ChrW(7919) & "a, b" & ChrW(7841) & "n h" & ChrW(227) & "y nh" & ChrW(7845) & "p ch" & ChrW(7885) & "n l" & ChrW(7879) & "nh " & ChrW(8220) & "Hu" & ChrW(7927) & ChrW(8221)
        Application.Assistant.DoAlert "Hý" & ChrW(7899) & "ng d" & ChrW(7851) & "n", msg, 0, 4, 0, 0, 0
    Else
        keysearch.Hide
        If CheckBox1 = False Then
            Call End_cau_DA
        Else
            Call End_cau_De
        End If
        Call Danh_dau_cau_hoi_chua_tu_khoa(TextBox1)
        TextBox1 = ""
        CheckBox1 = False
    End If
End Sub
Private Sub Danh_dau_cau_hoi_chua_tu_khoa(ByRef TextBox1 As String)
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    Tmp = TextBox1
' Doan code nay dung de xoa bo khoang trang du thua o 2 bia cua tu khoa
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
' Doan code dung de do tim cau hoi chua tu khoa va to mau highlight
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = Tmp
        .Replacement.text = " " & Tmp & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .text = "(Câu [0-9]{1,4})(*)" & Tmp
        .Replacement.text = "~.|.~" & "\1\2" & Tmp
        .Forward = False
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = " " & Tmp & " "
        .Replacement.text = Tmp
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
End Sub
Private Sub End_cau_DA()
' Convert_McMix Macro
Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    With Selection.Find
        .text = "z.zz^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "^p "
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = " Câu "
        .Replacement.text = " Câu$"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWholeWord = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "Câu "
        .Replacement.text = "z.zz^pCâu "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWholeWord = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = " Câu$"
        .Replacement.text = " Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText text:="z.zz"
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.zz^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceOne
    End With
    For i = 1 To ActiveDocument.Tables.Count
    ActiveDocument.Tables(i).Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "z.zz^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
        Selection.TypeText "z.zz"
    Next i
    Selection.HomeKey Unit:=wdStory
End Sub
Private Sub End_cau_De()
' Convert_McMix Macro
Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    With Selection.Find
        .text = "z.zz^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "^p "
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:]*)(A.*)(B.*)(C.*)(D.*)(^13)"
        .Replacement.text = "\1\2\3\4\5\6" & "z.zz^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
