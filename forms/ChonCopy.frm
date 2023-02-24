VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChonCopy 
   Caption         =   "  "
   ClientHeight    =   1995
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6540
   OleObjectBlob   =   "ChonCopy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChonCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
        ChonCopy.Hide
End Sub

Private Sub CommandButton2_Click()
        If OptionButton1 = False And OptionButton2 = False Then
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
            msg = "Vui l" & ChrW(242) & "ng ch" & ChrW(7885) & "n c" & ChrW(225) & "ch ch" & ChrW(233) & "p c" & ChrW(226) & "u h" & ChrW(7887) & "i."
            Application.Assistant.DoAlert Title, msg, 0, 8, 0, 0, 0
        End If
        If OptionButton1 = True Then
            ChonCopy.Hide
            Call End_cau_DA
        End If
        If OptionButton2 = True Then
            ChonCopy.Hide
            Call End_cau_De
        End If
End Sub

Private Sub Help_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label7_Click()
Dim ktW As Boolean
If ktW = True Then
ChonCopy.Height = 128
ktW = False
Else
ChonCopy.Height = 256
ktW = True
End If
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub UserForm_Click()

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
