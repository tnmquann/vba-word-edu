Attribute VB_Name = "S_Bank"
Option Explicit
Dim i As Integer
Dim ab() As Integer
Dim dapanmoi() As Integer
Dim Title, msg As String
Dim InAns() As String

Sub Xoadongtrang(ByVal control As Office.IRibbonControl)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^p "
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
   
    With Selection.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
End Sub
Sub Xoakhoangtrang(ByVal control As Office.IRibbonControl)
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "  "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
End Sub
Sub ChuanhoaCSDL(ByVal control As Office.IRibbonControl)
 Call ChuanDATA
End Sub
Sub ChuanDATA()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0)
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
            .text = "([AaBbCcDd])([.:\)])"
            .Replacement.text = "\1\2" & " "
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
            .text = "^11"
            .Replacement.text = "^13"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
            .text = "(\[\<)([Bb])([Rr])(\>\])"
            .Replacement.text = "\1\2\3\4" & "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "( )([.:,\)])"
        .Replacement.text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "( "
        .Replacement.text = "("
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
            .text = "([^32^9])([BCD])(.)"
            .Replacement.text = "^p" & "\2" & ". "
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "C©u"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
        
    With Selection.Find
        .text = "Caâu"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    With Selection.Find
        .ClearFormatting
        .text = "^p "
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    
    With Selection.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "  "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
   With Selection.Find
        .text = "^p "
        .Replacement.text = "^p"
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
        .text = " ^p"
        .Replacement.text = "^p"
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
        .text = "^p^p"
        .Replacement.text = "^p"
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
        .text = " ."
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = "([.:\)])( )"
        .Replacement.text = "\1\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = " "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    '''''
    Selection.WholeStory
    Selection.ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(1.5)
        .LeftMargin = CentimetersToPoints(1.5)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.6)
        .FooterDistance = CentimetersToPoints(0.6)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With
        'Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 12
        Selection.HomeKey Unit:=wdStory
    Call S_SerialHDD
    If ktBanQuyen = False Then S_NoteRig.Show
End Sub

Sub DelBr(ByVal control As Office.IRibbonControl)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[<Br>]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub InsBr(ByVal control As Office.IRibbonControl)
    Dim myRange As Range
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
        Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.HomeKey Unit:=wdLine
        Selection.TypeParagraph
    Next i
    With Selection.Find
        .text = "  "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "^p "
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "( )([.:\),])"
        .Replacement.text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find.Replacement.Font
       .Bold = True
       .Color = wdColorBlue
    End With
    With Selection.Find
        .text = "(^13)([CcBb])([âaà©])([ui])( [0-9]{1,4})"
        .Replacement.text = "\1" & "$#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "(Caâu [0-9]{1,4})([.:\)])"
        .Replacement.text = "$#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
For i = 1 To ActiveDocument.Tables.Count
    ActiveDocument.Tables(i).Cell(1, 1).Select
    Selection.Find.Execute FindText:="(Câu [0-9]{1,4})([.:\)])", MatchWildcards:=True
    If Selection.Find.Found = True Then
        'Selection.Find.Execute Replace:=wdReplaceOne
        Selection.TypeBackspace
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=2
        Selection.TypeParagraph
        Selection.TypeText text:="$#"
    End If
Next i
    Selection.HomeKey Unit:=wdStory
'Exit Sub
    With Selection.Find
        .text = "$#"
    End With
    i = 1
    Do While Selection.Find.Execute = True
        Selection.TypeText text:="[<Br>]" & Chr(13) & "Câu " & i & "."
        Selection.EndKey Unit:=wdLine, Extend:=wdMove
        i = i + 1
    Loop
    Selection.EndKey Unit:=wdStory, Extend:=wdMove
    Call ClearBlankBf
    Selection.TypeParagraph
    Selection.TypeText text:="[<Br>]"
    With Selection.Find
        .text = "\[\<Br\>\]" & Chr(13) & "(Câu 1.)"
        .Replacement.text = "Câu 1."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".."
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
         .text = ".:"
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
         .text = ".)"
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4})(.)"
        .Replacement.text = "\1\2" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
    Call S_SerialHDD
    If ktBanQuyen = False Then S_NoteRig.Show
    
End Sub
Sub Dem_cau()
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove 'Dua con tro ve dau van ban
    Selection.Find.ClearFormatting
    i = 0
    Selection.Find.ClearFormatting
    With Selection.Find
            .text = "[<br>]"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = False
            .MatchWildcards = False
    End With
    Do While Selection.Find.Execute = True
       i = i + 1
    Loop
    MsgBox "Tông sô câu: " & i
End Sub
Sub Index(ByVal control As Office.IRibbonControl)
    Dim title2, msg As String
    Dim ktMsg As Byte
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
            .text = "[<br>]"
            '.Forward = True
            '.Wrap = wdFindContinue
            .MatchWildcards = False
    End With
    If Selection.Find.Execute = False Then
            title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "D" & ChrW(7919) & " li" & ChrW(7879) & "u c" & _
            ChrW(7911) & "a b" & ChrW(7841) & "n không có ký hi" & ChrW(7879) & _
            "u ""[<Br>]"". N" & ChrW(7871) & "u b" & ChrW(7841) & "n mu" & ChrW(7889) _
             & "n " & ChrW(273) & "ánh th" & ChrW(7913) & " t" & ChrW(7921) & " cho d" _
             & ChrW(7919) & " li" & ChrW(7879) & "u có chèn mã [DS10.C1.1...a] thì ch" & ChrW(7885) & _
             "n ""Yes"". B" & ChrW(7841) & "n có ti" & ChrW(7871) & "p t" & ChrW(7909) & "c ?"
            ktMsg = Application.Assistant.DoAlert(title2, msg, 4, 2, 0, 0, 1)
            If ktMsg = 6 Then
                Call Index_IDQ
                Exit Sub
            Else
                Exit Sub
            End If
    End If
        
    i = 1
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "[<br>] "
            .Replacement.text = "[<br>]"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "[<br>]" & "^p"
            .Replacement.text = "[<br>]"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        'Do While .Execute
            .Execute Replace:=wdReplaceAll
        'Loop
        End With
        
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
        With Selection.Find
            .text = "[<br>]"
            .Replacement.text = "#$"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "#$"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        End With
    Do While Selection.Find.Execute = True
        i = i + 1
        Selection.TypeText text:="[<Br>]" & Chr(13) & "Câu " & i & ". "
        Selection.EndKey Unit:=wdLine
    Loop
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    Selection.TypeText text:="Câu 1. "
End Sub
Private Sub Index_IDQ()
    i = 0
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "( )(\[)(??1?.)(*)([abcd]\])"
            .Replacement.text = "\2\3\4\5"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
   
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    Selection.Find.ClearFormatting
   
        With Selection.Find
            .text = "(\[)(??1?.)(*)([abcd])(\])"
            .MatchWildcards = True
        End With
    Do While Selection.Find.Execute = True
        i = i + 1
        Selection.HomeKey Unit:=wdLine
        Selection.TypeText text:="Câu " & i & ". "
        Selection.EndKey Unit:=wdLine
    Loop
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .text = "(Câu )([0-9]{1,3})(.)"
        .Replacement.text = "\1\2\3"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
End Sub
Sub Index2(ByVal control As Office.IRibbonControl)
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find.Replacement.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = True
        .Italic = False
        .Color = 13382400
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:\)])"
        .Replacement.text = "$000#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "$000#^t"
        .Replacement.text = "$000#"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "$000# "
        .Replacement.text = "$000#"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    With Selection.Find
            .text = "$000#"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    i = 0
    Do While Selection.Find.Execute = True
        i = i + 1
        Selection.TypeText text:="Câu " & i & ". "
        Selection.EndKey Unit:=wdLine, Extend:=wdMove
    Loop
End Sub

Sub TestGroup(ByVal control As Office.IRibbonControl)
S_QG.Show
End Sub
Sub S_BankNew()
On Error GoTo S_Quit
Dim MADEAUTO As String
Dim QBank() As String
Dim L1, L2, L3, L4, lmax, socot, sodong As Integer
Dim lv1, lv2, lv3, lv4, i1, i2, i3, i4, j, tam2 As Integer
Dim tonglv1, tonglv2, tonglv3, tonglv4 As Integer
Dim thutucau As Integer
Dim Bai, f_dich As String
Dim Dapan() As Integer
Dim InAns() As String
Dim www As New Word.Application
Dim bank As New Word.Document
Dim S_sode As Byte
ReDim InAns(Val(S_matran.ComboSode.Value))
ReDim Dapan(Val(S_matran.Ltong))
ReDim QBank(Val(S_matran.Ltong))
ReDim dapanmoi(Val(S_matran.Ltong))
S_matran.Hide
MADEAUTO = S_matran.ComboMADE.text
'tao thu muc luu de
Select Case ktlop
        Case 12
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_matran.TextMon.text
        Case 11
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_matran.TextMon.text
        Case 10
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_matran.TextMon.text
        Case Else
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Other\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Other\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Other\" & S_matran.TextMon.text
End Select
''''''''''''''''''''''
'IN HEADER
''''''''''''''''''''''
'Kiem tra ton tai Header
        Dim docOpener As Document
        If FExists(S_Drive & "S_Bank&Test\S_Templates\default_Header_" & Right(S_matran.ComboHead, 1) & ".docx") = False Then
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "Header " & ChrW(273) & "ã ch" & ChrW(7885) & "n không có trong th" & _
            ChrW(432) & " m" & ChrW(7909) & "c S_Templates. B" & ChrW(7841) & "n ki" _
            & ChrW(7875) & "m tra và th" & ChrW(7921) & "c hi" & ChrW(7879) & "n l" & _
            ChrW(7841) & "i."
            Application.Assistant.DoAlert Title, msg, 0, 3, 0, 0, 0
            www.Quit
            Set www = Nothing
            Exit Sub
        End If
        'Kiem tre Header dang mo thi dong lai
        If docIsOpen("default_Header_" & Right(S_matran.ComboHead, 1) & ".docx") Then
            Set docOpener = Application.Documents("default_Header_" & Right(S_matran.ComboHead, 1) & ".docx")
            docOpener.Close
            Set docOpener = Nothing
        End If
Dim S_Header As Document
Dim ktAns As Boolean
Dim myRange As Range
Dim MD_in As String
'Phan noi dung
For S_sode = 1 To Val(S_matran.ComboSode.Value)
        thutucau = 0
        Documents.add
        Call S_PageSetup

        Select Case S_matran.ComboHead
            Case "Header 1"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_1.docx")
            Case "Header 2"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_2.docx")
            Case "Header 3"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_3.docx")
            Case "Header 4"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_4.docx")
            Case "Header 5"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_5.docx")
        End Select
        www.Selection.Tables(1).Select
        www.Selection.Copy
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        S_Header.Close
        Set S_Header = Nothing
        
        ktAns = False
        If S_matran.ComboAns = "Before" And Val(S_matran.Ltong) <= 50 Then
                ktAns = True
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx")
                Set myRange = www.ActiveDocument.Tables(Int(((Val(S_matran.Ltong) - 1) / 5)) + 1).Range
                myRange.Copy
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
                Selection.TypeParagraph
                Set myRange = Nothing
        End If
        Dim MadeTmp As String
        MadeTmp = ""
        If S_matran.ComboMADE = "Auto" Then
            MadeTmp = S_sode Mod 10 & Int(89 * Rnd() + 10)
        Else
            MadeTmp = MADEAUTO - 1 + S_sode
        End If
        ActiveDocument.Variables("MADE") = MadeTmp
        ActiveDocument.Variables("<lop>") = ktlop
        ActiveDocument.Fields.Update
        If ktAns = True Then
        S_Header.Close
        ktAns = False
        Set S_Header = Nothing
        End If
''''''''''''''''''''''
Dim S_Khode As String
If S_matran.Kho1 Then S_Khode = "S_Bank"
If S_matran.Kho2 Then S_Khode = "S_Bank 2"
If S_matran.Kho3 Then S_Khode = "S_Bank 3"
Dim BT_LT As Byte
For i = 1 To S_matran.ListBox1.ListCount
    lv1 = Val(S_matran.ListBox1.list(i - 1, 3))
    'MsgBox lv1
    If lv1 > 0 Then
        Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 10, 1)
        If Right(S_matran.ListBox1.list(i - 1, 0), 3) = "LT]" Then
            BT_LT = 0
        Else
            BT_LT = 4
        End If
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 3)))
        For j = 1 To lv1
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".1." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 8) & "].dat"
        Next j
    End If
Next i
tonglv1 = thutucau
For i = 1 To S_matran.ListBox1.ListCount
    lv2 = Val(S_matran.ListBox1.list(i - 1, 5))
    If lv2 > 0 Then
    Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 10, 1)
    If Right(S_matran.ListBox1.list(i - 1, 0), 3) = "LT]" Then
        BT_LT = 0
    Else
        BT_LT = 4
    End If
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 5)))
        For j = 1 To lv2
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".2." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 8) & "].dat"
        Next j
    End If
Next i
tonglv2 = thutucau - tonglv1
For i = 1 To S_matran.ListBox1.ListCount
    lv3 = Val(S_matran.ListBox1.list(i - 1, 7))
    If lv3 > 0 Then
    Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 10, 1)
    If Right(S_matran.ListBox1.list(i - 1, 0), 3) = "LT]" Then
        BT_LT = 0
    Else
        BT_LT = 4
    End If
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 7)))
        For j = 1 To lv3
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".3." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 8) & "].dat"
        Next j
    End If
Next i
tonglv3 = thutucau - tonglv2 - tonglv1
For i = 1 To S_matran.ListBox1.ListCount
    lv4 = Val(S_matran.ListBox1.list(i - 1, 9))
    If lv4 > 0 Then
    Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 10, 1)
    If Right(S_matran.ListBox1.list(i - 1, 0), 3) = "LT]" Then
        BT_LT = 0
    Else
        BT_LT = 4
    End If
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 9)))
        For j = 1 To lv4
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".4." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 8) & "].dat"
        Next j
    End If
Next i
tonglv4 = thutucau - tonglv3 - tonglv2 - tonglv1
Select Case S_matran.ComboLevel
Case "(1,2)(3,4)"
    For i = 1 To tonglv1 + tonglv2
        Call XaoMang(QBank(i), QBank(Int(Rnd * (tonglv1 + tonglv2)) + 1))
    Next i
    For i = 1 To tonglv3 + tonglv4
        Call XaoMang(QBank(i + tonglv1 + tonglv2), QBank(Int(Rnd * (tonglv3 + tonglv4) + 1) + tonglv1 + tonglv2))
    Next i
Case "(1,2)(3)(4)"
    For i = 1 To tonglv1 + tonglv2
        Call XaoMang(QBank(i), QBank(Int(Rnd * (tonglv1 + tonglv2)) + 1))
    Next i
    For i = 1 To tonglv3
        Call XaoMang(QBank(i + tonglv1 + tonglv2), QBank(Int(Rnd * (tonglv3) + 1) + tonglv1 + tonglv2))
    Next i
    For i = 1 To tonglv4
        Call XaoMang(QBank(i + tonglv1 + tonglv2 + tonglv3), QBank(Int(Rnd * (tonglv4) + 1) + tonglv1 + tonglv2 + tonglv3))
    Next i
Case "(1)(2)(3)(4)"
    For i = 1 To tonglv1
        Call XaoMang(QBank(i), QBank(Int(Rnd * (tonglv1)) + 1))
    Next i
    For i = 1 To tonglv2
        Call XaoMang(QBank(i + tonglv1), QBank(Int(Rnd * (tonglv2) + 1) + tonglv1))
    Next i
    For i = 1 To tonglv3
        Call XaoMang(QBank(i + tonglv1 + tonglv2), QBank(Int(Rnd * (tonglv3) + 1) + tonglv1 + tonglv2))
    Next i
    For i = 1 To tonglv4
        Call XaoMang(QBank(i + tonglv1 + tonglv2 + tonglv3), QBank(Int(Rnd * (tonglv4) + 1) + tonglv1 + tonglv2 + tonglv3))
    Next i
Case Else
    For i = 1 To thutucau
        Call XaoMang(QBank(i), QBank(Int(Rnd * thutucau) + 1))
    Next i
End Select
'''''
    Dim addBank, addBank_Old As String
    Dim tam() As String
    addBank_Old = ""
    Call RandNum(thutucau)
For i = 1 To thutucau
    tam = Split(QBank(i), ".")
    addBank = tam(4) & "." & tam(5) & "." & tam(6)
    ab(i) = ab(i) - 4 * Int((ab(i) + 3) / 4) + 4
    If addBank <> addBank_Old Then
        addBank_Old = addBank
        Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & Mid(tam(4), 4, 2) & "\" & addBank, PasswordDocument:="159")
    End If
    bank.Tables(8 * (tam(0) - 1) + tam(1) + tam(3)).Cell(6 * (tam(2) - 1) + 7, 2).Select
    If S_matran.Theo_Bai Then
        www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Else
        www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
        www.Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If
    Select Case Trim(www.Selection)
    Case "A"
    Dapan(i) = 1
    Case "B"
    Dapan(i) = 2
    Case "C"
    Dapan(i) = 3
    Case "D"
    Dapan(i) = 4
    End Select
    i1 = 1
    i2 = 2
    i3 = 3
    i4 = 4
    Select Case Dapan(i) - ab(i)
            Case 1
                tam2 = i1
                i1 = i2
                i2 = i3
                i3 = i4
                i4 = tam2
                dapanmoi(i) = ab(i)
                
            Case 2
                tam2 = i1
                i1 = i3
                i3 = tam2
                tam2 = i2
                i2 = i4
                i4 = tam2
                dapanmoi(i) = ab(i)
                
            Case 3
                tam2 = i4
                i4 = i1
                i1 = tam2
                dapanmoi(i) = ab(i)
                
            Case -1
                tam2 = i4
                i4 = i3
                i3 = i2
                i2 = i1
                i1 = tam2
                dapanmoi(i) = ab(i)
                
            Case -2
                tam2 = i1
                i1 = i3
                i3 = tam2
                tam2 = i2
                i2 = i4
                i4 = tam2
                dapanmoi(i) = ab(i)
                
            Case -3
                tam2 = i4
                i4 = i1
                i1 = tam2
                dapanmoi(i) = ab(i)
            Case 0
                dapanmoi(i) = ab(i)
            Case Else
                dapanmoi(i) = "9"
            End Select
    'Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & Mid(tam(4), 4, 2) & "\" & addBank, passworddocument:="159")
    'Do chieu dai cac phuong an
    bank.Tables(8 * (tam(0) - 1) + tam(1) + tam(3)).Cell(6 * (tam(2) - 1) + 3, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank1"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L1 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L1 = L1 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L1 = L1 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
    bank.Tables(8 * (tam(0) - 1) + tam(1) + tam(3)).Cell(6 * (tam(2) - 1) + 4, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank2"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L2 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L2 = L2 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L2 = L2 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
    bank.Tables(8 * (tam(0) - 1) + tam(1) + tam(3)).Cell(6 * (tam(2) - 1) + 5, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank3"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L3 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L3 = L3 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L3 = L3 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
    bank.Tables(8 * (tam(0) - 1) + tam(1) + tam(3)).Cell(6 * (tam(2) - 1) + 6, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank4"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L4 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L4 = L4 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L4 = L4 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
        
        lmax = L1
            If Val(lmax) < L2 Then lmax = L2
            If Val(lmax) < L3 Then lmax = L3
            If Val(lmax) < L4 Then lmax = L4
            If Val(lmax) < 10 Then lmax = "0" & lmax
            If Val(lmax) > 60 Then lmax = 60
            If Val(lmax) = 0 Then GoTo S_Quit
        If lmax < 18 Then
                sodong = 1
                socot = 4
        ElseIf lmax >= 18 And lmax < 45 Then
                sodong = 2
                socot = 2
        Else
                sodong = 4
                socot = 1
        End If
        'Bat dau in de
        'Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & Mid(tam(4), 4, 2) & "\" & addBank, passworddocument:="159")
        Call FontFormat
        Selection.TypeText text:="Câu " & i & ". "
        If S_matran.InmaCH.Value Then
            Select Case tam(1)
                Case "1"
                MD_in = "a"
                Case "2"
                MD_in = "b"
                Case "3"
                MD_in = "c"
                Case "4"
                MD_in = "d"
            End Select
            
            If tam(3) = "0" Then
                Selection.TypeText text:=Left(addBank, 8) & "." & tam(0) & ".LT." & MD_in & "] "
            Else
                Selection.TypeText text:=Left(addBank, 8) & "." & tam(0) & ".BT." & MD_in & "] "
            End If
        End If
        bank.Tables(8 * (tam(0) - 1) + tam(1) + tam(3)).Cell(6 * (tam(2) - 1) + 2, 2).Select
        www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        If Selection.Tables.Count > 0 Then
                Selection.Tables(1).Select
                Selection.MoveUp Unit:=wdLine, Count:=1
                Selection.HomeKey Unit:=wdLine
                Selection.EndKey Unit:=wdLine, Extend:=wdExtend
                If Left(Selection, 4) = "Câu " And Len(Selection) <= 11 Then
                    Selection.TypeBackspace
                    Selection.TypeBackspace
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.Tables(1).Cell(1, 1).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                    End With
                    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(-0.2)
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1 ', Extend:=wdExtend
                    With Selection.Font
                            .Name = "Times New Roman"
                            .Size = 12
                            .Bold = True
                            .Color = 13382400
                    End With
                    Selection.TypeText text:="Câu " & i & ". "
                End If
                    Selection.EndKey Unit:=wdStory
                    Selection.TypeParagraph
        Else
                Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
                Selection.TypeParagraph
        End If
 
        Dim myTable As Table
        Set myTable = ActiveDocument.Tables.add(Range:=Selection.Range, NumRows:=sodong, NumColumns:=socot)
        Select Case socot
            Case 4
            myTable.Columns(1).SetWidth ColumnWidth:=131, RulerStyle:=wdAdjustNone
            myTable.Columns(2).SetWidth ColumnWidth:=120, RulerStyle:=wdAdjustNone
            myTable.Columns(3).SetWidth ColumnWidth:=120, RulerStyle:=wdAdjustNone
            Case 2
            myTable.Columns(1).SetWidth ColumnWidth:=251, RulerStyle:=wdAdjustNone
            End Select
        Application.Keyboard (1033)
        Selection.TypeText text:="A."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i1
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
       
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:="B."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i2
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:="C."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i3
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:="D."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i4
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=False
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        Call S_ParagaphFormat
        Selection.MoveDown Unit:=wdLine, Count:=1
        Set myTable = Nothing
Next i
    
''''''''''''''''''
     'IN FOOTER
''''''''''''''''''
        Selection.TypeParagraph
        Dim ktFooter As Boolean
        Dim S_Footer As Document
        ktFooter = False
        Select Case S_matran.ComboFooter
        Case "Footer 1"
            ktFooter = True
            Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Footer_1.docx")
            www.Selection.WholeStory
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Case "Footer 2"
            ktFooter = True
            Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Footer_2.docx")
            www.Selection.WholeStory
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Case "Default"
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.TypeText text:="---------- " & "H" & ChrW(7870) & "T" & " ----------"
        End Select
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Call FontFormat2
        Selection.TypeText text:="Trang "
        Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            "PAGE  ", PreserveFormatting:=True
        Selection.TypeText text:="/"
        Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            "NUMPAGES  ", PreserveFormatting:=True
        Selection.TypeText text:=" - M" & ChrW(227) & " " & ChrW(273) & ChrW(7873) & " thi "
        Selection.TypeText text:=MadeTmp
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        If ktFooter = True Then S_Footer.Close
'''''''''''''''''''
        If S_matran.ComboAns = "After" And Val(S_matran.Ltong) <= 50 Then
                ktAns = True
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx")
                Set myRange = www.ActiveDocument.Tables(Int(((S_matran.Ltong - 1) / 5)) + 1).Range
                myRange.Copy
                Selection.TypeParagraph
                Selection.TypeParagraph
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End If
        If ktAns = True Then
        S_Header.Close
        Set myRange = Nothing
        End If
        Set S_Header = Nothing
        Set S_Footer = Nothing
        If S_matran.InmaCH.Value Then
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdPink
            With Selection.Find
                .text = "(\[)([DGH])([STH])(?????????)(T)(.)([abcd])(\])(^32)"
                .Replacement.text = "\1\2\3\4\5\6\7\8\9"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        End If
'''''''''''''''''''
    Selection.EndKey Unit:=wdStory, Extend:=wdMove
    ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_matran.TextMon & "] Made " & MadeTmp & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    If S_sode > 1 Then ActiveDocument.Close
    InAns(S_sode) = MadeTmp
    For i = 1 To Val(S_matran.Ltong)
    InAns(S_sode) = InAns(S_sode) & dapanmoi(i)
    Next i
Next S_sode
    bank.Close (False)
    Documents.add
    Call S_PageSetup
    Selection.TypeText text:=ChrW(272) & "ÁP ÁN [" & S_matran.TextMon & "]:"
    ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_matran.TextMon & "] Dapan" & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    For i = 1 To Val(S_matran.ComboSode.Value)
        For j = 1 To Val(S_matran.Ltong)
            dapanmoi(j) = Mid(InAns(i), j + 3, 1)
        Next j
        Call in_dapanMT(Left(InAns(i), 3), S_matran.Ltong)
        Selection.MoveDown Unit:=wdLine, Count:=1
    Next i
    ActiveDocument.Save
    
''''''''''''''''''
    www.Quit (False)
    Set bank = Nothing
    Set www = Nothing
    Unload S_matran
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Các " & ChrW(273) & ChrW(7873) & " " & ChrW(273) & "ã l" & ChrW(432) & "u trong:" & Chr(13) & _
    f_dich & "\Made_xxx.docx"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Call S_SerialHDD
    If ktBanQuyen = False Then S_NoteRig.Show
Exit Sub
S_Quit:
    If www.Documents.Count > 0 Then
    For i = www.Documents.Count To 1 Step -1
        www.Documents(i).Close (False)
    Next i
    End If
    www.Quit (False)
    Set bank = Nothing
    Set www = Nothing
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Quá trình in " & ChrW(273) & ChrW(7873) & " x" & _
         ChrW(7843) & "y ra l" & ChrW(7895) & "i. Vui lòng th" & ChrW(7921) & _
        "c hi" & ChrW(7879) & "n l" & ChrW(7841) & "i!"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub
Sub S_BankNew_CD()
On Error GoTo S_Quit
Dim MADEAUTO As String
Dim QBank() As String
Dim lv1, lv2, lv3, lv4, j As Integer
Dim tonglv1, tonglv2, tonglv3, tonglv4 As Integer
Dim thutucau As Integer
Dim Bai, f_dich As String

Dim S_sode As Byte
ReDim InAns(Val(S_matran.ComboSode.Value))

ReDim QBank(Val(S_matran.Ltong))
ReDim dapanmoi(Val(S_matran.Ltong))
S_matran.Hide
MADEAUTO = S_matran.ComboMADE.text

'tao thu muc luu de
Select Case ktlop
        Case 12
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 12\" & S_matran.TextMon.text
        Case 11
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 11\" & S_matran.TextMon.text
        Case 10
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 10\" & S_matran.TextMon.text
        Case Else
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Other\" & S_matran.TextMon.text & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Other\" & S_matran.TextMon.text)
            End If
            f_dich = S_Drive & "S_Bank&Test\S_Test\Other\" & S_matran.TextMon.text
End Select
''''''''''''''''''''''
'IN HEADER
''''''''''''''''''''''
'Kiem tra ton tai Header
        Dim docOpener As Document
        If FExists(S_Drive & "S_Bank&Test\S_Templates\default_Header_" & Right(S_matran.ComboHead, 1) & ".docx") = False Then
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "Header " & ChrW(273) & "ã ch" & ChrW(7885) & "n không có trong th" & _
            ChrW(432) & " m" & ChrW(7909) & "c S_Templates. B" & ChrW(7841) & "n ki" _
            & ChrW(7875) & "m tra và th" & ChrW(7921) & "c hi" & ChrW(7879) & "n l" & _
            ChrW(7841) & "i."
            Application.Assistant.DoAlert Title, msg, 0, 3, 0, 0, 0
           
            Exit Sub
        End If
        'Kiem tra Header dang mo thi dong lai
        If docIsOpen("default_Header_" & Right(S_matran.ComboHead, 1) & ".docx") Then
            Set docOpener = Application.Documents("default_Header_" & Right(S_matran.ComboHead, 1) & ".docx")
            docOpener.Close
            Set docOpener = Nothing
        End If

Dim MD_in As String

'Phan noi dung
For S_sode = 1 To Val(S_matran.ComboSode.Value)
        thutucau = 0
        Documents.add
        Call S_PageSetup
        '''''''''''''
        Call inHeader
'''''''''''
        Dim MadeTmp As String
        MadeTmp = ""
        If S_matran.ComboMADE = "Auto" Then
            MadeTmp = S_sode Mod 10 & Int(89 * Rnd() + 10)
        Else
            MadeTmp = MADEAUTO - 1 + S_sode
        End If
        ActiveDocument.Variables("MADE") = MadeTmp
        ActiveDocument.Variables("<lop>") = ktlop
        ActiveDocument.Fields.Update
        
''''''''''''''''''''''
If S_matran.Kho1 Then S_Khode = "S_Bank"
If S_matran.Kho2 Then S_Khode = "S_Bank 2"
If S_matran.Kho3 Then S_Khode = "S_Bank 3"
Dim BT_LT As Byte
BT_LT = 0
For i = 1 To S_matran.ListBox1.ListCount
    lv1 = Val(S_matran.ListBox1.list(i - 1, 3))
    'MsgBox lv1
    If lv1 > 0 Then
        Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 13, 2)
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 2)))
        For j = 1 To lv1
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".1." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 14) & "].dat"
        Next j
    End If
Next i
tonglv1 = thutucau
For i = 1 To S_matran.ListBox1.ListCount
    lv2 = Val(S_matran.ListBox1.list(i - 1, 5))
    If lv2 > 0 Then
        Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 13, 2)
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 4)))
        For j = 1 To lv2
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".2." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 14) & "].dat"
        Next j
    End If
Next i
tonglv2 = thutucau - tonglv1

For i = 1 To S_matran.ListBox1.ListCount
    lv3 = Val(S_matran.ListBox1.list(i - 1, 7))
    If lv3 > 0 Then
        Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 13, 2)
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 6)))
        For j = 1 To lv3
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".3." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 14) & "].dat"
        Next j
    End If
Next i
tonglv3 = thutucau - tonglv2 - tonglv1
For i = 1 To S_matran.ListBox1.ListCount
    lv4 = Val(S_matran.ListBox1.list(i - 1, 9))
    If lv4 > 0 Then
        Bai = Mid(S_matran.ListBox1.list(i - 1, 0), 13, 2)
        Call RandNum(Val(S_matran.ListBox1.list(i - 1, 8)))
        For j = 1 To lv4
        thutucau = thutucau + 1
        QBank(thutucau) = Bai & ".4." & ab(j) & "." & BT_LT & "." & Left(S_matran.ListBox1.list(i - 1, 0), 14) & "].dat"
        Next j
    End If
Next i
tonglv4 = thutucau - tonglv3 - tonglv2 - tonglv1

Select Case S_matran.ComboLevel
Case "(1,2)(3,4)"
    For i = 1 To tonglv1 + tonglv2
        Call XaoMang(QBank(i), QBank(Int(Rnd * (tonglv1 + tonglv2)) + 1))
    Next i
    For i = 1 To tonglv3 + tonglv4
        Call XaoMang(QBank(i + tonglv1 + tonglv2), QBank(Int(Rnd * (tonglv3 + tonglv4) + 1) + tonglv1 + tonglv2))
    Next i
Case "(1,2)(3)(4)"
    For i = 1 To tonglv1 + tonglv2
        Call XaoMang(QBank(i), QBank(Int(Rnd * (tonglv1 + tonglv2)) + 1))
    Next i
    For i = 1 To tonglv3
        Call XaoMang(QBank(i + tonglv1 + tonglv2), QBank(Int(Rnd * (tonglv3) + 1) + tonglv1 + tonglv2))
    Next i
    For i = 1 To tonglv4
        Call XaoMang(QBank(i + tonglv1 + tonglv2 + tonglv3), QBank(Int(Rnd * (tonglv4) + 1) + tonglv1 + tonglv2 + tonglv3))
    Next i
Case "(1)(2)(3)(4)"
    For i = 1 To tonglv1
        Call XaoMang(QBank(i), QBank(Int(Rnd * (tonglv1)) + 1))
    Next i
    For i = 1 To tonglv2
        Call XaoMang(QBank(i + tonglv1), QBank(Int(Rnd * (tonglv2) + 1) + tonglv1))
    Next i
    For i = 1 To tonglv3
        Call XaoMang(QBank(i + tonglv1 + tonglv2), QBank(Int(Rnd * (tonglv3) + 1) + tonglv1 + tonglv2))
    Next i
    For i = 1 To tonglv4
        Call XaoMang(QBank(i + tonglv1 + tonglv2 + tonglv3), QBank(Int(Rnd * (tonglv4) + 1) + tonglv1 + tonglv2 + tonglv3))
    Next i
Case Else
    For i = 1 To thutucau
        Call XaoMang(QBank(i), QBank(Int(Rnd * thutucau) + 1))
    Next i
End Select
'Selection.WholeStory
With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(0)
            .LeftIndent = CentimetersToPoints(1.75)
            .MirrorIndents = False
End With
Call RandNum(thutucau)
For i = 1 To thutucau
    Call InCau(QBank(i), i)
Next i
    
''''''''''''''''''
With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(0)
            .LeftIndent = CentimetersToPoints(0)
            .MirrorIndents = False
End With
Call inFooter(MadeTmp)
'''''''''''''''''''
      
        If S_matran.InmaCH.Value Then
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdPink
            With Selection.Find
                .text = "(\[)([DGH])([STH])(?????????)(T)(.)([abcd])(\])(^32)"
                .Replacement.text = "\1\2\3\4\5\6\7\8\9"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        End If
'''''''''''''''''''
    Selection.EndKey Unit:=wdStory, Extend:=wdMove
    ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_matran.TextMon & "] Made " & MadeTmp & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    If S_sode > 1 Then ActiveDocument.Close
    InAns(S_sode) = MadeTmp
    For i = 1 To Val(S_matran.Ltong)
    InAns(S_sode) = InAns(S_sode) & dapanmoi(i)
    Next i
Next S_sode

    Documents.add
    Call S_PageSetup
    Selection.TypeText text:=ChrW(272) & "ÁP ÁN [" & S_matran.TextMon & "]:"
    ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_matran.TextMon & "] Dapan" & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    For i = 1 To Val(S_matran.ComboSode.Value)
        For j = 1 To Val(S_matran.Ltong)
            dapanmoi(j) = Mid(InAns(i), j + 3, 1)
        Next j
        Call in_dapanMT(Left(InAns(i), 3), S_matran.Ltong)
        Selection.MoveDown Unit:=wdLine, Count:=1
    Next i
    ActiveDocument.Save
   
''''''''''''''''''
    Unload S_matran
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Các " & ChrW(273) & ChrW(7873) & " " & ChrW(273) & "ã l" & ChrW(432) & "u trong:" & Chr(13) & _
    f_dich & "\Made_xxx.docx"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Call S_SerialHDD
    If ktBanQuyen = False Then S_NoteRig.Show
Exit Sub
S_Quit:

    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Quá trình in " & ChrW(273) & ChrW(7873) & " x" & _
         ChrW(7843) & "y ra l" & ChrW(7895) & "i. Vui lòng th" & ChrW(7921) & _
        "c hi" & ChrW(7879) & "n l" & ChrW(7841) & "i!"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub

Private Sub XaoMang(ByRef so1 As String, ByRef so2 As String)
Dim tg As String
    tg = so1
    so1 = so2
    so2 = tg
End Sub
Private Sub in_dapanMT(ByRef md As String, ByRef socau As Integer)
    Dim T As String
    Dim ida As Integer
    Selection.TypeParagraph
    Selection.TypeText text:="M" & ChrW(227) & " " & ChrW(273) & ChrW(7873) & " [" & md & "]"
    Selection.TypeParagraph
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=15, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Application.Keyboard (1033)
    For ida = 1 To socau
        Select Case dapanmoi(ida)
            Case 1
                T = "A"
            Case 2
                T = "B"
            Case 3
                T = "C"
            Case 4
                T = "D"
            Case Else
                T = "_"
        End Select
        Selection.TypeText text:=ida & T
        If ida < socau Then Selection.MoveRight Unit:=wdCell
    Next ida
    Selection.WholeStory
    If socau > 105 Then
        Call FontFormat3
    Else
        Call FontFormat
    End If
    Selection.EndKey Unit:=wdLine
End Sub
Private Sub XaoSo(ByRef so1 As Integer, ByRef so2 As Integer)
Dim tg As Integer
    tg = so1
    so1 = so2
    so2 = tg
End Sub
Private Sub RandNum(ByRef n As Integer)
ReDim ab(n) As Integer
Dim iRannum, ii As Integer
Randomize
    For iRannum = 1 To n
        ab(iRannum) = iRannum
    Next
    For iRannum = 1 To n
        Call XaoSo(ab(iRannum), ab(Int(Rnd * n) + 1))
    Next
    'Dim tex As String
    'tex = ""
    'For ii = 1 To n
    'tex = tex & ab(ii) & "_"
    'Next ii
    'MsgBox tex
End Sub
Private Sub InCau(ByRef idtext As String, ByRef idex As Integer)
On Error GoTo S_Quit
Dim L1, L2, L3, L4, lmax, socot, sodong As Integer
Dim i1, i2, i3, i4, tam2 As Integer
Dim Dapan() As Integer
Dim www As New Word.Application
Dim bank As New Word.Document

ReDim Dapan(Val(S_matran.Ltong))

Dim addBank As String
Dim ktHDG_bm As Boolean
Dim tam() As String
Dim myRange As Range
tam = Split(idtext, ".")
    addBank = Mid(tam(4), 2, 4) & "." & tam(5) & "\" & tam(4) & "." & tam(5) & "." & tam(6) & "." & tam(7) & "." & tam(8)
    ab(idex) = ab(idex) - 4 * Int((ab(idex) + 3) / 4) + 4
        Set bank = www.Documents.Open(S_Drive & "S_Bank&Test\" & S_Khode & "\Lop " & Mid(tam(4), 4, 2) & "\Chuyen de\" & addBank, PasswordDocument:="159")
    If S_matran.inHDG.Value = True Then
        bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 7, 2).Select
        Set myRange = www.Selection.Range
        ktHDG_bm = True
        If Len(myRange) > 4 Then
            myRange.MoveEnd Unit:=wdCharacter, Count:=-1
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.Select
            With www.ActiveDocument.Bookmarks
                .add Range:=www.Selection.Range, Name:="bankHDG"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        Else
            ktHDG_bm = False
        End If
    End If
    bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 7, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
    www.Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    Select Case Trim(www.Selection)
    Case "A"
    Dapan(idex) = 1
    Case "B"
    Dapan(idex) = 2
    Case "C"
    Dapan(idex) = 3
    Case "D"
    Dapan(idex) = 4
    End Select
    i1 = 1
    i2 = 2
    i3 = 3
    i4 = 4
    Select Case Dapan(idex) - ab(idex)
            Case 1
                tam2 = i1
                i1 = i2
                i2 = i3
                i3 = i4
                i4 = tam2
                dapanmoi(idex) = ab(idex)
                
            Case 2
                tam2 = i1
                i1 = i3
                i3 = tam2
                tam2 = i2
                i2 = i4
                i4 = tam2
                dapanmoi(idex) = ab(idex)
                
            Case 3
                tam2 = i4
                i4 = i1
                i1 = tam2
                dapanmoi(idex) = ab(idex)
                
            Case -1
                tam2 = i4
                i4 = i3
                i3 = i2
                i2 = i1
                i1 = tam2
                dapanmoi(idex) = ab(idex)
                
            Case -2
                tam2 = i1
                i1 = i3
                i3 = tam2
                tam2 = i2
                i2 = i4
                i4 = tam2
                dapanmoi(idex) = ab(idex)
                
            Case -3
                tam2 = i4
                i4 = i1
                i1 = tam2
                dapanmoi(idex) = ab(idex)
            Case 0
                dapanmoi(idex) = ab(idex)
            Case Else
                dapanmoi(idex) = "9"
            End Select

    bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 3, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank1"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L1 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L1 = L1 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L1 = L1 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
    bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 4, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank2"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L2 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L2 = L2 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L2 = L2 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
    bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 5, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank3"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L3 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L3 = L3 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L3 = L3 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
    bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 6, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With www.ActiveDocument.Bookmarks
        .add Range:=www.Selection.Range, Name:="bank4"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    
    L4 = www.Selection.Characters.Count
        Select Case www.Selection.InlineShapes.Count
            Case 1
            L4 = L4 + Round(www.Selection.InlineShapes(1).Width / 5.8)
            Case 2
            L4 = L4 + Round((www.Selection.InlineShapes(1).Width + www.Selection.InlineShapes(2).Width) / 6.3)
        End Select
        
        lmax = L1
            If Val(lmax) < L2 Then lmax = L2
            If Val(lmax) < L3 Then lmax = L3
            If Val(lmax) < L4 Then lmax = L4
            If Val(lmax) < 10 Then lmax = "0" & lmax
            If Val(lmax) > 60 Then lmax = 60
            If Val(lmax) = 0 Then GoTo S_Quit
        If lmax < 16 Then
                sodong = 1
                socot = 4
        ElseIf lmax < 37 Then
                sodong = 2
                socot = 2
        Else
                sodong = 4
                socot = 1
        End If
        'Bat dau in de
        With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = True
        .Color = wdColorBlue
        End With
        'Call FontFormat
        Selection.TypeText text:="Câu " & idex & ". " & vbTab
        With ActiveDocument.Bookmarks
        .add Range:=Selection.Range, Name:="ttcau"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
        End With
        Dim MD_in As String
        If S_matran.InmaCH.Value Then
            Select Case tam(1)
                Case "1"
                MD_in = "a"
                Case "2"
                MD_in = "b"
                Case "3"
                MD_in = "c"
                Case "4"
                MD_in = "d"
            End Select
            
            If tam(3) = "0" Then
                Selection.TypeText text:=Left(addBank, 8) & "." & tam(0) & ".LT." & MD_in & "] "
            Else
                Selection.TypeText text:=Left(addBank, 8) & "." & tam(0) & ".BT." & MD_in & "] "
            End If
        End If
        bank.Tables(tam(1)).Cell(6 * (tam(2) - 1) + 2, 2).Select
        www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        www.Selection.Copy
                
        Selection.Paste 'AndFormat (wdFormatPlainText)
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        Selection.EndKey Unit:=wdStory
        Dim myTable As Table
        Set myTable = ActiveDocument.Tables.add(Range:=Selection.Range, NumRows:=sodong, NumColumns:=socot)
        Select Case socot
            Case 4
            myTable.Columns(1).SetWidth ColumnWidth:=155, RulerStyle:=wdAdjustNone
            myTable.Columns(2).SetWidth ColumnWidth:=110, RulerStyle:=wdAdjustNone
            myTable.Columns(3).SetWidth ColumnWidth:=110, RulerStyle:=wdAdjustNone
            Case 2
            myTable.Columns(1).SetWidth ColumnWidth:=265, RulerStyle:=wdAdjustNone
            End Select
        Application.Keyboard (1033)
        Selection.TypeText text:="A."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i1
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
       
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:="B."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i2
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:="C."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i3
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        
        Selection.MoveRight Unit:=wdCell
        Selection.TypeText text:="D."
        www.Selection.GoTo what:=wdGoToBookmark, Name:="bank" & i4
        www.Selection.Copy
        Selection.Paste 'AndFormat (wdFormatPlainText)
        Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=False
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(0)
            .LeftIndent = CentimetersToPoints(1.75)
        End With
        Selection.MoveDown Unit:=wdLine, Count:=1
        Set myTable = Nothing
        If ktHDG_bm And S_matran.inHDG Then
            
            www.Selection.GoTo what:=wdGoToBookmark, Name:="bankHDG"
            www.Selection.Copy
            Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
            Selection.GoTo what:=wdGoToBookmark, Name:="bankHDG"
            With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(0)
            .LeftIndent = CentimetersToPoints(1.75)
            End With
            ActiveDocument.Bookmarks("bankHDG").Delete
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.TypeParagraph
            
        End If
    Selection.GoTo what:=wdGoToBookmark, Name:="ttcau"
    With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(-1.75)
            .LeftIndent = CentimetersToPoints(1.75)
            .MirrorIndents = False
    End With
    Selection.EndKey Unit:=wdStory, Extend:=wdMove
        
bank.Close (False)
www.Quit (False)
Set bank = Nothing
Set www = Nothing
Set myRange = Nothing
Exit Sub
S_Quit:
www.Quit (False)
End Sub
Private Sub inHeader()
    Dim www As New Word.Application
    Dim S_Header As Document
    Dim ktAns As Boolean
    Dim myRange As Range
            Select Case S_matran.ComboHead
            Case "Header 1"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_1.docx")
            Case "Header 2"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_2.docx")
            Case "Header 3"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_3.docx")
            Case "Header 4"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_4.docx")
            Case "Header 5"
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_5.docx")
        End Select
        www.Selection.Tables(1).Select
        www.Selection.Copy
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        S_Header.Close
        Set S_Header = Nothing
        ktAns = False
        If S_matran.ComboAns = "Before" And Val(S_matran.Ltong) <= 50 Then
                ktAns = True
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx")
                Set myRange = www.ActiveDocument.Tables(Int(((Val(S_matran.Ltong) - 1) / 5)) + 1).Range
                myRange.Copy
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
                Selection.TypeParagraph
                Set myRange = Nothing
        End If
   www.Quit (False)
End Sub
Private Sub inFooter(ByRef Smade As String)
    Dim www As New Word.Application
    Dim ktAns As Boolean
    Dim myRange As Range
    Selection.TypeParagraph
    Dim ktFooter As Boolean
    Dim S_Footer As Document
        ktFooter = False
        Select Case S_matran.ComboFooter
        Case "Footer 1"
            ktFooter = True
            Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Footer_1.docx")
            www.Selection.WholeStory
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Case "Footer 2"
            ktFooter = True
            Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Footer_2.docx")
            www.Selection.WholeStory
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Case "Default"
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.TypeText text:="---------- " & "H" & ChrW(7870) & "T" & " ----------"
        End Select
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Call FontFormat2
        Selection.TypeText text:="Trang "
        Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            "PAGE  ", PreserveFormatting:=True
        Selection.TypeText text:="/"
        Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            "NUMPAGES  ", PreserveFormatting:=True
        Selection.TypeText text:=" - M" & ChrW(227) & " " & ChrW(273) & ChrW(7873) & " thi "
        Selection.TypeText text:=Smade
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        If ktFooter = True Then S_Footer.Close
      If S_matran.ComboAns = "After" And Val(S_matran.Ltong) <= 50 Then
                ktAns = True
                Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx")
                Set myRange = www.ActiveDocument.Tables(Int(((S_matran.Ltong - 1) / 5)) + 1).Range
                myRange.Copy
                Selection.TypeParagraph
                Selection.TypeParagraph
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End If
        If ktAns = True Then
        'S_Header.Close
        Set myRange = Nothing
        End If
        'Set S_Header = Nothing
        Set S_Footer = Nothing
    www.Quit (False)
End Sub
