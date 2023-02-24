Attribute VB_Name = "Module2"
Sub Tach_lay_cau_hoi(ByVal control As Office.IRibbonControl)
' Chep cau hoi trac nghiem ra mot file moi (khong chep HDG)
Application.ScreenUpdating = False
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)"
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)" & "  "
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Set ThisDoc = ActiveDocument
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "(Câu [0-9]{1,4}*)(A.*)(B.*)(C.*)(D.*)(^13)"
        .MatchWildcards = True
    If Selection.Find.Execute = True Then
    Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    Else
    Exit Sub
    End If
    ThisDoc.Activate
    Selection.Copy
    Do
    Selection.Copy
    ThatDoc.Activate
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    ThisDoc.Activate
    Selection.Copy
    Loop While Selection.Find.Execute(Forward:=True) = True
    End With
    ThatDoc.Activate
    Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Font.Name = "Times New Roman"
    Selection.TypeText text:="C" & ChrW(193) & "C C" & ChrW(194) & "U H" & ChrW(7886) & "I TR" & ChrW(7854) & "C NGHI" & ChrW(7878) & "M L" & ChrW(7844) & "Y RA T" & ChrW(7914) & " T" & ChrW(192) & "I LI" & ChrW(7878) & "U"
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Size = 16
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(225) & "c c" & ChrW(226) & "u h" & ChrW(7887) & "i tr" & ChrW(7855) & "c nghi" & ChrW(7879) & "m " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c ch" & ChrW(233) & "p ra m" & ChrW(7897) & "t file m" & ChrW(7899) & "i. B" & ChrW(7841) & "n nh" & ChrW(7899) & " save file l" & ChrW(7841) & "i nh" & ChrW(233)
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Tao_bang_dap_an(ByVal control As Office.IRibbonControl)
    GachDA.Show
End Sub
Sub Danh_dau_dap_an(ByVal control As Office.IRibbonControl)
If ActiveDocument.Tables.Count = 0 Then Exit Sub
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Application.ScreenUpdating = False
' Code giup han che nhan dang nham phuong an
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
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
    Selection.Find.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)"
        .Replacement.ClearFormatting
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)" & "  "
        .Replacement.ClearFormatting
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "(A.*)(B.*)(C.*)(D.*)"
        .Replacement.text = "#" & "\1" & "#" & "\2" & "#" & "\3" & "#" & "\4"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
' Chep bang dap an sang file tam và Text hoa no
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Set ThisDoc = ActiveDocument
        ActiveDocument.Tables(ActiveDocument.Tables.Count).Select
        Selection.Copy
    Set ThatDoc = Documents.add
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        socot = ActiveDocument.Tables(1).Columns.Count
        sohang = ActiveDocument.Tables(1).Rows.Count
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .text = "Câu"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .text = "([.,:_;^32-])"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        C = 0
        With Selection.Find
            .ClearFormatting
            .text = "([0-9]{1,4})"
            .Replacement.text = "#" & "\1" & "#"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        Do While Selection.Find.Execute = True
            C = C + 1
        Loop
        If socot > 1 Then
        ActiveDocument.Tables(1).Rows(1).Cells(2).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
        O2 = Selection.text
        Else
        O2 = ""
        End If
        If sohang > 1 And O2 = "#2#" Then
        ActiveDocument.Tables(1).Select
        For i = 1 To socot
            Selection.Tables(1).Columns(i).Select
            Selection.Cells.Merge
        Next i
        End If
        ActiveDocument.Tables(1).Select
            Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=True
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .text = "([^13^32^9])"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
'Ghi dap an vao mang
        Selection.HomeKey Unit:=wdStory
        Dim Arr(1 To 999)
        For j = 1 To C
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "#" & j & "#"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Arr(j) = Selection.text
        Next j
    ThatDoc.Close (No)
    ThisDoc.Activate
        Selection.HomeKey Unit:=wdStory
        For i = 1 To C
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "(Câu )" & i & "([.:])"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            If Selection.Find.Execute = True Then
                With Selection.Find
                    .text = "#" & Arr(i) & "."
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
            End If
            End With
            Selection.Find.Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Color = wdColorRed
            Selection.Font.Bold = True
            Selection.Font.Underline = wdUnderlineSingle
        Next i
        With Selection.Find
            .ClearFormatting
            .text = "#"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        Selection.Find.ClearFormatting
        Selection.Find.Font.Underline = wdUnderlineSingle
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.Underline = wdUnderlineNone
        With Selection.Find
            .text = "([^9^32])"
            .Replacement.text = "\1"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Them_br(ByVal control As Office.IRibbonControl)
' Convert_McMix Macro
Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    With Selection.Find
        .text = "[<br>]^p"
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
    With Selection.Find
        .text = "Câu "
        .Replacement.text = "[<br>]^pCâu "
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
    Selection.TypeText text:="[<br>]"
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "[<br>]^p"
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
        .text = "[<br>]^p"
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
        Selection.TypeText "[<br>]"
    Next i
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Size = 12
        .Color = wdColorGreen
    End With
    With Selection.Find
        .text = "[<br>]"
        .Replacement.text = "[<br>]"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Xoa_br(ByVal control As Office.IRibbonControl)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[\<" & "(br)" & "\>\]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
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
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Kiem_tra_loi(ByVal control As Office.IRibbonControl)
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Application.ScreenUpdating = False
Dim Cau, Traloi, PAnA, PAnB, PAnC, PAnD, msgA, msgB, msgC, msgD, TBLoiPA, msgTraloi
'Xoa Highilght
    Selection.WholeStory
    Options.DefaultHighlightColorIndex = wdNoHighlight
    Selection.Range.HighlightColorIndex = wdNoHighlight
' Chuan hoa cac tu khoa A.B.C.D.
    Selection.HomeKey Unit:=wdStory
        With Selection.Find
        .ClearFormatting
        .text = "([A-D].)"
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)" & "  "
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
' Dem so cau hoi co trong tai lieu
    Cau = 0
    With Selection.Find
        .ClearFormatting
        .text = "(Câu [0-9]{1,4}[.:])"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
    End With
    Do While Selection.Find.Execute = True
        Cau = Cau + 1
    Loop
    Selection.HomeKey Unit:=wdStory
' Dem so tu khoa A. co trong tai lieu
    PAnA = 0
    With Selection.Find
        .ClearFormatting
        .text = "A. "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
    End With
    Do While Selection.Find.Execute = True
        PAnA = PAnA + 1
    Loop
    Selection.HomeKey Unit:=wdStory
    If PAnA <> Cau Then              ' Thong bao loi khong khop nhau giua so A. va so cau
        msgA = "S" & ChrW(7889) & " t" & ChrW(7915) & " kho" & ChrW(225) & " ph" & ChrW(432) & "" & ChrW(417) & "ng " & ChrW(225) & "n A. kh" & ChrW(225) & "c v" & ChrW(7899) & "i s" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i (" & PAnA & " v" & ChrW(224) & " " & Cau & ")" & vbCrLf
        Options.DefaultHighlightColorIndex = wdYellow
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Highlight = True
        With Selection.Find
            .text = "A. "
            .Replacement.text = "A. "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
        msgA = ""
    End If
' Dem so tu khoa B. co trong tai lieu
    PAnB = 0
    With Selection.Find
        .ClearFormatting
        .text = "B. "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
    End With
    Do While Selection.Find.Execute = True
        PAnB = PAnB + 1
    Loop
    Selection.HomeKey Unit:=wdStory
    If PAnB <> Cau Then              ' Thong bao loi khong khop nhau giua so B. va so cau
        msgB = "S" & ChrW(7889) & " t" & ChrW(7915) & " kho" & ChrW(225) & " ph" & ChrW(432) & "" & ChrW(417) & "ng " & ChrW(225) & "n B. kh" & ChrW(225) & "c v" & ChrW(7899) & "i s" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i (" & PAnB & " v" & ChrW(224) & " " & Cau & ")" & vbCrLf
        Options.DefaultHighlightColorIndex = wdBrightGreen
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Highlight = True
        With Selection.Find
            .text = "B. "
            .Replacement.text = "B. "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
        msgB = ""
    End If
' Dem so tu khoa C. co trong tai lieu
    PAnC = 0
    With Selection.Find
        .ClearFormatting
        .text = "C. "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
    End With
    Do While Selection.Find.Execute = True
        PAnC = PAnC + 1
    Loop
    Selection.HomeKey Unit:=wdStory
    If PAnC <> Cau Then              ' Thong bao loi khong khop nhau giua so C. va so cau
        msgC = "S" & ChrW(7889) & " t" & ChrW(7915) & " kho" & ChrW(225) & " ph" & ChrW(432) & "" & ChrW(417) & "ng " & ChrW(225) & "n C. kh" & ChrW(225) & "c v" & ChrW(7899) & "i s" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i (" & PAnC & " v" & ChrW(224) & " " & Cau & ")" & vbCrLf
        Options.DefaultHighlightColorIndex = wdTurquoise
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Highlight = True
        With Selection.Find
            .text = "C. "
            .Replacement.text = "C. "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
        msgC = ""
    End If
' Dem so tu khoa D. co trong tai lieu
    PAnD = 0
    With Selection.Find
        .ClearFormatting
        .text = "D. "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
    End With
    Do While Selection.Find.Execute = True
        PAnD = PAnD + 1
    Loop
    Selection.HomeKey Unit:=wdStory
    If PAnD <> Cau Then              ' Thong bao loi khong khop nhau giua so C. va so cau
        msgD = "S" & ChrW(7889) & " t" & ChrW(7915) & " kho" & ChrW(225) & " ph" & ChrW(432) & "" & ChrW(417) & "ng " & ChrW(225) & "n D. kh" & ChrW(225) & "c v" & ChrW(7899) & "i s" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i (" & PAnD & " v" & ChrW(224) & " " & Cau & ")" & vbCrLf
        Options.DefaultHighlightColorIndex = wdGray25
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Highlight = True
        With Selection.Find
            .text = "D. "
            .Replacement.text = "D. "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
        msgD = ""
    End If
' Dem so bo phuong an A.B.C.D. duoc sap xep dung thu tu
    Traloi = 0
    With Selection.Find
        .ClearFormatting
        .text = "(A.)(*)(B.)(*)(C.)(*)(D.)(*)(^13)"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
    End With
    Do While Selection.Find.Execute = True
        Traloi = Traloi + 1
    Loop
    Selection.HomeKey Unit:=wdStory
    If Traloi <> Cau Then
        msgTraloi = "Th" & ChrW(7913) & " t" & ChrW(7921) & " A-B-C-D c" & ChrW(7911) & "a c" & ChrW(225) & "c ph" & ChrW(432) & "" & ChrW(417) & "ng " & ChrW(225) & "n ch" & ChrW(432) & "a " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c " & ChrW(273) & "" & ChrW(7843) & "m b" & ChrW(7843) & "o" & vbCrLf
    Else
        msgTraloi = ""
    End If
' Thong bao loi
    Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    If Traloi = Cau And PAnA = Cau And PAnB = Cau And PAnC = Cau And PAnD = Cau Then
        msg1 = "Ch" & ChrW(250) & "c m" & ChrW(7915) & "ng b" & ChrW(7841) & "n. C" & ChrW(225) & "c c" & ChrW(226) & "u h" & ChrW(7887) & "i c" & ChrW(7911) & "a b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(7843) & "m b" & ChrW(7843) & "o t" & ChrW(7889) & "t v" & ChrW(7873) & " m" & ChrW(7863) & "t k" & ChrW(7929) & " thu" & ChrW(7853) & "t."
        Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg1, 0, 4, 0, 0, 0
    Else
        TBLoiPA = "---------------------" & vbCrLf & "B" & ChrW(7841) & "n h" & ChrW(227) & "y ki" & ChrW(7875) & "m tra l" & ChrW(7841) & "i v" & ChrW(259) & "n b" & ChrW(7843) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c g" & ChrW(245) & " v" & ChrW(224) & " " & ChrW(273) & "" & ChrW(7843) & "m b" & ChrW(7843) & "o r" & ChrW(7857) & "ng c" & ChrW(225) & "c k" & ChrW(253) & " t" & ChrW(7921) & " A. B. C. D." & vbCrLf & "ch" & ChrW(7881) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c d" & ChrW(249) & "ng " & ChrW(273) & "" & ChrW(7875) & " l" & ChrW(224) & "m t" & ChrW(7915) & " kho" & ChrW(225) & " cho c" & ChrW(225) & "c ph" & ChrW(432) & "" & ChrW(417) & "ng " & ChrW(225) & "n tr" & ChrW(7843) & " l" & ChrW(7901) & "i m" & ChrW(224) & " th" & ChrW(244) & "i." & vbCrLf & "---------------------" & vbCrLf & _
            "Sau khi ch" & ChrW(7881) & "nh s" & ChrW(7917) & "a, b" & ChrW(7841) & "n nh" & ChrW(7899) & " ch" & ChrW(7841) & "y ch" & ChrW(7913) & "c n" & ChrW(259) & "ng ki" & ChrW(7875) & "m tra l" & ChrW(7895) & "i l" & ChrW(7841) & "i l" & ChrW(7847) & "n n" & ChrW(7919) & "a nh" & ChrW(233)
        msg2 = msgTraloi & msgA & msgB & msgC & msgD & TBLoiPA
        Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg2, 0, 4, 0, 0, 0
    End If
End Sub
Sub GhepHoi_Dap(ByVal control As Office.IRibbonControl)

' Macro ghep de bai va huong dan giai
' Luu lai de danh co dip su dung

Dim ThisDoc As Document
Set ThisDoc = ActiveDocument
Dim ThatDoc As Document
Dim TamDoc As Document
Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
Set TamDoc = Documents.add(DocumentType:=wdNewBlankDocument)
ThisDoc.Activate
Selection.WholeStory
Selection.Copy
TamDoc.Activate
Selection.PasteAndFormat (wdListCombineWithExistingList)
ThisDoc.ConvertNumbersToText
ThisDoc.Activate
Dim socau As Integer
socau = 0
With Selection.Find
.text = "Câu [0-9]{1,4}?"
.Replacement.text = "Câu%00"
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = True
Do While .Execute
.Execute Replace:=wdReplaceOne
socau = socau + 1
Loop
End With
If (socau Mod 2 <> 0) Then
MsgBox ("Loi. Tong so cau hoi va tra loi la " & socau)
Exit Sub
End If
Selection.HomeKey Unit:=wdStory
For j = 1 To socau / 2
With Selection.Find
.text = "%00"
.Replacement.text = " " & j & "."
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Execute Replace:=wdReplaceOne
End With
Next
Dim cautraloi As Integer
cautraloi = 9900
For j = socau / 2 + 1 To socau
With Selection.Find
cautraloi = cautraloi + 1
.text = "%00"
.Replacement.text = " " & cautraloi & "."
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Execute Replace:=wdReplaceOne
End With
Next
socau = socau / 2
For j = 1 To socau - 1
ThisDoc.Activate
Selection.HomeKey Unit:=wdStory
With Selection.Find
.ClearFormatting
.text = "(Câu)"
.Replacement.ClearFormatting
.Replacement.text = "^13\1"
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
.Execute Replace:=wdReplaceOne
End With
Selection.HomeKey Unit:=wdStory
With Selection.Find
.ClearFormatting
.text = "(^13)(Câu [0-9]{1,2}.)(*)(D.*)(^13)"
.Replacement.ClearFormatting
.Replacement.text = "\2\3\4\5"
.Replacement.Font.ColorIndex = wdBlue
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
.Execute Replace:=wdReplaceOne
End With
Selection.Cut
ThatDoc.Activate
Selection.PasteAndFormat (wdFormatOriginalFormatting)
ThisDoc.Activate
Selection.HomeKey Unit:=wdStory
With Selection.Find
.ClearFormatting
.text = "(Câu 99)"
.Replacement.ClearFormatting
.Replacement.text = "^13\1"
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
.Execute Replace:=wdReplaceOne
End With
With Selection.Find
.ClearFormatting
.text = "(Câu 99)([0-9]{2}.)(*)(^13)(Câu 99[0-9]{2}.)"
.Replacement.ClearFormatting
.Font.Name = "Times New Roman"
.Font.Size = 12
.Replacement.text = "^13" & "H" & ChrW(432) & ChrW(417) & ChrW(769) & "ng dâ" _
& ChrW(771) & "n gia" & ChrW(777) & "i" & "^13" & "Tra loi \2\3\4\5"
.Replacement.Font.Color = wdColorViolet
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
.Execute Replace:=wdReplaceOne
End With
Selection.MoveLeft Unit:=wdWord, Count:=3, Extend:=wdExtend
Selection.Cut
ThatDoc.Activate
Selection.PasteAndFormat (wdFormatOriginalFormatting)
ThatDoc.Range.InsertAfter Chr(13)
Next
ThisDoc.Activate
Selection.HomeKey Unit:=wdStory
With Selection.Find
.ClearFormatting
.text = "(Câu [0-9]{1,2}.)(*)(D.*)(^13)"
.Replacement.ClearFormatting
.Replacement.text = "\1\2\3\4"
.Replacement.Font.ColorIndex = wdBlue
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
.Execute Replace:=wdReplaceOne
End With
Selection.Cut
ThatDoc.Activate
Selection.PasteAndFormat (wdFormatOriginalFormatting)
ThisDoc.Activate
Selection.EndKey Unit:=wdStory
Selection.TypeParagraph
Selection.Font.Bold = True
Selection.TypeText text:="Câu 9999."
With Selection.Find
.ClearFormatting
.text = "(Câu 99[0-9]{2}.)(*)(Câu 9999.)"
.Replacement.ClearFormatting
.Replacement.text = "^13" & "H" & ChrW(432) & ChrW(417) & ChrW(769) & "ng dâ" _
& ChrW(771) & "n gia" & ChrW(777) & "i" & "^13" & "Tra loi \2"
.Replacement.Font.Color = wdColorViolet
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
.Execute Replace:=wdReplaceOne
End With
Selection.Cut
ThatDoc.Activate
Selection.PasteAndFormat (wdFormatOriginalFormatting)
ThatDoc.Range.InsertAfter Chr(13)
ThisDoc.Activate
Selection.WholeStory
Selection.Delete
TamDoc.Activate
Selection.WholeStory
Selection.Copy
ThisDoc.Activate
Selection.PasteAndFormat (wdListCombineWithExistingList)
Selection.EndKey Unit:=wdStory
Selection.TypeText text:="PHÂ" & ChrW(768) & "N CÂU HO" & ChrW(777) & _
"I CO" & ChrW(769) & " H" & ChrW(431) & ChrW(416) & ChrW(769) & "NG DÂ" & _
ChrW(771) & "N GIA" & ChrW(777) & "I"
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
Selection.Font.Bold = wdToggle
Selection.Font.Name = "Times New Roman"
Selection.Font.Size = 14
Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
Selection.MoveRight Unit:=wdCharacter, Count:=1
Selection.TypeParagraph
ThatDoc.Activate
Selection.WholeStory
Selection.Copy
ThisDoc.Activate
Selection.PasteAndFormat (wdListCombineWithExistingList)
TamDoc.Close (No)
ThatDoc.Close (No)
End Sub
