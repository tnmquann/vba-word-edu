Attribute VB_Name = "PS_Hotro2"
Sub Tach_lay_cau_hoi_new(ByVal control As Office.IRibbonControl)
' Chep cau hoi trac nghiem ra mot file moi (khong chep HDG)
Application.Visible = False
On Error Resume Next
Application.ScreenUpdating = False
    Dim FormDoc As Document, ThisDoc As Document, ThatDoc As Document
    Dim oData As New DataObject
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Set FormDoc = ActiveDocument
    Selection.WholeStory
    Selection.Copy
    Set ThisDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    FormDoc.Close (No)
    ThisDoc.Activate
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
    ActiveDocument.UndoClear
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
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    Selection.Find.Replacement.ParagraphFormat.TabStops.ClearAll
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(1.75), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(6), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(10), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(14), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    With Selection.Find
        .text = "(A.*)(B.*)(C.*)(D.*)(^13)"
        .Replacement.text = "\1\2\3\4\5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
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
    oData.SetText text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
    
    ThisDoc.Activate
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(-1.75)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
    End With
    Selection.Find.Replacement.ParagraphFormat.TabStops.ClearAll
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(1.75), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])(*)(A.*)(D.*)(^13)"
        .Replacement.text = "\1^9\5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(^13Hý" & ChrW(7899) & "ng d" & ChrW(7851) & "n*^13)"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(^13L" & ChrW(7901) & "i gi" & ChrW(7843) & "i*^13)"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    ThatDoc.Activate
    msg = "C" & ChrW(243) & " 2 file m" & ChrW(7899) & "i " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c t" & ChrW(7841) & "o ra" & vbCrLf & "+ File 1: ch" & ChrW(7913) & "a " & ChrW(273) & "" & ChrW(7873) & " b" & ChrW(224) & "i" & vbCrLf & "+ File 2: ch" & ChrW(7913) & "a l" & ChrW(7901) & "i gi" & ChrW(7843) & "i" & vbCrLf & "B" & ChrW(7841) & "n nh" & ChrW(7899) & " Save l" & ChrW(7841) & "i nh" & ChrW(233) & " !"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
    Application.Visible = True
End Sub
Sub Tao_bang_dap_an_new(ByVal control As Office.IRibbonControl)
    On Error Resume Next
    GachDA.CheckBox1.Enabled = False
    GachDA.CheckBox2.Enabled = False
    GachDA.CheckBox3.Enabled = False
    GachDA.Label2.Enabled = False
    GachDA.Label3.Enabled = False
    GachDA.Label4.Enabled = False
    GachDA.Label5.Enabled = False
    GachDA.Label6.Enabled = False
    GachDA.Label7.Enabled = False
    GachDA.Label8.Enabled = False
    GachDA.Label9.Enabled = False
    GachDA.Label10.Enabled = False
    GachDA.CheckBox4.Enabled = False
    GachDA.CheckBox5.Enabled = False
    GachDA.CheckBox6.Enabled = False
    GachDA.TextBox1.Enabled = False
    GachDA.TextBox2.Enabled = False
    GachDA.TextBox3.Enabled = False
    GachDA.Label12.Enabled = False
    GachDA.Label18.Enabled = False
    GachDA.Show
    ActiveDocument.UndoClear
    End
End Sub
Sub Danh_dau_dap_an_new(ByVal control As Office.IRibbonControl)
On Error Resume Next
Application.Visible = False
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Dim CountImage
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
        .Replacement.text = "#\1#\2#\3#\4"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
' Chep bang dap an sang file tam và Text hoa no
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
    Dim oData   As New DataObject 'object to use the clipboard
    oData.SetText text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
ActiveDocument.UndoClear
    msg1 = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg1, 0, 4, 0, 0, 0
    Application.Visible = True
End Sub
Sub GhepHoi_Dap_new(ByVal control As Office.IRibbonControl)
On Error Resume Next
ActiveDocument.Range.ListFormat.ConvertNumbersToText
' Kiem tra xem co cau nao nam trong Table hay khong de bao loi
    Application.Visible = False
    Loi = 0
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
        With Selection
            Options.DefaultHighlightColorIndex = wdYellow
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Highlight = True
            With Selection.Find
                .text = "(Câu [0-9]{1,4}[.:])"
                .Replacement.text = "\1"
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .MatchCase = True
                .MatchWildcards = True
            If Selection.Find.Execute = True Then
                .Execute Replace:=wdReplaceOne
                Loi = Loi + 1
            End If
            End With
        End With
    Next i
    If Loi <> 0 Then
        'Danh dau cau hoi nam trong Table (neu co)
        Options.DefaultHighlightColorIndex = wdYellow
        Selection.Find.ClearFormatting
        Selection.Find.Highlight = True
        With Selection.Find
            .text = "(Câu [0-9]{1,4}[.:])"
            .Replacement.text = "\1"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = True
            .MatchWildcards = True
        End With
        Selection.Find.Execute
        Selection.HomeKey Unit:=wdLine
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg = "B" & ChrW(7841) & "n ph" & ChrW(7843) & "i chuy" & ChrW(7875) & "n t" & ChrW(7915) & " c" & ChrW(7909) & "m t" & ChrW(7915) & " " & ChrW(8220) & "C" & ChrW(226) & "u xx " & ChrW(8230) & "[lMc" & ChrW(8230) & "]" & ChrW(8221) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c " & ChrW(273) & "" & ChrW(225) & "nh d" & ChrW(7845) & "u n" & ChrW(7873) & "n v" & ChrW(224) & "ng" & vbCrLf & "ra ngo" & ChrW(224) & "i table tr" & ChrW(432) & "" & ChrW(7899) & "c khi ch" & ChrW(7841) & "y ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh t" & ChrW(225) & "ch c" & ChrW(226) & "u h" & ChrW(7887) & "i tr" & ChrW(7855) & "c nghi" & ChrW(7879) & "m!"
        Application.Assistant.DoAlert Title, msg, 0, 1, 0, 0, 0
        Application.Visible = True
        End
    Else
        'Neu Cau hoi nam hoan toan ngoai Table thi chay chuong trinh
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .text = "z.zz^13"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .text = "^13 "
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
            .ClearFormatting
            .text = "^13^13"
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
            .text = "(Câu [0-9]{1,4}[.:])"
            .Replacement.text = "z.zz^p\1"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        Selection.TypeText text:="z.zz"
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .text = "z.zz^13"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceOne
        End With
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.Bold = True
        Selection.Find.Replacement.Font.Size = 12
        Selection.Find.Replacement.Font.Color = wdColorGreen
        With Selection.Find
            .text = "z.zz"
            .Replacement.text = "z.zz"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        Application.ScreenUpdating = True
    End If
' Bat dau xu ly ghep De bai va Dap an lai voi nhau
    Set ThisDoc = ActiveDocument
    Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    For i = 1 To 9999
        ThisDoc.Activate
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .text = "Câu " & i & "([.:])(*z.zz)"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        If Selection.Find.Execute = True Then
            Do
                Selection.Cut
                ThatDoc.Activate
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
                Selection.EndKey Unit:=wdStory
                Selection.TypeParagraph
                Call ClearClipBoard
                ThisDoc.Activate
            Loop While .Execute
        Else
            i = 9999
        End If
        End With
        ThatDoc.Activate
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .text = "(Câu )" & i & "([.:]*)(z.zz^13Câu )" & i & "([.:][^32^t])"
            .Replacement.text = "\1" & i & "\2" & "H" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i^p"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        Selection.EndKey Unit:=wdStory
    Next i
    ThatDoc.Activate
    Selection.WholeStory
    Selection.Cut
    ThisDoc.Activate
    With Selection.Find
        .ClearFormatting
        .text = "^13^13"
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.EndKey Unit:=wdStory
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    ThatDoc.Close (No)
    ThisDoc.Activate
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "H" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i" & "(*)(z.zz^13)"
        .Replacement.text = "H" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i" & "\1"
        .Replacement.ParagraphFormat.LeftIndent = CentimetersToPoints(1.75)
        .Replacement.ParagraphFormat.RightIndent = CentimetersToPoints(0)
        .Replacement.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
        .Replacement.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        .Replacement.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(A.*B.*C.*D.*^13)"
        .Replacement.text = "\1"
        .Replacement.ParagraphFormat.LeftIndent = CentimetersToPoints(1.75)
        .Replacement.ParagraphFormat.RightIndent = CentimetersToPoints(0)
        .Replacement.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
        .Replacement.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        .Replacement.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
        .Replacement.ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = "\1"
        .Replacement.ParagraphFormat.LeftIndent = CentimetersToPoints(1.75)
        .Replacement.ParagraphFormat.RightIndent = CentimetersToPoints(0)
        .Replacement.ParagraphFormat.FirstLineIndent = CentimetersToPoints(-1.75)
        .Replacement.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        .Replacement.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
        .Replacement.ParagraphFormat.Alignment = wdAlignParagraphJustify
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "H" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i"
        .Replacement.text = "H" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i"
        .Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Replacement.ParagraphFormat.LeftIndent = CentimetersToPoints(1.75)
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "z.zz^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Application.Visible = True
    Application.ScreenUpdating = True
    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg2 = "Ph" & ChrW(7847) & "n " & ChrW(273) & "" & ChrW(7873) & " b" & ChrW(224) & "i v" & ChrW(224) & " " & ChrW(272) & "" & ChrW(225) & "p " & ChrW(225) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c gh" & ChrW(233) & "p l" & ChrW(7841) & "i v" & ChrW(7899) & "i nhau" & vbCrLf & "B" & ChrW(7841) & "n n" & ChrW(234) & "n ki" & ChrW(7875) & "m tra l" & ChrW(7841) & "i tr" & ChrW(432) & "" & ChrW(7899) & "c khi Save v" & ChrW(224) & " " & ChrW(273) & "" & ChrW(243) & "ng file."
    Application.Assistant.DoAlert title2, msg2, 0, 4, 0, 0, 0
End Sub
Private Sub ClearClipBoard()
Dim oData   As New DataObject 'object to use the clipboard
    oData.SetText text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
End Sub
