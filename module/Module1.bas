Attribute VB_Name = "Module1"
Sub TT_cau_Text(ByVal control As Office.IRibbonControl)
' Chuyen tu danh STT tu dong sang dang Text
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
End Sub
Sub TT_cau_Auto(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .text = "(Câu [0-9]{1,4}[.:])"
    .Replacement.text = "#"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
End With
With Selection.Find
    .text = "(^13)([0-9]{1,4}[/.:)])"
    .Replacement.text = "\1" & "#"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
End With
Selection.Find.ClearFormatting
With Selection.Find
    .text = "#"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWildcards = False
    If Selection.Find.Execute = False Then Exit Sub
End With
With Selection.Find
    .text = "#^t"
    .Replacement.text = "#"
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
    .text = "# "
    .Replacement.text = "#"
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
Set danhsach = ActiveDocument.Content
Tiep:
danhsach.Find.Execute FindText:="#", Forward:=True
If danhsach.Find.Found = True Then
danhsach.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
.NumberFormat = "Câu " & "%1:"
.TrailingCharacter = wdTrailingTab
.NumberStyle = wdListNumberStyleArabic
.NumberPosition = CentimetersToPoints(0)
.Alignment = wdListLevelAlignLeft
.TextPosition = CentimetersToPoints(1.75)
.TabPosition = wdUndefined
.ResetOnHigher = 0
.StartAt = 1
.LinkedStyle = ""
.Font.Bold = True
.Font.Color = wdColorBlue
End With
ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
wdWord10ListBehavior
Selection.Delete Unit:=wdCharacter, Count:=1
GoTo Tiep
Else
Selection.HomeKey Unit:=wdStory
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c chuy" & ChrW(7875) & "n th" & ChrW(7913) & " t" & ChrW(7921) & " c" & ChrW(226) & "u sang d" & ChrW(7841) & "ng t" & ChrW(7921) & " " & ChrW(273) & "" & ChrW(7897) & "ng " & ChrW(273) & "" & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
Exit Sub
End If
End Sub
Sub Sap_lai_TT_cau(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .text = "(Câu [0-9]{1,4}[.:])"
    .Replacement.text = "#"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
End With
With Selection.Find
    .text = "(^13)([0-9]{1,4}[/.:)])"
    .Replacement.text = "\1" & "#"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
End With
Selection.Find.ClearFormatting
With Selection.Find
    .text = "#"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWildcards = False
    If Selection.Find.Execute = False Then Exit Sub
End With
With Selection.Find
    .text = "#^t"
    .Replacement.text = "#"
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
    .text = "# "
    .Replacement.text = "#"
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
Set danhsach = ActiveDocument.Content
Tiep:
danhsach.Find.Execute FindText:="#", Forward:=True
If danhsach.Find.Found = True Then
danhsach.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
.NumberFormat = "Câu " & "%1:"
.TrailingCharacter = wdTrailingTab
.NumberStyle = wdListNumberStyleArabic
.NumberPosition = CentimetersToPoints(0)
.Alignment = wdListLevelAlignLeft
.TextPosition = CentimetersToPoints(1.75)
.TabPosition = wdUndefined
.ResetOnHigher = 0
.StartAt = 1
.LinkedStyle = ""
.Font.Bold = True
.Font.Color = wdColorBlue
End With
ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
wdWord10ListBehavior
Selection.Delete Unit:=wdCharacter, Count:=1
GoTo Tiep
Else
ActiveDocument.Range.ListFormat.ConvertNumbersToText
With Selection.Find
        .text = ":^t"
        .Replacement.text = ": "
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
Selection.HomeKey Unit:=wdStory
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(225) & "c c" & ChrW(226) & "u h" & ChrW(7887) & "i " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c s" & ChrW(7855) & "p x" & ChrW(7871) & "p l" & ChrW(7841) & "i theo th" & ChrW(7913) & " t" & ChrW(7921)
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
Exit Sub
End If
End Sub
Sub Xoa_duong_ke_bang(ByVal control As Office.IRibbonControl)
If ActiveDocument.Tables.Count = 0 Then Exit Sub
For i = 1 To ActiveDocument.Tables.Count
ActiveDocument.Tables(1).Select
Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=True
Next i
    msg = "T" & ChrW(7845) & "t c" & ChrW(7843) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7901) & "ng k" & ChrW(7867) & " b" & ChrW(7843) & "ng c" & ChrW(243) & " trong v" & ChrW(259) & "n b" & ChrW(7843) & "n " & ChrW(273) & "" & ChrW(7873) & "u " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c xo" & ChrW(225) & " m" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Chuan_hoa_BTN(ByVal control As Office.IRibbonControl)
    BTN.Show
End Sub
Sub Chuan_hoa_VDC(ByVal control As Office.IRibbonControl)
    VDC.Show
End Sub
Sub Tiet_kiem_giay_A5(ByVal control As Office.IRibbonControl)
    Call Chuan_hoa
    Call Canh_Tab_PageSetup_A5_A5
End Sub
Sub Tiet_kiem_giay_A4(ByVal control As Office.IRibbonControl)
    Call Chuan_hoa
    Call Canh_Tab_PageSetup_A4
End Sub
Private Sub Chuan_hoa()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Application.ScreenUpdating = False
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
' Xoa ky hieu thua cua file goc
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .text = "^l"
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
        .text = "^t"
        .Replacement.text = " "
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
        .text = "( )([.:,;\?])"
        .Replacement.text = "\2"
        .Replacement.Font.Underline = wdUnderlineNone
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
' Giu lai dinh dang gach chan cho Dap an
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([A-D])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([A-D])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
'Xu ly dinh dang cac phuong an (co chu y tranh nhan dang nham noi dung de)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    With Selection.Find
        .text = "([A-D].)"
        .Replacement.text = "\1\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "( [A-D].)"
        .Replacement.text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pA."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pB."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pD."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Underline = wdUnderlineNone
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = "^t"
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
    With Selection.Find
        .text = "A.A. "
        .Replacement.text = "A."
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
    With Selection.Find
        .text = "B.B. "
        .Replacement.text = "B."
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
    With Selection.Find
        .text = "C.C. "
        .Replacement.text = "C."
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
    With Selection.Find
        .text = "D.D. "
        .Replacement.text = "D."
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
' Xoa ky hieu thua phat sinh
    With Selection.Find
        .ClearFormatting
        .text = "  "
        .Replacement.ClearFormatting
        .Replacement.text = " "
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
    With Selection.Find
        .ClearFormatting
        .text = " ."
        .Replacement.ClearFormatting
        .Replacement.text = "."
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
    With Selection.Find
        .ClearFormatting
        .text = ";.^t"
        .Replacement.ClearFormatting
        .Replacement.text = ".^t"
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
    With Selection.Find
        .ClearFormatting
        .text = ";^t"
        .Replacement.ClearFormatting
        .Replacement.text = "^t"
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
    With Selection.Find
        .ClearFormatting
        .text = ".."
        .Replacement.ClearFormatting
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
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
        .ClearFormatting
        .text = " ^p"
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
' Tra lai dinh dang ban dau cho bang
  If ActiveDocument.Tables.Count > 0 Then
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
    End With
    Next i
  End If
End Sub
Private Sub Canh_Tab_PageSetup_A5_A5()
' Canh Tab cho cac phuong an
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineNone
        .Color = wdColorBlue
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.1)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(-0.6)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
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
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.ParagraphFormat.TabStops.ClearAll
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(3.7), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(6.9), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(10.1), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
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
' Dinh dang lai trang in
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(0.5)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.6)
        .FooterDistance = CentimetersToPoints(0.6)
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
        .TwoPagesOnOne = True
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
' Danh lai thu tu cua cac cau hoi (danh tu dong)
    Selection.Find.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}.)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}:)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "#^t"
        .Replacement.text = "#"
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
        .text = "# "
        .Replacement.text = "#"
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
    Set danhsach = ActiveDocument.Content
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
    danhsach.Select
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.ParagraphFormat.TabStops.ClearAll
        ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "Câu %1:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Underline = wdUnderlineSingle
            .Color = wdColorBlue
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    Selection.Delete Unit:=wdCharacter, Count:=1
GoTo Tiep
    Else
    End If
    Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Private Sub Canh_Tab_PageSetup_A4()
' Canh Tab cho cac phuong an
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineNone
        .Color = wdColorBlue
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.1)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(-0.6)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
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
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.ParagraphFormat.TabStops.ClearAll
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(4.85), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(9.2), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(13.55), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
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
' Dinh dang lai trang in
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.6)
        .FooterDistance = CentimetersToPoints(0.6)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = True
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
' Danh lai thu tu cua cac cau hoi (danh tu dong)
    Selection.Find.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}.)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}:)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "#^t"
        .Replacement.text = "#"
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
        .text = "# "
        .Replacement.text = "#"
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
    Set danhsach = ActiveDocument.Content
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
    danhsach.Select
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.ParagraphFormat.TabStops.ClearAll
        ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "Câu %1:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Underline = wdUnderlineSingle
            .Color = wdColorBlue
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    Selection.Delete Unit:=wdCharacter, Count:=1
GoTo Tiep
    Else
    End If
    Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
