Attribute VB_Name = "Module18"
Sub Chuan_hoa_0106(ByVal control As Office.IRibbonControl)
    On Error Resume Next
    Application.ScreenUpdating = True
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.WholeStory
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0)
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Font.ColorIndex = wdRed
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
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
        .text = "^p^t"
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
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "([.:,\)])( )"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "([^13^32^9])([Aa])([.:\)])"
        .Replacement.text = "#A."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Bb])([.:\)])"
        .Replacement.text = "#B."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Cc])([.:\)])"
        .Replacement.text = "#C."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Dd])([.:\)])"
        .Replacement.text = "#D."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
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
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    
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
       .Bold = True
       .Color = wdColorBlue
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
       .Bold = True
       .Color = wdColorBlue
       .Italic = False
    End With
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "(^13)([ABCD])"
        .Replacement.text = "\1\2"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "#A."
        .Replacement.text = "^pA. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#B."
        .Replacement.text = "^tB. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
        .text = "^pB."
        .Replacement.text = "^pB. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#C."
        .Replacement.text = "^tC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^pC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#D."
        .Replacement.text = "^tD. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pD."
        .Replacement.text = "^pD. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
       .Bold = False
       .Color = wdColorBlack
    End With
    Selection.Find.ClearFormatting
   
    With Selection.Find
        .text = ".^t"
        .Replacement.text = ".^t"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([ABCD])(. .)"
        .Replacement.text = "\1" & ". "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".."
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    
        Selection.WholeStory
        Selection.ParagraphFormat.TabStops.ClearAll
        ActiveDocument.DefaultTabStop = CentimetersToPoints(0.5)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(0.5) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(9.5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(14), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
 
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
    
    
    Set danhsach = ActiveDocument.Content
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
    danhsach.Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.ParagraphFormat.TabStops.ClearAll
    'ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
    'Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
    .NumberFormat = "Câu " & "%1."
    .TrailingCharacter = wdTrailingTab
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = CentimetersToPoints(0)
    .Alignment = wdListLevelAlignLeft
    .TextPosition = CentimetersToPoints(0)
    .TabPosition = wdUndefined
    .ResetOnHigher = 0
    .StartAt = 1
    .LinkedStyle = ""
    .Font.Bold = True
    .Font.Color = wdColorBlue
    .Font.Italic = False
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
    True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
    wdWord10ListBehavior
    Selection.Delete Unit:=wdCharacter, Count:=1
    GoTo Tiep
    End If
   
    Selection.WholeStory
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
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
        Selection.WholeStory
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 12
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    If ActiveDocument.Tables.Count > 0 Then
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
        With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0)
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        End With
    Next i
    ActiveDocument.Tables(1).Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
    Selection.HomeKey Unit:=wdStory
    'If ktBanQuyen = False Then Call S_SerialHDD
    'If ktBanQuyen = False Then S_NoteRig.Show
End Sub

