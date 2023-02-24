Attribute VB_Name = "Module16"
Sub sang_do(ByVal control As Office.IRibbonControl)
Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorRed
     With Selection.Find
        .text = "([A-D])"
        .Replacement.text = "\1."
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
        .Replacement.text = "\1."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
     Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = ""
        .Replacement.text = ""
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
                 
End Sub

