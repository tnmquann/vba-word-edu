Attribute VB_Name = "Lay_ma_TV"
Sub chuyen_ma_tieng_Viet()
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "&#"
        .Replacement.text = """ & ChrW("
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
        .text = ";"
        .Replacement.text = ") & """
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
    Selection.EndKey Unit:=wdStory
    Application.Keyboard (1066)
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeParagraph
    Application.Keyboard (1033)
    Selection.TypeText text:="& VbCrLf"
End Sub


