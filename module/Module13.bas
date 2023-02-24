Attribute VB_Name = "Module13"
Sub DPS_BTPro(ByVal control As Office.IRibbonControl)
On Error Resume Next
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "\[([012])([DH])([0-9]{1,4})-([0-9]{1,4}).([0-9]{1,4})-([1-4])\]"
        .Replacement.text = "[\2H1\1.C\3.\4.D\5.\6]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = "\[DH1([012].C[0-9]{1,4}.)"
        .Replacement.text = "[DS1\1"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".D([1-9].[1-4]\])"
        .Replacement.text = ".D0\1"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".1]"
        .Replacement.text = ".a]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".2]"
        .Replacement.text = ".b]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".3]"
        .Replacement.text = ".c]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .text = ".4]"
        .Replacement.text = ".d]"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineNone
        .Color = wdColorPink
    End With
    With Selection.Find
        .text = "(\[[DH]*.[abcd]\])"
        .Replacement.text = "\1"
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    Selection.HomeKey Unit:=wdStory
End Sub

