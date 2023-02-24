Attribute VB_Name = "Module17"
Sub dasmarttest(ByVal control As Office.IRibbonControl)
On Error Resume Next
    Selection.HomeKey Unit:=wdStory
           With Selection.Find
        .text = "A.."
        .Replacement.text = "A."
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
      With Selection.Find
        .text = "B.."
        .Replacement.text = "B."
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
      With Selection.Find
        .text = "C.."
        .Replacement.text = "C."
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
      With Selection.Find
        .text = "D.."
        .Replacement.text = "D."
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
      Selection.HomeKey Unit:=wdStory
End Sub
