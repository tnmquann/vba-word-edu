Attribute VB_Name = "Module20"
Sub xoaid4(ByVal control As Office.IRibbonControl)
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[[012][DH][0-9]{1,2}-[1-4]\])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
