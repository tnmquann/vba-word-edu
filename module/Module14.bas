Attribute VB_Name = "Module14"
Sub BTPro_DPS(ByVal control As Office.IRibbonControl)
On Error Resume Next
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DH])([SH]1)(?.C?.?.???.)"
        .Forward = True
        .Replacement.text = "[" & "\2\4"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.???.)(a)(\])"
        .Forward = True
        .Replacement.text = "\1" & "1"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.???.)(b)(\])"
        .Forward = True
        .Replacement.text = "\1" & "2"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.???.)(c)(\])"
        .Forward = True
        .Replacement.text = "\1" & "3"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(.???.)(d)(\])"
        .Forward = True
        .Replacement.text = "\1" & "4"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DH])([012])(.C)([0-9]{1})(.)([0-9]{1}.D)([0-9]{2}.)([1234])"
        .Forward = True
        .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "\8" & "-" & "\9" & "]"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DH])([012])(.C)([0-9]{1})(.)([0-9]{1}.D0)([0-9]{1}.)([1234])"
        .Forward = True
        .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "\8" & "-" & "\9" & "]"
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".D0"
        .Forward = True
        .Replacement.text = "."
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".D"
        .Forward = True
        .Replacement.text = "."
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ".-"
        .Forward = True
        .Replacement.text = "-"
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
End Sub
