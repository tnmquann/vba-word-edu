Attribute VB_Name = "Module8"
Sub HideGiai(ByVal control As Office.IRibbonControl)
ActiveDocument.ConvertNumbersToText
Selection.WholeStory
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.text = "(L?i gi?i)"
.Replacement.text = "#\1"
.Forward = False
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.text = "(Câu [0-9]{1,2})"
.Replacement.text = "~\1"
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWholeWord = True
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
Selection.HomeKey Unit:=wdStory
With Selection.Find
.text = "~"
.Replacement.text = ""
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
Selection.EndKey Unit:=wdStory
Selection.TypeParagraph
Selection.TypeText "~"
With Selection.Find
.text = "(#*~)"
.Replacement.text = "\1"
.Replacement.Font.Hidden = True
.Forward = False
.Wrap = wdFindContinue
.Format = True
.MatchCase = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
End Sub
