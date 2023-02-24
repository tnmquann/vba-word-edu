Attribute VB_Name = "showgiai"
Sub showgiai(ByVal control As Office.IRibbonControl)
Selection.WholeStory
Selection.Font.Hidden = False
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.text = "#"
.Replacement.text = ""
.Forward = False
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWildcards = False
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.text = "~"
.Replacement.text = ""
.Forward = False
.Wrap = wdFindContinue
.Format = False
.MatchCase = False
.MatchWildcards = False
.Execute Replace:=wdReplaceAll
End With
End Sub

