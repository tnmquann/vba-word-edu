Attribute VB_Name = "Module7"
Sub Sua_Math_lech_dong(ByVal control As Office.IRibbonControl)
Selection.Font.Position = 0 'Chon Position = Normal
Application.Run MacroName:="MTCommand_TeXToggle"
With Selection.Find
.text = "\["
.Replacement.text = "${"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchWildcards = False
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.text = "\]"
.Replacement.text = "}$"
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchWildcards = False
.Execute Replace:=wdReplaceAll
End With
Application.Run MacroName:="MTCommand_TeXToggle"
End Sub
