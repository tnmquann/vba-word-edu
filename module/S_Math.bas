Attribute VB_Name = "S_Math"
Option Explicit
Sub MathTypeConvert(ByVal control As Office.IRibbonControl)
 Call S_MTConvert
End Sub
Sub S_MTConvert()
    On Error GoTo S_Quit
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = " "
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Cut
    Application.Run MacroName:="MTCommand_InsertInlineEqn"
    SendKeys "^v"
    SendKeys "^a"
    SendKeys "^+="
    SendKeys "^{F4}"
Exit Sub
S_Quit:
Dim Title, msg As String
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a ch" & _
        ChrW(7885) & "n công th" & ChrW(7913) & "c c" & ChrW(7847) & "n chuy" & _
        ChrW(7875) & "n"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub
Sub Lechdong(ByVal control As Office.IRibbonControl)
On Error GoTo S_Quit
    Selection.ClearFormatting
    Application.Run MacroName:="MTCommand_FormatEqns"
Exit Sub
S_Quit:
Dim Title, msg As String
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a ch" & _
        ChrW(7885) & "n công th" & ChrW(7913) & "c c" & ChrW(7847) & "n chuy" & _
        ChrW(7875) & "n"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Sub
