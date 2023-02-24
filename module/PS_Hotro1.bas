Attribute VB_Name = "PS_Hotro1"
Sub TT_cau_Text_new(ByVal control As Office.IRibbonControl)
' Chuyen tu danh STT tu dong sang dang Text
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
End Sub
Sub TT_cau_Auto_new(ByVal control As Office.IRibbonControl)
    On Error Resume Next
    Key_cau.CheckBox4 = True
    Key_cau.Show
    Application.Visible = False
    If Key_cau.CheckBox4 = False And Key_cau.CheckBox5 = False And Key_cau.CheckBox6 = False Then Exit Sub
    Call Chay_TT_cau_Auto(Key_cau.CheckBox4, Key_cau.CheckBox5, Key_cau.CheckBox6, True)
    Application.Visible = True
    End
End Sub
Private Sub Chay_TT_cau_Auto(ByVal KeyCau As Boolean, KeyBai As Boolean, KeyNumber As Boolean, DemCau As Boolean)
    Application.ScreenUpdating = False
    On Error Resume Next
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    If KeyCau = True Then
        With Selection.Find
            .text = "(Câu [0-9]{1,4}[.:])"
            .Replacement.text = "#"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End If
    If KeyBai = True Then
        With Selection.Find
            .text = "(Bài [0-9]{1,4}[.:])"
            .Replacement.text = "#"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End If
    If KeyNumber = True Then
        With Selection.Find
            .text = "(^13)([0-9]{1,4}[/.:)])"
            .Replacement.text = "\1" & "#"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End If
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = False
    If Selection.Find.Execute = False Then Exit Sub
    End With
    With Selection.Find
        .text = "#^t"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "# "
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Set danhsach = ActiveDocument.Content
    socau = 0
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
        socau = socau + 1
        danhsach.Select
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.ParagraphFormat.TabStops.ClearAll
            ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
                , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
            .NumberFormat = "Câu " & "%1."
            .TrailingCharacter = wdTrailingTab
            .NumberStyle = wdListNumberStyleArabic
            .NumberPosition = CentimetersToPoints(0)
            .Alignment = wdListLevelAlignLeft
            .TextPosition = CentimetersToPoints(1.75)
            .TabPosition = wdUndefined
            .ResetOnHigher = 0
            .StartAt = 1
            .LinkedStyle = ""
            .Font.Bold = True
            .Font.Color = wdColorGreen
        End With
        ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
        Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
        Selection.Delete Unit:=wdCharacter, Count:=1
        GoTo Tiep
    Else
        Selection.HomeKey Unit:=wdStory
        Application.ScreenUpdating = True
        Selection.HomeKey Unit:=wdStory
        If DemCau = True Then
            msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t." & vbCrLf & "S" & ChrW(7889) & " c" & ChrW(226) & "u " & ChrW(273) & "" & ChrW(227) & " chuy" & ChrW(7875) & "n: " & socau
            Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
        End If
        ActiveDocument.Save
        Exit Sub
    End If
End Sub
Sub Chuan_hoa_BTN_new(ByVal control As Office.IRibbonControl)
    On Error Resume Next
    BTN.Show
    ActiveDocument.UndoClear
End Sub
Private Sub Chuan_hoa()
    On Error Resume Next
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Application.ScreenUpdating = False
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
' Xoa ky hieu thua cua file goc
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .text = "^l"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "  "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "  "
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "( )([.:,;\?])"
        .Replacement.text = "\2"
        .Replacement.Font.Underline = wdUnderlineNone
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^p "
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
' Giu lai dinh dang gach chan cho Dap an
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([A-D])"
        .Replacement.text = "\1"
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
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
'Xu ly dinh dang cac phuong an (co chu y tranh nhan dang nham noi dung de)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = wdColorGreen
    End With
    With Selection.Find
        .text = "([A-D].)"
        .Replacement.text = "\1\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "( [A-D].)"
        .Replacement.text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pA."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pB."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pD."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Underline = wdUnderlineNone
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "A.A. "
        .Replacement.text = "A."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "B.B. "
        .Replacement.text = "B."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "C.C. "
        .Replacement.text = "C."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "D.D. "
        .Replacement.text = "D."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
' Xoa ky hieu thua phat sinh
    With Selection.Find
        .ClearFormatting
        .text = "  "
        .Replacement.ClearFormatting
        .Replacement.text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = " ."
        .Replacement.ClearFormatting
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = ";.^t"
        .Replacement.ClearFormatting
        .Replacement.text = ".^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = ";^t"
        .Replacement.ClearFormatting
        .Replacement.text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = ".."
        .Replacement.ClearFormatting
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .ClearFormatting
        .text = " ^p"
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
' Tra lai dinh dang ban dau cho bang
  If ActiveDocument.Tables.Count > 0 Then
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
    End With
    Next i
  End If
End Sub
Private Sub End_cau()
    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Dim Loi As Integer, i As Integer
    Loi = 0
    For i = 1 To ActiveDocument.Tables.Count
        ActiveDocument.Tables(i).Select
        With Selection
            Options.DefaultHighlightColorIndex = wdYellow
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Highlight = True
            With Selection.Find
                .text = "(Câu[^32^s][0-9]{1,4}*\[)([0-2][DHL][1-9]*-[1-9]\])"
                .Replacement.text = "\1\2"
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = True
                .MatchWildcards = True
            If Selection.Find.Execute = True Then
                .Execute Replace:=wdReplaceOne
                Loi = Loi + 1
            End If
            End With
        End With
    Next i
    If Loi <> 0 Then
        Exit Sub
    Else
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.zz^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "^13 "
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .ClearFormatting
        .text = "^13^13"
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "(Câu[^32^s][0-9]{1,4}[.:])"
        .Replacement.text = "z.zz^p\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText text:="z.zz"
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.zz^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceOne
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Size = 12
        .Color = wdColorGreen
    End With
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = "z.zz"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    End If
End Sub
