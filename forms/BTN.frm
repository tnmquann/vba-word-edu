VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BTN 
   Caption         =   "Chuan hoa nhom toan hoc BTN"
   ClientHeight    =   1455
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7815
   OleObjectBlob   =   "BTN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BTN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    BTN.Hide
    Call Chuan_hoa_BTN0
    Call Chuan_hoa_BTN1
    Call Chuan_hoa_BTN3
    Call DapanBTN
End Sub
Private Sub CommandButton2_Click()
    BTN.Hide
    Call Chuan_hoa_BTN0
    Call Chuan_hoa_BTN2
    Call Chuan_hoa_BTN3
    Call DapanBTN
End Sub
Private Sub CommandButton3_Click()
    BTN.Hide
    Exit Sub
End Sub
Private Sub Chuan_hoa_BTN0()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Application.ScreenUpdating = False
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
    End With
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.75)
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = " "
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
        .text = "  "
        .Replacement.text = " "
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
        .text = "  "
        .Replacement.text = " "
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
        .text = "( )([.:,;\?])"
        .Replacement.text = "\2"
        .Replacement.Font.Underline = wdUnderlineNone
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^p "
        .Replacement.text = "^p"
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
        .text = "^p^p"
        .Replacement.text = "^p"
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
        .Color = wdColorBlue
    End With
    With Selection.Find
        .text = "([A-D].)"
        .Replacement.text = "\1\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Private Sub Chuan_hoa_BTN1()
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    With Selection.Find
        .text = "( [A-D].)"
        .Replacement.text = ".^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pA."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pB."
        .Replacement.text = ".^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pC."
        .Replacement.text = ".^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pD."
        .Replacement.text = ".^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pCâu "
        .Replacement.text = ".^pCâu "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .text = "(.[^9^13])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Private Sub Chuan_hoa_BTN2()
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "( [A-D].)"
        .Replacement.text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pA."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pB."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^pD."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Private Sub Chuan_hoa_BTN3()
    Application.ScreenUpdating = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "A.A. "
        .Replacement.text = "A."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "B.B. "
        .Replacement.text = "B."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "C.C. "
        .Replacement.text = "C."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "D.D. "
        .Replacement.text = "D."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
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
' Canh Tab cho cac phuong an
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineNone
        .Color = wdColorBlue
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.ParagraphFormat.TabStops.ClearAll
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(6), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(10), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    Selection.Find.Replacement.ParagraphFormat.TabStops.add Position:= _
        CentimetersToPoints(14), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Dinh dang lai trang in

    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(1.5)
        .LeftMargin = CentimetersToPoints(1.5)
        .RightMargin = CentimetersToPoints(1.5)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.6)
        .FooterDistance = CentimetersToPoints(0.6)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    
' Danh lai thu tu cua cac cau hoi (danh tu dong)
   
    Selection.Find.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
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
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
    danhsach.Select
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.ParagraphFormat.TabStops.ClearAll
        ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "Câu " & "%1:"
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
        .Font.Color = wdColorBlue
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
    True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
    wdWord10ListBehavior
    Selection.Delete Unit:=wdCharacter, Count:=1
GoTo Tiep
    Else
    End If
Application.ScreenUpdating = True
    msg = "V" & ChrW(259) & "n b" & ChrW(7843) & "n c" & ChrW(7911) & "a b" & ChrW(7841) & "n c" & ChrW(417) & " b" & ChrW(7843) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c " & ChrW(273) & "" & ChrW(7883) & "nh d" & ChrW(7841) & "ng theo chu" & ChrW(7849) & "n " & ChrW(273) & "" & ChrW(7883) & "nh d" & ChrW(7841) & "ng c" & ChrW(7911) & "a BTN"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
Selection.HomeKey Unit:=wdStory
End Sub
Private Sub Dapan()
    With Selection.Find
        .text = "Ch" & ChrW(7885) & "n^tA."
        .Replacement.text = "Ch" & ChrW(7885) & "n A"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
       With Selection.Find
        .text = "Ch" & ChrW(7885) & "n^tB."
        .Replacement.text = "Ch" & ChrW(7885) & "n B"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .text = "Ch" & ChrW(7885) & "n^tC."
        .Replacement.text = "Ch" & ChrW(7885) & "n C"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .text = "Ch" & ChrW(7885) & "n^tD."
        .Replacement.text = "Ch" & ChrW(7885) & "n D"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Private Sub DapanBTN()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = False
        .Color = wdColorBlue
    End With
    With Selection.Find.Replacement.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
    End With
    Call Dapan
Dim BTNFindText As String
Selection.HomeKey wdStory
BTNFindText = "Ch" & ChrW(7885) & "n A"
Selection.Find.Execute BTNFindText
Do Until Selection.Find.Found = False
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Range.HighlightColorIndex = wdYellow
        Selection.ParagraphFormat.TabStops.ClearAll
        Selection.MoveRight
        Selection.Find.Execute
Loop
Selection.HomeKey wdStory
BTNFindText = "Ch" & ChrW(7885) & "n B"
Selection.Find.Execute BTNFindText
Do Until Selection.Find.Found = False
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Range.HighlightColorIndex = wdYellow
        Selection.ParagraphFormat.TabStops.ClearAll
        Selection.MoveRight
        Selection.Find.Execute
Loop
Selection.HomeKey wdStory
BTNFindText = "Ch" & ChrW(7885) & "n C"
Selection.Find.Execute BTNFindText
Do Until Selection.Find.Found = False
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Range.HighlightColorIndex = wdYellow
        Selection.ParagraphFormat.TabStops.ClearAll
        Selection.MoveRight
        Selection.Find.Execute
Loop
Selection.HomeKey wdStory
BTNFindText = "Ch" & ChrW(7885) & "n D"
Selection.Find.Execute BTNFindText
Do Until Selection.Find.Found = False
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Range.HighlightColorIndex = wdYellow
        Selection.ParagraphFormat.TabStops.ClearAll
        Selection.MoveRight
        Selection.Find.Execute
Loop
End Sub



    
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
