Attribute VB_Name = "Module4"
Sub Tim_va_Highlight(ByVal control As Office.IRibbonControl)
    keysearch.Show
    Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "~.|.~"
        .MatchWildcards = False
    If Selection.Find.Execute = False Then
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
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
            .text = "z.zz"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        Exit Sub
    End If
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)"
        .Replacement.ClearFormatting
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)" & "  "
        .Replacement.ClearFormatting
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    With Selection.Find
        .text = "~.|.~" & "(Câu [0-9]{1,4}[.:]*)(A.*)(B.*)(C.*)(D.*)(z.zz)"
        .Replacement.text = "\1\2\3\4\5\6"
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
    With Selection.Find
        .ClearFormatting
        .text = "~.|.~"
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
        .text = "z.zz"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert Title, msg, 0, 8, 0, 0, 0
End Sub
Sub Chep_cau_HighLight(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
    ChonCopy.Show
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.zz^13"
        .MatchWildcards = False
    If Selection.Find.Execute = False Then Exit Sub
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .text = "Câu "
        .Replacement.text = "~.|.~Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .Execute Replace:=wdReplaceAll
    End With
    Call Copy_cau("~.|.~", "")
    Application.ScreenUpdating = True
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n nh" & ChrW(7899) & " save file m" & ChrW(7899) & "i l" & ChrW(7841) & "i nh" & ChrW(233)
    Application.Assistant.DoAlert Title, msg, 0, 8, 0, 0, 0
End Sub
Sub Delete_cau_HighLight(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
    XoaHighlight.Show
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.zz^13"
        .MatchWildcards = False
    If Selection.Find.Execute = False Then Exit Sub
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .text = "Câu "
        .Replacement.text = "~.|.~Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)"
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)" & "  "
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "(~.|.~Câu [0-9]{1,4}*)(A.*)(B.*)(C.*)(D.*)(z.zz^13)"
        .Replacement.ClearFormatting
        .Replacement.text = ""
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "~.|.~"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.zz^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert Title, msg, 0, 1, 0, 0, 0
End Sub
Sub Highlight_cau_tuong_tu(ByVal control As Office.IRibbonControl)
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Application.ScreenUpdating = False

' Bien moi cau dan thanh 1 paragraph duy nhat, Phan tra loi la mot paragraph duy nhat
' Co gang lam cho phan tra loi o cac cau hoi deu khac nhau
'  (muc tieu: chi can word nhan dang cau dan giong nhau la duoc, khong nhan dang phan tra loi)

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "A."
        .Replacement.text = "A.A."
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
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^p"
        .Replacement.text = "^l"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^lA."
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])(^t)"
        .Replacement.text = "\1" & "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
Selection.HomeKey Unit:=wdStory

' Tim paragraph bi lap lai va to mau highlight (xanh)

Dim p1 As Paragraph
Dim p2 As Paragraph
Dim DupCount As Long

DupCount = 0

For Each p1 In ActiveDocument.Paragraphs
  If p1.Range.text <> vbCr Then
    
    For Each p2 In ActiveDocument.Paragraphs
      If p1.Range.text = p2.Range.text Then
        DupCount = DupCount + 1
        If p1.Range.text = p2.Range.text And DupCount > 1 Then
        p1.Range.Select
        Options.DefaultHighlightColorIndex = wdTurquoise
        Selection.Range.HighlightColorIndex = wdTurquoise
        End If
      End If
    Next p2
    
  End If
  
  'Reset Duplicate Counter
    DupCount = 0

Next p1

' Reset dinh dang cac cau hoi

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = CentimetersToPoints(-1.75)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = ":^p"
        .Replacement.text = ":^t"
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
Selection.HomeKey Unit:=wdStory
Application.ScreenUpdating = True

Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t. C" & ChrW(225) & "c c" & ChrW(226) & "u h" & ChrW(7887) & "i c" & ChrW(243) & " kh" & ChrW(7843) & " n" & ChrW(259) & "ng l" & ChrW(7863) & "p l" & ChrW(7841) & "I" & vbCrLf & "ho" & ChrW(7863) & "c t" & ChrW(432) & "" & ChrW(417) & "ng t" & ChrW(7921) & " c" & ChrW(226) & "u kh" & ChrW(225) & "c " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c t" & ChrW(244) & " n" & ChrW(7873) & "n m" & ChrW(224) & "u v" & ChrW(224) & "ng."
    Application.Assistant.DoAlert Title, msg, 0, 3, 0, 0, 0
End Sub
Sub Lay_cau_chua_key(ByVal control As Office.IRibbonControl)
' Chep cac cau hoi theo mot CUM TU cho truoc ra 1 file moi
' (Cac cau hoi do van con duoc danh dau o file cu)
    keysearch.Show
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    Options.DefaultHighlightColorIndex = wdNoHighlight
    Selection.Range.HighlightColorIndex = wdNoHighlight
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "~.|.~"
        .MatchWildcards = False
    If Selection.Find.Execute = False Then
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
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
            .text = "z.zz"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        Exit Sub
    End If
    End With
    Call Copy_cau("~.|.~", "CH" & ChrW(7912) & "A T" & ChrW(7914) & " KHO" & ChrW(193) & " C" & ChrW(7846) & "N T" & ChrW(204) & "M")
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert Title, msg, 0, 8, 0, 0, 0
End Sub
Sub CopyRed(ByVal control As Office.IRibbonControl)
Selection.HomeKey Unit:=wdStory
Dim ThisDoc As Document
Dim ThatDoc As Document
Selection.Find.ClearFormatting
Selection.Find.Font.Color = wdColorRed
Set ThisDoc = ActiveDocument
With Selection.Find
.text = ""
.Replacement.text = ""
If Selection.Find.Execute = True Then
Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
Else
Exit Sub
End If
ThisDoc.Activate
Selection.Copy
Do
Selection.Copy
ThatDoc.Activate
Selection.PasteAndFormat (wdFormatOriginalFormatting)
ThatDoc.Range
ThisDoc.Activate
Selection.Copy
Loop While Selection.Find.Execute(Forward:=True) = True
End With
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert Title, msg, 0, 8, 0, 0, 0
End Sub
Sub CopyBlue(ByVal control As Office.IRibbonControl)
Selection.HomeKey Unit:=wdStory
Dim ThisDoc As Document
Dim ThatDoc As Document
Selection.Find.ClearFormatting
Selection.Find.Font.Color = wdColorBlue
Set ThisDoc = ActiveDocument
With Selection.Find
.text = ""
.Replacement.text = ""
If Selection.Find.Execute = True Then
Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
Else
Exit Sub
End If
ThisDoc.Activate
Selection.Copy
Do
Selection.Copy
ThatDoc.Activate
Selection.PasteAndFormat (wdFormatOriginalFormatting)
ThatDoc.Range
ThisDoc.Activate
Selection.Copy
Loop While Selection.Find.Execute(Forward:=True) = True
End With
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert Title, msg, 0, 8, 0, 0, 0
End Sub
Sub Gioi_thieu(ByVal control As Office.IRibbonControl)
    ganid6.Show
End Sub
Sub GanID(ByVal control As Office.IRibbonControl)
    GanIDfrm.Show vbModeless
End Sub

Private Sub Copy_cau(ByVal Key As String, ByVal Titles As String)
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    If Key <> "~.|.~" Then
    With Selection.Find
        .ClearFormatting
        .text = "(Câu [0-9]{1,4})(*)" & Key
        .Replacement.text = Key & "\1\2" & Key
        .Forward = False
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
    If Selection.Find.Execute = False Then Exit Sub
        .Execute Replace:=wdReplaceAll
    End With
    End If
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)"
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .text = "([A-D].)" & "  "
        .Replacement.text = "\1" & " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Set ThisDoc = ActiveDocument
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = Key & "(Câu [0-9]{1,4}*)(A.*)(B.*)(C.*)(D.*)(z.zz^13)"
        .Replacement.ClearFormatting
        .Replacement.text = "\1\2\3\4\5\6"
        .MatchWildcards = True
    If Selection.Find.Execute = True Then
    Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    Else
    Exit Sub
    End If
    ThisDoc.Activate
    Selection.Copy
    Do
    Selection.Copy
    ThatDoc.Activate
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    ThisDoc.Activate
    Selection.Copy
    Loop While Selection.Find.Execute(Forward:=True) = True
    End With
    ThisDoc.Activate
    With Selection.Find
        .text = Key & "Câu"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.zz^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    ThatDoc.Activate
    With Selection.Find
        .text = Key
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.zz^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    Selection.WholeStory
    Options.DefaultHighlightColorIndex = wdNoHighlight
    Selection.Range.HighlightColorIndex = wdNoHighlight
    If Titles <> "" Then
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Font.Name = "Times New Roman"
    Selection.TypeText text:="C" & ChrW(193) & "C C" & ChrW(194) & "U H" & ChrW(7886) & "I " & Titles
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Size = 16
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    Else
    End If
    Selection.HomeKey Unit:=wdStory
End Sub
Private Sub To_Red_de_bai()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
With Selection.Find
.ClearFormatting
.text = " D."
.Replacement.text = "ZZ. "
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Format = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.ClearFormatting
.text = "^9D."
.Replacement.text = "^9ZZ. "
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Format = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.ClearFormatting
.text = "D."
.Replacement.text = "XX."
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Format = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With

With Selection.Find
.ClearFormatting
.text = ".  "
.Replacement.text = ". "
.Forward = True
.Wrap = wdFindContinue
.Format = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.ClearFormatting
.text = "(Câu [0-9]{1,4})(*)(ZZ.*)(^13)"
.Replacement.ClearFormatting
.Replacement.text = "\1\2\3\4"
.Replacement.Font.Color = wdColorRed
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchCase = True
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.ClearFormatting
.text = "(Câu [0-9]{1,4})(*)(ZZ.*)(^13)"
.Replacement.ClearFormatting
.Replacement.text = "\1\2\3\4"
.Replacement.Font.Color = wdColorRed
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchCase = True
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.ClearFormatting
.text = "ZZ."
.Replacement.text = "D."
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Format = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.ClearFormatting
.text = "XX."
.Replacement.text = "D."
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.Format = False
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With

End Sub

Sub Tach_cau_theo_4_muc_do()
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
' Kiem tra thong bao yeu cau ve to mau cho moi loai cau hoi
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .MatchWildcards = False
        .text = "LÝU " & ChrW(221) & " V" & ChrW(7872) & " CÁC K" _
         & ChrW(221) & " HI" & ChrW(7878) & "U NH" & ChrW(7852) & "N D" & ChrW( _
        7840) & "NG"
    If Selection.Find.Execute = False Then
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Select
    Selection.SplitTable
    Else
    End If
' Hien thi thong bao yeu cau to mau cho moi loai cau hoi
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Bold = False
    Selection.Font.Size = 13
    Selection.Font.Color = wdColorAutomatic
    Selection.TypeText text:="LÝU " & ChrW(221) & " V" & ChrW(7872) & " CÁC K" _
         & ChrW(221) & " HI" & ChrW(7878) & "U NH" & ChrW(7852) & "N D" & ChrW( _
        7840) & "NG"
    Selection.TypeParagraph
    Selection.TypeText text:="1. C" & ChrW(7847) & "n g" & ChrW(245) & _
        " thêm k" & ChrW(253) & " hi" & ChrW(7879) & "u nh" & ChrW(7853) & "n d" _
        & ChrW(7841) & "ng vào phía sau s" & ChrW(7889) & " th" & ChrW(7913) & _
        " t" & ChrW(7921) & " c" & ChrW(7911) & "a "
    Selection.TypeText text:="m" & ChrW(7895) & "i câu h" & ChrW(7887) & "i."
    Selection.TypeParagraph
    Selection.TypeText text:="2. Các k" & ChrW(253) & " hi" & ChrW(7879) & _
        "u nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng bao g" & ChrW(7891) & _
        "m : [NB] , [TH] , [VDT] , [VDC]."
    Selection.TypeParagraph
    Selection.HomeKey Unit:=wdStory
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Bold = wdToggle
    Selection.Font.Size = 14
    Selection.Font.Color = wdColorBlue
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.Range.HighlightColorIndex = wdYellow
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.HomeKey Unit:=wdStory
    Exit Sub
    Else
    End If
    End With
' Xoa thong bao to mau cho cac cau hoi
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1
' Bat dau thao tac tach cau hoi
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "["
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "]"
        .Replacement.text = "~"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Call Copy_cau("#NB~", "NH" & ChrW(7852) & "N BI" & ChrW(7870) & "T")
    Call Copy_cau("#TH~", "THÔNG HI" & ChrW(7874) & "U")
    Call Copy_cau("#VDT~", "V" & ChrW(7852) & "N D" & ChrW(7908) & "NG TH" & ChrW(7844) & "P")
    Call Copy_cau("#VDC~", "V" & ChrW(7852) & "N D" & ChrW(7908) & "NG CAO")
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "#"
        .Replacement.text = "["
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "~"
        .Replacement.text = "]"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Call To_mau_4_muc_do
Application.ScreenUpdating = True
End Sub
Sub To_mau_4_muc_do()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = wdColorRed
    End With
    With Selection.Find
        .text = "[VDC]"
        .Replacement.text = "[VDC]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 49407
    End With
    With Selection.Find
        .text = "[VDT]"
        .Replacement.text = "[VDT]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 5287936
    End With
    With Selection.Find
        .text = "[TH]"
        .Replacement.text = "[TH]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = wdColorBlue
    End With
    With Selection.Find
        .text = "[NB]"
        .Replacement.text = "[NB]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Xoa_ky_hieu_4_muc_do()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "[VDC]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "[VDT]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "[TH]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "[NB]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

