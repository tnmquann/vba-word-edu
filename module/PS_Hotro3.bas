Attribute VB_Name = "PS_Hotro3"
Option Explicit
Sub Ghi_chu_thich_new(ByVal control As Office.IRibbonControl)
    On Error Resume Next
    Chuthich.OptionButton1 = True
    Chuthich.Show
End Sub
Sub Them_ID_4_Tham_so(ByVal control As Office.IRibbonControl)
    Them_ID.Show
End Sub
Sub huong_dan_nhap_lieu_new(ByVal control As Office.IRibbonControl)
    Huongdan.Show
End Sub
Sub Gach_DA_bang(ByVal control As Office.IRibbonControl)
On Error Resume Next
Application.Visible = False
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Dim CountImage As Integer, C As Integer, O2 As String, i As Integer, j As Integer, msg As String, socot As Integer, sohang As Integer
If ActiveDocument.Tables.Count = 0 Then Exit Sub
ActiveDocument.Range.ListFormat.ConvertNumbersToText
Application.ScreenUpdating = False
' Code giup han che nhan dang nham phuong an
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
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
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
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
    With Selection.Find
        .ClearFormatting
        .text = "(A.*)(B.*)(C.*)(D.*)"
        .Replacement.text = "#\1#\2#\3#\4"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
' Chep bang dap an sang file tam và Text hoa no
    Set ThisDoc = ActiveDocument
        ActiveDocument.Tables(ActiveDocument.Tables.Count).Select
        Selection.Copy
    Set ThatDoc = Documents.add
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        socot = ActiveDocument.Tables(1).Columns.Count
        sohang = ActiveDocument.Tables(1).Rows.Count
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .text = "Câu"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .text = "([.,:_;^32-])"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        C = 0
        With Selection.Find
            .ClearFormatting
            .text = "([0-9]{1,4})"
            .Replacement.text = "#" & "\1" & "#"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        Do While Selection.Find.Execute = True
            C = C + 1
        Loop
        If socot > 1 Then
        ActiveDocument.Tables(1).Rows(1).Cells(2).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
        O2 = Selection.text
        Else
        O2 = ""
        End If
        If sohang > 1 And O2 = "#2#" Then
        ActiveDocument.Tables(1).Select
        For i = 1 To socot
            Selection.Tables(1).Columns(i).Select
            Selection.Cells.Merge
        Next i
        End If
        ActiveDocument.Tables(1).Select
            Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=True
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .text = "([^13^32^9])"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
'Ghi dap an vao mang
        Selection.HomeKey Unit:=wdStory
        Dim Arr(1 To 999)
        For j = 1 To C
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "#" & j & "#"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Arr(j) = Selection.text
        Next j
    ThatDoc.Close (No)
    ThisDoc.Activate
        Selection.HomeKey Unit:=wdStory
        For i = 1 To C
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "(Câu )" & i & "([.:])"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            If Selection.Find.Execute = True Then
                With Selection.Find
                    .text = "#" & Arr(i) & "."
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
            End If
            End With
            Selection.Find.Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Font.Color = wdColorRed
            Selection.Font.Bold = True
            Selection.Font.Underline = wdUnderlineSingle
        Next i
        With Selection.Find
            .ClearFormatting
            .text = "#"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        Selection.Find.ClearFormatting
        Selection.Find.Font.Underline = wdUnderlineSingle
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.Underline = wdUnderlineNone
        With Selection.Find
            .text = "([^9^32])"
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
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
    Dim oData   As New DataObject 'object to use the clipboard
    oData.SetText text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
ActiveDocument.UndoClear
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
    Application.Visible = True
End Sub
Sub Gach_DA_highlight(ByVal control As Office.IRibbonControl)
    Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(Ch?n?)([A-D])"
        .Replacement.text = "Ch" & ChrW(7885) & "n " & "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Ch?n)([A-D])"
        .Replacement.text = "Ch" & ChrW(7885) & "n " & "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Ch?n?)([A-D])."
        .Replacement.text = "Ch" & ChrW(7885) & "n " & "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Underline = wdUnderlineSingle
        .Color = wdColorRed
    End With
    With Selection.Find
        .text = "([A-D].)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    'Selection.WholeStory
    'Options.DefaultHighlightColorIndex = wdNoHighlight
    'Selection.Range.HighlightColorIndex = wdNoHighlight
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Bold = True
        .Font.Underline = wdUnderlineSingle
        .Font.Color = wdColorRed
        .text = "."
        .Replacement.text = "."
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.Color = wdColorBlue
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Options.DefaultHighlightColorIndex = wdBrightGreen
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    With Selection.Find
        .text = "(Ch?n [A-D])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True
End Sub
Sub Ghep_File_In_Folder_new(ByVal control As Office.IRibbonControl)
    On Error Resume Next
    Application.Visible = False
    Dim PathFolder, OldFileFolder, NewFileName, NewDoc As Document, i As Integer
    Dim FileNguon1, FileNguon2, FileName, msg
    PathFolder = ActiveDocument.path & "\"
    For i = 1 To Len(ActiveDocument.path)
        If Right(ActiveDocument.path, 1) = ":" Then
            NewFileName = ActiveDocument.Name
            Exit For
        Else
            If Mid(ActiveDocument.path, Len(ActiveDocument.path) - i, 1) = "\" Then
                NewFileName = Right(ActiveDocument.path, i) & ".docx"
                Exit For
            End If
        End If
    Next i
    OldFileFolder = ActiveDocument.path & "\Old Files\"
    If DirExists(ActiveDocument.path & "\Old Files\") = False Then
            MkDir (ActiveDocument.path & "\Old Files\")
    End If
    ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set NewDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    For i = 1 To 9999
        Selection.Find.ClearFormatting
        With Selection.Find
            .text = "^13^13"
            .Replacement.text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
        Selection.EndKey Unit:=wdStory
        FileNguon1 = PathFolder & "*.docx"
        FileNguon2 = PathFolder & "*.doc"
        If Dir(FileNguon1) <> "" Or Dir(FileNguon2) <> "" Then
            If Dir(FileNguon1) <> "" Then
                FileName = Dir(FileNguon1)
            Else
                FileName = Dir(FileNguon2)
            End If
            Selection.InsertFile (PathFolder & FileName)
            Name (PathFolder & FileName) As (OldFileFolder & FileName)
        Else
            Exit For
        End If
    Next i
    ' Save file voi ten ban dau
    ActiveDocument.SaveAs PathFolder & NewFileName
    If i = 2 Then
        ' Thong bao Thu muc hien hanh chi chua 1 file word
        msg = "File b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " m" & ChrW(7903) & " l" & ChrW(224) & " file word duy nh" & ChrW(7845) & "t trong th" & ChrW(432) & " m" & ChrW(7909) & "c ch" & ChrW(7913) & "a n" & ChrW(243)
    Else
        ' Thong bao hoan thanh viec ghep file
        msg = "C" & ChrW(225) & "c file c" & ChrW(7911) & "a b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c gh" & ChrW(233) & "p th" & ChrW(224) & "nh c" & ChrW(244) & "ng." & vbCrLf & "B" & ChrW(7841) & "n n" & ChrW(234) & "n s" & ChrW(7855) & "p l" & ChrW(7841) & "i th" & ChrW(7913) & " t" & ChrW(7921) & " cho c" & ChrW(225) & "c c" & ChrW(226) & "u h" & ChrW(7887) & "i nh" & ChrW(233)
    End If
    Application.Assistant.DoAlert "", msg, 0, 4, 0, 0, 0
    Selection.HomeKey Unit:=wdStory
    Application.Visible = True
End Sub
Private Sub Chay_TT_cau_Auto(ByVal KeyCau As Boolean, KeyBai As Boolean, KeyNumber As Boolean, DemCau As Boolean)
    Dim danhsach, socau, msg
    On Error Resume Next
    Application.ScreenUpdating = False
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
Public Sub Do_cau_trung(ByVal control As Office.IRibbonControl)
    Dim p1 As Paragraph, p2 As Paragraph, OldName As String, Title As String, msg As String
    Dim DupCount As Long, STTtuongtu As Long, OldDoc As Document, ThisDoc As Document, ThatDoc As Document

' Bien moi cau dan thanh 1 paragraph duy nhat, Phan tra loi la mot paragraph duy nhat
' Co gang lam cho phan tra loi o cac cau hoi deu khac nhau
'  (muc tieu: chi can word nhan dang cau dan giong nhau la duoc, khong nhan dang phan tra loi)
    Application.ScreenUpdating = False
    Application.Visible = False
    Wait.Show (0)
    Wait.Repaint
    On Error Resume Next
    Selection.WholeStory
        Selection.Copy
        Selection.HomeKey Unit:=wdStory
    Set OldDoc = ActiveDocument
    Set ThisDoc = Documents.add(DocumentType:=wdNewBlankDocument)
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Selection.WholeStory
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        ActiveDocument.Range.ListFormat.ConvertNumbersToText
    ThisDoc.Activate
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "(A.*[B-C].*D.*)(Câu[^32^s][0-9]{1,4}[.:])"
        .Replacement.text = "z.zz^p\2"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .ClearFormatting
        .text = "(A.*[B-C].*D.*^13)"
        .Replacement.text = "z.zz"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .ClearFormatting
        .text = "^9"
        .Replacement.text = ""
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .ClearFormatting
        .text = "(\[*\])"
        .Replacement.text = ""
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "  "
        .Replacement.text = " "
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll, Forward:=True
    Loop
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "^13"
        .Replacement.text = "^l"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=False
    End With
    With Selection.Find
        .ClearFormatting
        .text = "z.zz^l"
        .Replacement.text = "z.zz^p"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    With Selection.Find
        .ClearFormatting
        .text = "(Câu[^32^s][0-9]{1,4}[.:])"
        .Replacement.text = "\1^p"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll, Forward:=True
    End With
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "(^32)([^32,.;:?])"
        .Forward = True
        .Replacement.text = "\2"
        .Replacement.ClearFormatting
        .Wrap = wdFindContinue
        .MatchWildcards = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.HomeKey Unit:=wdStory

' Tim paragraph bi lap lai va to mau highlight (xanh)

    Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    ThisDoc.Activate
    Selection.HomeKey Unit:=wdStory
        STTtuongtu = 0
    For Each p1 In ActiveDocument.Paragraphs
        DupCount = 0
        If p1.Range.text <> vbCr Then
            For Each p2 In ActiveDocument.Paragraphs
                If p1.Range.text = p2.Range.text Then
                    DupCount = DupCount + 1
                    If DupCount > 1 Then
                        p2.Range.Select
                        Selection.MoveUp Unit:=wdLine, Count:=1
                        Selection.HomeKey Unit:=wdLine
                        With Selection.Find
                            .ClearFormatting
                            .text = "(Câu*z.zz^13)"
                            .Replacement.text = ""
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        End With
                            Selection.Find.Execute
                            Selection.Cut
                        ThatDoc.Activate
                            Selection.PasteAndFormat (wdFormatOriginalFormatting)
                        ThisDoc.Activate
                    End If
                End If
            Next p2
            If DupCount > 1 Then
                STTtuongtu = STTtuongtu + 1
                p1.Range.Select
                Selection.MoveUp Unit:=wdLine, Count:=1
                Selection.HomeKey Unit:=wdLine
                With Selection.Find
                    .ClearFormatting
                    .text = "(Câu*z.zz^13)"
                    .Replacement.text = ""
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                End With
                    Selection.Find.Execute
                    Selection.Cut
                ThatDoc.Activate
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    Selection.TypeText text:="-----------------------------"
                    Selection.TypeParagraph
                ThisDoc.Activate
            End If
    End If
    Next p1
    ThisDoc.Close SaveChanges:=wdDoNotSaveChanges
    
    If STTtuongtu = 0 Then
        ThatDoc.Close SaveChanges:=wdDoNotSaveChanges
        Wait.Hide
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(250) & "c m" & ChrW(7915) & "ng b" & ChrW(7841) & "n! C" & ChrW(243) & " l" & ChrW(7869) & " file c" & ChrW(7911) & "a b" & ChrW(7841) & "n kh" & ChrW(244) & "ng c" & ChrW(243) & " c" & ChrW(226) & "u tr" & ChrW(249) & "ng."
        Application.Assistant.DoAlert Title, msg, 0, 3, 0, 0, 0
        OldDoc.Activate
        Application.Visible = True
    Else
        OldDoc.Activate
            ActiveDocument.Range.ListFormat.ConvertNumbersToText
        ThatDoc.Activate
        ' Reset dinh dang cac cau hoi
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "^l"
            .Replacement.text = "^p"
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=False
        End With
        With Selection.Find
            .text = "z.zz^13"
            .Replacement.text = ""
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=False
        End With
        With Selection.Find
            .ClearFormatting
            .text = "(Câu[^32^s][0-9]{1,4}[.:])(^13)"
            .Replacement.text = "\1"
            .Replacement.ClearFormatting
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll, Forward:=True
        End With
        Selection.HomeKey Unit:=wdStory
        Application.ScreenUpdating = True
        Wait.Hide
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t. C" & ChrW(225) & "c c" & ChrW(226) & "u h" & ChrW(7887) & "i c" & ChrW(243) & " kh" & ChrW(7843) & " n" & ChrW(259) & "ng l" & ChrW(7863) & "p l" & ChrW(7841) & "I" & vbCrLf & "ho" & ChrW(7863) & "c t" & ChrW(432) & "" & ChrW(417) & "ng t" & ChrW(7921) & " c" & ChrW(226) & "u kh" & ChrW(225) & "c " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c t" & ChrW(244) & " n" & ChrW(7873) & "n m" & ChrW(224) & "u v" & ChrW(224) & "ng."
        Application.Assistant.DoAlert Title, msg, 0, 3, 0, 0, 0
        ThatDoc.Activate
        Application.Visible = True
    End If
    End
End Sub
Private Sub Chay_Sap_lai_thu_tu_cau_TN()
    Application.Visible = False
    Wait.Show (0)
    Wait.Repaint
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Dim i, m, j, b, n, k
    Dim kLop, kMon, kChuong, kBai, kDang, kMucdo
    Dim SoChuong, SoBai, SoDang, SoMucdo
    SoLop = 3
    SoMon = 3
    SoChuong = 6
    SoBai = 9
    SoDang = 30
    SoMucdo = 4
    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
' Them ky hieu nhan dang het cau
    Call End_cau
' Chuyen ma hieu nhan dang ra dau dong
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    With Selection.Find
        .ClearFormatting
        .text = "(Câu[^32^s][0-9]{1,4}*)(\[)([0-SoLop][DHL][1-SoChuong])(*)(-[1-SoMucdo])(\])"
        .Replacement.text = "##" & "\3\4\5" & "~" & "\1\2\3\4\5\6"
        .Replacement.ClearFormatting
        .Forward = False
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
    If Selection.Find.Execute = False Then
        Title1 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg1 = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a th" & ChrW(234) & "m k" & ChrW(253) & " hi" & ChrW(7879) & "u nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng c" & ChrW(226) & "u h" & ChrW(7887) & "i ho" & ChrW(7863) & "c k" & ChrW(253) & " hi" & ChrW(7879) & "u m" & ChrW(224) & vbCrLf & "b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " th" & ChrW(234) & "m ch" & ChrW(432) & "a " & ChrW(273) & "" & ChrW(250) & "ng theo h" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n c" & ChrW(7911) & "a ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh."
        Application.Assistant.DoAlert Title1, msg1, 0, 4, 0, 0, 0
        Huongdan.Show
        Exit Sub
    Else
        .Execute Replace:=wdReplaceAll
    End If
    End With
    Set ThisDoc = ActiveDocument
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "(##*Câu)"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
    Selection.Cut
    Set ThatDoc = Documents.add(DocumentType:=wdNewBlankDocument)
    Selection.PasteAndFormat (wdUseDestinationStylesRecovery)
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
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    ThisDoc.Activate
' Xoa ky hieu nhan dang thua
    If Sap_xep_cau.OptionButton1 = True Then
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .text = "(##)(*)(-[1-SoMucdo])(~)"
            .Replacement.text = "\1\3\4"
            .Replacement.ClearFormatting
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=True
        End With
    Else
        If Sap_xep_cau.OptionButton2 = True Then
            With Selection.Find
                .ClearFormatting
                .text = "(##)([0-SoLop])([DHL])([1-SoChuong])(*)(~)"
                .Replacement.text = "\1\2\3\4\6"
                .Replacement.ClearFormatting
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=True
            End With
        Else
            If Sap_xep_cau.OptionButton3 = True Then
                With Selection.Find
                    .ClearFormatting
                    .text = "(##)([0-SoLop])([DHL])([1-SoChuong])(.[1-SoBai])(*)(~)"
                    .Replacement.text = "\1\2\3\4\5\7"
                    .Replacement.ClearFormatting
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=True
                End With
            End If
        End If
    End If
' Chep tung nhom cau cung muc do cua moi chuong
    For i = 0 To (SoLop - 1)
        If Sap_xep_cau.OptionButton1 = True Then
            kLop = ""
            i = (SoLop - 1)
        Else
            kLop = i
        End If
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .text = "##" & kLop & "(*)(~)"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        If Selection.Find.Execute = True Then
        For m = 1 To SoMon
            If Sap_xep_cau.OptionButton1 = True Then
                kMon = ""
                m = SoMon
            Else
                If m = 1 Then
                        kMon = "D"
                Else
                    If m = 2 Then
                        kMon = "H"
                    Else
                        kMon = "L"
                    End If
                End If
            End If
            Selection.HomeKey Unit:=wdStory
            With Selection.Find
                .text = "##" & kLop & kMon & "(*)(~)"
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            If Selection.Find.Execute = True Then
            For j = 1 To SoChuong
                If Sap_xep_cau.OptionButton1 = True Then
                    kChuong = ""
                    j = SoChuong
                Else
                    kChuong = j
                End If
                Selection.HomeKey Unit:=wdStory
                With Selection.Find
                    .text = "##" & kLop & kMon & kChuong & "(*)(~)"
                    .Replacement.text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWildcards = True
                If Selection.Find.Execute = True Then
                For b = 1 To SoBai
                    If Sap_xep_cau.OptionButton1 = True Or Sap_xep_cau.OptionButton2 = True Then
                        kBai = ""
                        b = SoBai
                    Else
                        kBai = "." & b
                    End If
                    Selection.HomeKey Unit:=wdStory
                    With Selection.Find
                        .text = "##" & kLop & kMon & kChuong & kBai & "(*)(~)"
                        .Replacement.text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = True
                    If Selection.Find.Execute = True Then
                    For n = 1 To SoDang
                            kDang = ""
                            n = SoDang
                        Selection.HomeKey Unit:=wdStory
                        With Selection.Find
                            .text = "##" & kLop & kMon & kChuong & kBai & kDang & "(*)(~)"
                            .Replacement.text = ""
                            .Forward = True
                            .Wrap = wdFindContinue
                            .MatchWildcards = True
                        If Selection.Find.Execute = True Then
                        For k = 1 To SoMucdo
                            If Sap_xep_cau.OptionButton1 = False Then
                                kMucdo = ""
                                k = SoMucdo
                            Else
                                kMucdo = "-" & k
                            End If
                            Tukhoa = "##" & kLop & kMon & kChuong & kBai & kDang & kMucdo & "~" & "(Câu[^32^s][0-9]{1,4})(*)(z.zz)"
                            Selection.HomeKey Unit:=wdStory
                            With Selection.Find
                                .text = Tukhoa
                                .Replacement.text = ""
                                .Forward = True
                                .Wrap = wdFindContinue
                                .MatchWildcards = True
                            If Selection.Find.Execute = True Then
                                Selection.Find.ClearFormatting
                                With Selection.Find
                                    .text = Tukhoa
                                    .Replacement.ClearFormatting
                                    .Replacement.text = Tukhoa
                                    .MatchWildcards = True
                                Do
                                    Selection.Cut
                                    ThatDoc.Activate
                                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                                    Call ClearClipBoard
                                    ThisDoc.Activate
                                Loop While .Execute
                                End With
                                ActiveDocument.UndoClear
                            End If
                            End With
                        Next k
                        End If
                        End With
                    Next n
                    End If
                    End With
                Next b
                End If
                End With
            Next j
            End If
            End With
        Next m
        End If
        End With
    Next i
    ThisDoc.Close (No)
    ThatDoc.Activate
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=False
    End With
    With Selection.Find
        .text = "(##*~)"
        .Replacement.text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll, Forward:=True, MatchWildcards:=True
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    ActiveDocument.UndoClear
    Application.ScreenUpdating = True
    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7901) & "ng d" & ChrW(7851) & "n file"
    msg2 = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t. B" & ChrW(7841) & "n nh" & ChrW(7899) & " save file m" & ChrW(7899) & "i n" & ChrW(224) & "y l" & ChrW(7841) & "i nh" & ChrW(233) & "!"
    Application.Assistant.DoAlert title2, msg2, 0, 4, 0, 0, 0
    Wait.Hide
    Application.Visible = True
End Sub

Public Sub Sap_lai_thu_tu_cau_TN_BTp(ByVal control As Office.IRibbonControl)
    Sap_xep_cau.Show
    End
End Sub
Sub MT_Convert1_new(ByVal control As Office.IRibbonControl)
    Call MT_Convert
End Sub
Sub MT_Convert()
    On Error Resume Next
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    If Len(Selection.text) = 1 Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Exit Sub
    End If
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
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
    SendKeys "%{F4}"
    End
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    With Selection.Find
        .text = "( )([.,;:^32])"
        .Replacement.text = "\2"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=2
End Sub
Sub Gioi_thieu_TG(ByVal control As Office.IRibbonControl)
    GioiThieu.Show
End Sub

Public Function DirExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
DirExists = fs.folderexists(OrigFile)
End Function
Private Sub ClearClipBoard()
Dim oData   As New DataObject 'object to use the clipboard
    oData.SetText text:=Empty 'Clear
    oData.PutInClipboard 'take in the clipboard to empty it
End Sub
Function FileThere(FileName As String) As Boolean
    FileThere = (Dir(FileName) <> "")
End Function
