VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_QTable 
   Caption         =   "Advanced"
   ClientHeight    =   7695
   ClientLeft      =   135
   ClientTop       =   3060
   ClientWidth     =   3210
   OleObjectBlob   =   "S_QTable.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "S_QTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ABCDChange_Click()
Call ConvertABCD
End Sub

Private Sub Auto2Text_Click()
Call ConvertAuto2Text
End Sub


Private Sub Label14_Click()
    Call delTable
End Sub

Private Sub Label15_Click()
Call ConvertTestPro
End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label19_Click()
Call ConvertABCD
End Sub

Private Sub Label23_Click()
Call Chuan_hoa_2
End Sub

Private Sub Label24_Click()
    If S_QTable.OptionButton1 = False And S_QTable.OptionButton2 = False Then
    MsgBox "Chua chon dang cau hoi"
    Else
    Call QuestionTable
    End If
End Sub

Private Sub Label25_Click()
Call ConvertText2Auto2
End Sub

Private Sub Label27_Click()
    Call S_PageSetup
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.WholeStory
    Selection.Delete
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(18), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.TypeText text:="Lê Hoài S" & ChrW(417) & "n - THPT Ngu"
    Selection.TypeText text:="y" & ChrW(7877) & "n Hu" & ChrW(7879) & " - Hu" & ChrW(7871) & Chr(9)
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
    End With
    'Selection.MoveRight unit:=wdCell, Count:=1
    'Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.InsertSymbol Font:="Wingdings", CharacterNumber:=-4056, Unicode:=True
    Selection.TypeText text:=" 0914 114 008"
    Selection.WholeStory
    
    Selection.Font.Name = "CommercialScript"
    Selection.Font.Size = 13
    Selection.Font.ColorIndex = wdBlue

    
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.WholeStory
    Selection.Delete
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        .Columns(1).Width = 420
        .Columns(2).Width = 100
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.TypeText text:=S_QTable.TextBox3.text
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(-0.2)
    End With
    Selection.MoveRight Unit:=wdCell, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.TypeText text:="Trang "
    Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
        "PAGE  ", PreserveFormatting:=True
    Selection.TypeText text:="/"
    Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
        "NUMPAGES  ", PreserveFormatting:=True
    Selection.Tables(1).Select
    Selection.Font.Name = "Ariston"
    Selection.Font.Size = 9.5
    Selection.Font.ColorIndex = wdViolet
    
    Options.DefaultBorderColor = 6299648
    Options.DefaultBorderColor = 12611584
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Selection.Tables(1).Cell(1, 1).Select
    Selection.Font.Name = "AvantGarde-Demi"
    Selection.Font.Size = 10
    Selection.Font.ColorIndex = wdViolet
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Private Sub Label28_Click()
        On Error Resume Next
        Dim C As Integer
        Call RemoveMarks
        Application.ScreenUpdating = True
        ActiveDocument.Range.ListFormat.ConvertNumbersToText
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        C = 0
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "([^13^32^9])([AaBbCcDd])([.:\)\/])"
            .Replacement.text = "\1\2\3" & " "
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "^p Câu"
            .Replacement.text = "^pCâu"
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
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "^pCâu"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Do While Selection.Find.Execute = True
            Selection.Collapse Direction:=wdCollapseStart
            C = C + 1
            With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
            Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.EndKey Unit:=wdLine
        Loop
        Selection.EndKey Unit:=wdStory
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C + 1 & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
        Dim i As Integer
        Dim myRange As Range
        
    For i = 1 To C
        Set myRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
        Selection.Find.ClearFormatting
        myRange.Find.Execute FindText:="([^13^32^9])([Dd])([.:\)])", MatchWildcards:=True
        If myRange.Find.Found = True Then
        myRange.Select
        Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="s1"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.ColorIndex = wdBlue
        Selection.TypeText text:="H" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i"
        Selection.TypeParagraph
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            If xoaHDG Then
                Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
                myRange.Delete
            End If
        End If
    Next i
    If themHDG Then
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.ColorIndex = wdRed
        Selection.Find.Replacement.Font.Bold = True
        With Selection.Find
            .text = "H" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i:"
            .Replacement.text = "H" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i:"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Else
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
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
    End If
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
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    MsgBox "Done!"
End Sub

Private Sub Label29_Click()
    If ktBanQuyen = False Then Call S_SerialHDD
    If ktBanQuyen = False Then
        S_NoteRig.Show
        Ch3A5.Value = False
        Exit Sub
    End If
    S_QTable.Hide
        Dim L1, L2, L3, L4, d_a, i As Byte
        Dim C As Integer
        Dim lmax As String
        Dim Shape1, Shape2, Shape3, Shape4 As Byte
        Dim title2, msg As String
        Dim ktMsg As Byte
        On Error Resume Next
        Dim myRange As Range
        
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = ChrW(272) & ChrW(7875) & " th" & ChrW(7921) & _
        "c hi" & ChrW(7879) & "n " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ch" _
        & ChrW(7913) & "c n" & ChrW(259) & "ng này các ph" & ChrW(432) & ChrW(417 _
        ) & "ng án ph" & ChrW(7843) & "i " & ChrW(273) & ChrW(432) & ChrW(417) & _
        "c xu" & ChrW(7889) & "ng dòng. N" & _
         ChrW(7871) & "u d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(7911 _
        ) & "a b" & ChrW(7841) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(432 _
        ) & ChrW(7907) & "c chu" & ChrW(7849) & "n hóa hãy ch" & ChrW(7885) & "n ""Yes"", n" & ChrW(7871) & _
        "u d" & ChrW(7919) & " li" & ChrW(7879) & "u " & ChrW(273) & "ã chu" & _
        ChrW(7849) & "n hóa r" & ChrW(7891) & "i hãy ch" & ChrW(7885) & _
        "n ""No""."
        ktMsg = Application.Assistant.DoAlert(Title, msg, 3, 2, 0, 2, 0)
        If ktMsg = 6 Then
            Call ChuanDATA
        ElseIf ktMsg = 2 Then
            Exit Sub
        Else
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            
            With Selection.Find
            .text = "^p "
            .Replacement.text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            Do While .Execute
            .Execute Replace:=wdReplaceAll
            Loop
            End With
        End If
        
        Call RemoveMarks
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "Câu "
                .MatchWildcards = False
            End With
        Do While Selection.Find.Execute = True
            Selection.HomeKey Unit:=wdLine
            Selection.TypeParagraph
            Exit Do
        Loop
        For i = 1 To ActiveDocument.Tables.Count
            ActiveDocument.Tables(i).Select
            Selection.MoveDown Unit:=wdLine, Count:=1
            Selection.TypeParagraph
        Next i
        C = 1
        ActiveDocument.Range.ListFormat.ConvertNumbersToText
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "([^13^32^9])([AaBbCcDd])([.:\)\/])"
            .Replacement.text = "\1\2\3" & " "
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.WholeStory
        Selection.ParagraphFormat.TabStops.ClearAll
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "^pCâu "
                .MatchWildcards = False
            End With
            Do While Selection.Find.Execute = True
                Selection.Collapse Direction:=wdCollapseEnd
                With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & C & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
                End With
                C = C + 1
            Loop
        Selection.EndKey Unit:=wdStory
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & C & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            choiceA = "([^13])([Aa])([.:\)])(*)([^13])([Bb])([.:\)])"
            choiceB = "([^13])([Bb])([.:\)])(*)([^13])([Cc])([.:\)])"
            choiceC = "([Cc])([.:\)])(*)([^13])([Dd])([.:\)])"
            choiceD = "([Dd])([.:\)])"
        For i = 1 To C - 1
  
            Selection.Find.ClearFormatting
            
            'Danh dau phuong an A
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
       
            myRange.Find.Execute FindText:=choiceA, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "a"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
   
            L1 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
                Case 1
                L1 = L1 + Round(myRange.InlineShapes(1).Width / 5.8)
                Case 2
                L1 = L1 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
             With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
            'Danh dau phuong an B
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Find.Execute FindText:=choiceB, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "b"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
           
            L2 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
                Case 1
                L2 = L2 + Round(myRange.InlineShapes(1).Width / 6)
                Case 2
                L2 = L2 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
           'Danh dau phuong an C
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Find.Execute FindText:=choiceC, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=0
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "c"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
           
            L3 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
               Case 1
                L3 = L3 + Round(myRange.InlineShapes(1).Width / 6)
                Case 2
                L3 = L3 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            
        End If
            'Danh dau phuong an D
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Find.Execute FindText:=choiceD, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.Select
            Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "d"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            
            L4 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
                Case 1
                L4 = L4 + Round(myRange.InlineShapes(1).Width / 6)
                Case 2
                L4 = L4 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            
        End If
            lmax = L1
            If Val(lmax) < L2 Then lmax = L2
            If Val(lmax) < L3 Then lmax = L3
            If Val(lmax) < L4 Then lmax = L4
            If Val(lmax) < 10 Then lmax = "0" & lmax
            If Val(lmax) > 60 Then lmax = 60
        ''''''''''
        Dim chia1, chia2  As Byte
        Dim tab2, tab3, tab4 As Long
        If S_QTable.Ch3A4 Then
            chia1 = 24
            chia2 = 45
            tab1 = 0.5
            tab2 = 5
            tab3 = 9.5
            tab4 = 14
        ElseIf S_QTable.Ch3A5 Then
            chia1 = 16
            chia2 = 30
            tab1 = 0.5
            tab2 = 3.2
            tab3 = 5.9
            tab4 = 8.6
        Else
            chia1 = 13
            chia2 = 25
            tab1 = 0.2
            tab2 = 2.4
            tab3 = 4.6
            tab4 = 6.8
        End If
        'Selection.WholeStory
        'Selection.ParagraphFormat.TabStops.ClearAll
        If Val(lmax) < chia1 Then
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
            'Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0.5)
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab
            
            ActiveDocument.DefaultTabStop = CentimetersToPoints(tab1)
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab1) _
            , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab2), _
            Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab3), _
            Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab4), _
            Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
            
            
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeBackspace
            Selection.TypeText text:=vbTab
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeBackspace
            Selection.TypeText text:=vbTab
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeBackspace
            Selection.TypeText text:=vbTab
        ElseIf Val(lmax) < chia2 Then
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab

            ActiveDocument.DefaultTabStop = CentimetersToPoints(tab1)
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab1) _
            , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab3), _
            Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeBackspace
            Selection.TypeText text:=vbTab
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"

            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab
            ActiveDocument.DefaultTabStop = CentimetersToPoints(tab1)
            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab1) _
            , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

            Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(tab3), _
            Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeBackspace
            Selection.TypeText text:=vbTab
            Else
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
            Selection.MoveLeft Unit:=wdCharacter
            Selection.TypeText text:=vbTab
            
        End If
    Next i
            Selection.Find.ClearFormatting
            Selection.Find.Font.Underline = wdUnderlineNone
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdBlue
            Selection.Find.Replacement.Font.Bold = True
            With Selection.Find
            .text = "([^9])([Aa])([.:\)\/])"
            .Replacement.text = "\1" & "A."
            .MatchCase = True
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
            End With
            
            With Selection.Find
            .text = "([^9])([Bb])([.:\)\/])"
            .Replacement.text = "\1" & "B."
            .MatchCase = True
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
            .text = "([^9])([Cc])([.:\)\/])"
            .Replacement.text = "\1" & "C."
            .MatchCase = True
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
            .text = "([^9])([Dd])([.:\)\/])"
            .Replacement.text = "\1" & "D."
            .MatchCase = True
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
            End With
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdBlue
            Selection.Find.Replacement.Font.Bold = True
             With Selection.Find
                .text = "(Câu [0-9]{1,4})"
                .Replacement.text = "\1" & "."
                .Forward = True
                .Format = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
            .text = ".."
            .Replacement.text = "."
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
            .text = ".:"
            .Replacement.text = "."
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
            .text = ".."
            .Replacement.text = "."
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
            End With
            With Selection.Find
            .text = "  "
            .Replacement.text = " "
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            Do While .Execute
            .Execute Replace:=wdReplaceAll
            Loop
            End With
            With Selection.Find
            .text = "^p^p"
            .Replacement.text = "^p"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            Do While .Execute
            .Execute Replace:=wdReplaceAll
            Loop
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Font.Underline = wdUnderlineSingle
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.Underline = wdUnderlineNone
            Selection.Find.Replacement.Font.ColorIndex = wdBlue
            With Selection.Find
            .text = "^t"
            .Replacement.text = "^t"
            .MatchCase = True
            .Forward = True
            .Format = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
            End With
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Font.Underline = wdUnderlineSingle
            Selection.Find.Replacement.Font.ColorIndex = wdBlue
            Selection.Find.Replacement.Font.Bold = True
            With Selection.Find
                .text = "([ABCD])"
                .Replacement.text = "\1"
                .Forward = True
                .Format = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            Selection.WholeStory
            Selection.Font.Name = "Times New Roman"
            Selection.Font.Size = 12
            
            Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        If ActiveDocument.Tables.Count > 0 Then
        For i = 1 To ActiveDocument.Tables.Count
            ActiveDocument.Tables(i).Select
            With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(0)
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            End With
        Next i
        ActiveDocument.Tables(1).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    If S_QTable.Ch3A4 Then
           Call S_PageSetup
           Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    ElseIf S_QTable.Ch3A5 Then
        Selection.WholeStory
        With ActiveDocument.PageSetup
            .LineNumbering.Active = False
            .Orientation = wdOrientPortrait
            .TopMargin = CentimetersToPoints(0.6)
            .BottomMargin = CentimetersToPoints(0.8)
            .LeftMargin = CentimetersToPoints(1.5)
            .RightMargin = CentimetersToPoints(0.86)
            .Gutter = CentimetersToPoints(0)
            .HeaderDistance = CentimetersToPoints(0.8)
            .FooterDistance = CentimetersToPoints(0.7)
            .PageWidth = CentimetersToPoints(14.8)
            .PageHeight = CentimetersToPoints(21)
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
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Else
        Selection.WholeStory
        With Selection.PageSetup.TextColumns
            .SetCount NumColumns:=2
            .EvenlySpaced = True
            .LineBetween = True
            .Width = CentimetersToPoints(5.97)
            .Spacing = CentimetersToPoints(0.5)
        End With
        With ActiveDocument.PageSetup
            .LineNumbering.Active = False
            .Orientation = wdOrientPortrait
            .TopMargin = CentimetersToPoints(0.6)
            .BottomMargin = CentimetersToPoints(0.8)
            .LeftMargin = CentimetersToPoints(1)
            .RightMargin = CentimetersToPoints(0.86)
            .Gutter = CentimetersToPoints(0)
            .HeaderDistance = CentimetersToPoints(0.8)
            .FooterDistance = CentimetersToPoints(0.7)
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
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    End If
    Selection.HomeKey Unit:=wdStory
    MsgBox "Xong."
Exit Sub
S_Quit:
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "D" & ChrW(7919) & " li" & ChrW(7879) & "u b" & _
        ChrW(7883) & " l" & ChrW(7895) & "i."
        Application.Assistant.DoAlert Title, msg, 4, 0, 0, 0, 0
End Sub



Private Sub Text2Auto_Click()
Call ConvertText2Auto
End Sub

Private Sub Standar1_Click()
Call Chuan_hoa_1
End Sub
Private Sub UserForm_Initialize()
S_QTable.Ch3A4 = True
End Sub
