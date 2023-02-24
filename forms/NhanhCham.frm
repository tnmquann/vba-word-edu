VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NhanhCham 
   Caption         =   " "
   ClientHeight    =   1545
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6720
   OleObjectBlob   =   "NhanhCham.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NhanhCham"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
' Gach chan cau hoi
  NhanhCham.Hide
  ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineDouble
    Selection.Find.Replacement.Font.Color = -603937025
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
' Gach chan dap an
  sodapan = 0
  If GachDA.CheckBox1 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Font.Underline = wdUnderlineSingle
        With Selection.Find
            .text = "([A-D])"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .Format = True
            .MatchWildcards = True
        Do While .Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=2
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.text = "." Then
                sodapan = sodapan + 1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.Font.Underline = wdUnderlineDouble
                Selection.Font.Color = -603937025
                Options.DefaultHighlightColorIndex = wdNoHighlight
                Selection.Range.HighlightColorIndex = wdNoHighlight
            Else
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
        Loop
        End With
  End If
  If GachDA.CheckBox2 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Font.Color = wdColorRed
        With Selection.Find
            .text = "([A-D])"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .Format = True
            .MatchWildcards = True
        Do While .Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=2
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.text = "." Then
                sodapan = sodapan + 1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.Font.Underline = wdUnderlineDouble
                Selection.Font.Color = -603937025
                Options.DefaultHighlightColorIndex = wdNoHighlight
                Selection.Range.HighlightColorIndex = wdNoHighlight
            Else
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
        Loop
        End With
  End If
  If GachDA.CheckBox3 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Highlight = True
        With Selection.Find
            .text = "([A-D])"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .Format = True
            .MatchWildcards = True
        Do While .Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=2
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.text = "." Then
                sodapan = sodapan + 1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.Font.Underline = wdUnderlineDouble
                Selection.Font.Color = -603937025
                Options.DefaultHighlightColorIndex = wdNoHighlight
                Selection.Range.HighlightColorIndex = wdNoHighlight
            Else
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
        Loop
        End With
  End If
  Selection.HomeKey Unit:=wdStory
  ' Kiem tra xem co danh dau dap an hay khong
    If sodapan = 0 Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg2 = "Ch" & ChrW(432) & "a c" & ChrW(243) & " c" & ChrW(226) & "u n" & ChrW(224) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c " & ChrW(273) & "" & ChrW(225) & "nh d" & ChrW(7845) & "u " & ChrW(273) & "" & ChrW(225) & "p " & ChrW(225) & "n theo c" & ChrW(225) & "ch b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " ch" & ChrW(7885) & "n"
        Application.Assistant.DoAlert Title, msg2, 0, 1, 0, 0, 0
        GachDA.CheckBox1 = False
        GachDA.CheckBox2 = False
        GachDA.CheckBox3 = False
        Exit Sub
    End If
  GachDA.CheckBox1 = False
  GachDA.CheckBox2 = False
  GachDA.CheckBox3 = False
  Call Bang_Dap_An(GachDA.CheckBox1, GachDA.CheckBox2, GachDA.CheckBox3)
  Application.ScreenUpdating = True
End Sub
Private Sub CommandButton2_Click()
Application.ScreenUpdating = False
' Gach chan cau hoi
  NhanhCham.Hide
  ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineDouble
    Selection.Find.Replacement.Font.Color = -603937025
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
' Gach chan dap an
  If GachDA.CheckBox1 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Font.Underline = wdUnderlineSingle
        Selection.Find.Replacement.Highlight = wdNoHighlight
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find.Replacement.Font
            .Underline = wdUnderlineDouble
            .Color = -603937025
        End With
        With Selection.Find
            .text = "([A-D])"
            .Replacement.text = "\1"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
  End If
  If GachDA.CheckBox2 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Font.Color = wdColorRed
        Selection.Find.Replacement.Highlight = wdNoHighlight
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find.Replacement.Font
            .Underline = wdUnderlineDouble
            .Color = -603937025
        End With
        With Selection.Find
            .text = "([A-D])"
            .Replacement.text = "\1"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
  End If
  If GachDA.CheckBox3 = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Highlight = True
        With Selection.Find
            .text = "([A-D])"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .Format = True
            .MatchWildcards = True
        Do While .Execute
            Selection.MoveRight Unit:=wdCharacter, Count:=2
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.text = "." Then
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.Font.Underline = wdUnderlineDouble
                Selection.Font.Color = -603937025
                Options.DefaultHighlightColorIndex = wdNoHighlight
                Selection.Range.HighlightColorIndex = wdNoHighlight
            Else
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
        Loop
        End With
  End If
  Selection.HomeKey Unit:=wdStory
' Kiem tra xem co danh dau dap an hay khong
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDouble
    Selection.Find.Font.Color = -603937025
    With Selection.Find
        .text = "([A-D])"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
    If Selection.Find.Execute = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg2 = "Ch" & ChrW(432) & "a c" & ChrW(243) & " c" & ChrW(226) & "u n" & ChrW(224) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c " & ChrW(273) & "" & ChrW(225) & "nh d" & ChrW(7845) & "u " & ChrW(273) & "" & ChrW(225) & "p " & ChrW(225) & "n theo c" & ChrW(225) & "ch b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " ch" & ChrW(7885) & "n"
        Application.Assistant.DoAlert Title, msg2, 0, 1, 0, 0, 0
        GachDA.CheckBox1 = False
        GachDA.CheckBox2 = False
        GachDA.CheckBox3 = False
        Exit Sub
    End If
    End With
  GachDA.CheckBox1 = False
  GachDA.CheckBox2 = False
  GachDA.CheckBox3 = False
  Call Bang_Dap_An(GachDA.CheckBox1, GachDA.CheckBox2, GachDA.CheckBox3)
Application.ScreenUpdating = True
End Sub
Private Sub Bang_Dap_An(ByVal CheckBox1 As Boolean, CheckBox2 As Boolean, CheckBox3 As Boolean)
Application.ScreenUpdating = False
' Chep_dap_an_ra_ngoai
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Set ThisDoc = ActiveDocument
    Dim StrTxt As String, Doc As Document
    With ThisDoc.Range
    With .Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = ""
        .Replacement.text = ""
        .Font.Underline = wdUnderlineDouble
        .Font.Color = -603937025
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .Execute
    End With
    Do While .Find.Found
        StrTxt = StrTxt & .text
        If Right(.text, 1) <> vbCr Then StrTxt = StrTxt & vbCr
        .Collapse wdCollapseEnd
        .Find.Execute
    Loop
    End With
    Set ThatDoc = Documents.add
    ThatDoc.Range.text = StrTxt
    ThatDoc.Activate
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
            .ClearFormatting
            .text = "([.,:_;^9^13^32-])"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "Câu"
        .Replacement.text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "," & "(1)"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceOne
    End With
    With Selection.Find
        .text = "([A-D])"
        .Replacement.text = "." & "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = ".."
        .Replacement.text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.WholeStory
    Selection.ConvertToTable Separator:=wdSeparateByCommas, NumColumns:=10, _
        NumRows:=4, AutoFitBehavior:=wdAutoFitFixed
    With Selection.Tables(1)
        .Style = "Table Grid"
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
    End With
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
    Selection.Copy
    ThisDoc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    Selection.Font.Italic = False
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    Selection.Font.Color = 49407
    Selection.TypeText text:="B" & ChrW(7842) & "NG " & ChrW(272) & "" & ChrW(193) & "P " & ChrW(193) & "N"
    Selection.TypeParagraph
    Selection.EndKey Unit:=wdStory
    Selection.PasteAndFormat (wdUseDestinationStylesRecovery)
    ThatDoc.Close (No)
    ThisDoc.Activate
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDouble
    Selection.Find.Font.Color = -603937025
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    Selection.Find.Replacement.Font.Color = wdColorGreen
    With Selection.Find
        .text = "(Câu [0-9]{1,4}[.:])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDouble
    Selection.Find.Font.Color = -603937025
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.Font.Color = wdColorRed
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .text = "([A-D])"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
Application.ScreenUpdating = True
    msg2 = "B" & ChrW(7843) & "ng " & ChrW(273) & "" & ChrW(225) & "p " & ChrW(225) & "n " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c t" & ChrW(7841) & "o xong"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg2, 0, 4, 0, 0, 0
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label1_Click()

End Sub
