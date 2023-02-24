Attribute VB_Name = "S_ShareSub"
Sub CheckDrive()
If S_Drive = "" Then S_Drive = "D:\"
If DirExists("D:\") = False Then S_Drive = "C:\"
End Sub
Public Function FExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
FExists = fs.fileexists(OrigFile)
End Function 'Returns a boolean - True if the file exists
Public Function DirExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
DirExists = fs.folderexists(OrigFile)
End Function
Sub OpenTestDir1(ByVal control As Office.IRibbonControl)
Call CheckDrive
ChangeFileOpenDirectory (S_Drive & "S_Bank&Test\S_Test\Lop 10\")
If Dialogs(wdDialogFileOpen).Show = -1 Then
ActiveWindow.View = wdPrintView
End If
End Sub
Sub OpenTestDir2(ByVal control As Office.IRibbonControl)
Call CheckDrive
ChangeFileOpenDirectory (S_Drive & "S_Bank&Test\S_Test\Lop 11\")
If Dialogs(wdDialogFileOpen).Show = -1 Then
ActiveWindow.View = wdPrintView
End If
End Sub
Sub OpenTestDir3(ByVal control As Office.IRibbonControl)
Call CheckDrive
ChangeFileOpenDirectory (S_Drive & "S_Bank&Test\S_Test\Lop 12\")
If Dialogs(wdDialogFileOpen).Show = -1 Then
ActiveWindow.View = wdPrintView
End If
End Sub
Sub OpenTestDir4(ByVal control As Office.IRibbonControl)
Call CheckDrive
ChangeFileOpenDirectory (S_Drive & "S_Bank&Test\S_Test\Other\")
If Dialogs(wdDialogFileOpen).Show = -1 Then
ActiveWindow.View = wdPrintView
End If
End Sub
Public Function DeleteFolder1(ByVal path As String)
    On Error GoTo Q
    Dim wkbk As Document
    path = IIf(Right$(path, 1) = "\", path, path & "\")
    For Each wkbk In Application.Documents
        wkbk.Close
    Next wkbk
    Do While Dir(path & "*.*") <> ""
        Kill path & Dir(path & "*.*")
    Loop
    path = Left(path, Len(path) - 1)
    RmDir (path)
Exit Function
Q:
    path = Left(path, Len(path) - 1)
    RmDir path
End Function
Public Function DeleteFolder(ByVal path As String)
    On Error GoTo Q
    Dim docActive As Document
    CreateObject("Scripting.FileSystemObject").DeleteFolder path
    Exit Function
Q:
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Mã môn mu" & ChrW(7889) & "n xóa có ch" & ChrW( _
        7913) & "a file " & ChrW(273) & "ang m" & ChrW(7903) & "." & Chr(13) & "B" & ChrW(7841) _
         & "n " & ChrW(273) & "óng file " & ChrW(273) & "ó và th" & ChrW(7921) & "c hi" & ChrW(7879) & "n l" & ChrW( _
        7841) & "i"
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
End Function
Sub RemoveMarks()
    Dim bkm As Bookmark
    For Each bkm In ActiveDocument.Bookmarks
    bkm.Delete
    Next bkm
End Sub
Sub FontFormat()
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = True
        .Color = 13382400
    End With
End Sub
Sub FontFormat2()
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 11
        .Color = 13382400
    End With
End Sub
Sub FontFormat3()
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
        .Color = 13382400
        .Bold = True
    End With
End Sub
Sub S_ParagaphFormat()
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0.5)
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
    End With
End Sub

Sub MkDr(ByVal control As Office.IRibbonControl)
Call MadeDir
End Sub
Sub QTable(ByVal control As Office.IRibbonControl)
S_QTable.Show
End Sub
Sub MadeDir()
    Call CheckDrive
    Dim path As String
    Dim Tb, Title, msg As String
    Dim ktMsg As Byte
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Ch" & ChrW(432) & ChrW(417) & "ng trình s" & _
        ChrW(7869) & " kh" & ChrW(7903) & "i t" & ChrW(7841) & "o th" & ChrW(432) _
         & " m" & ChrW(7909) & "c g" & ChrW(7889) & "c S_Bank&Test và các th" & _
        ChrW(432) & " m" & ChrW(7909) & "c con trong " & ChrW(7893) & " " & Left(S_Drive, 1)
    ktMsg = Application.Assistant.DoAlert(Title, msg, 4, 2, 0, 0, 1)
    If ktMsg = 6 Then
        path = "C:\Program Files\"
        Options.SaveInterval = 10
        Application.DefaultSaveFormat = ""
        If DirExists(S_Drive & "S_Bank&Test\") = False Then MkDir (S_Drive & "S_Bank&Test\")
        If DirExists(S_Drive & "S_Bank&Test\S_Test\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Test\")
        If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 10\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 10\")
        If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 11\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 11\")
        If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 12\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 12\")
        If DirExists(S_Drive & "S_Bank&Test\S_Test\Other\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Test\Other\")
        
        If DirExists(S_Drive & "S_Bank&Test\S_Data\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Data\")
        If DirExists(S_Drive & "S_Bank&Test\S_Data\Lop 10\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Data\Lop 10\")
        If DirExists(S_Drive & "S_Bank&Test\S_Data\Lop 11\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Data\Lop 11\")
        If DirExists(S_Drive & "S_Bank&Test\S_Data\Lop 12\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Data\Lop 12\")
        If DirExists(S_Drive & "S_Bank&Test\S_Data\Other\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Data\Other\")
        
        If DirExists(S_Drive & "S_Bank&Test\S_Bank\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Bank\")
        If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 10\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 10\")
        If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 11\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 11\")
        If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 12\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 12\")
        
        If DirExists(S_Drive & "S_Bank&Test\S_Templates\") = False Then MkDir (S_Drive & "S_Bank&Test\S_Templates\")
        
        If DirExists("C:\Program Files (x86)\") Then path = "C:\Program Files (x86)\"
        
        FileCopy path & "S_Bank&Test\S_Templates\Mark_Printer.docx", _
        S_Drive & "S_Bank&Test\S_Templates\Mark_Printer.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Answer.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Footer_1.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Footer_1.docx"
        
         FileCopy path & "S_Bank&Test\S_Templates\default_Footer_2.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Footer_2.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Header_1.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Header_1.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Header_2.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Header_2.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Header_3.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Header_3.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Header_4.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Header_4.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\default_Header_5.docx", _
        S_Drive & "S_Bank&Test\S_Templates\default_Header_5.docx"
        
        
        FileCopy path & "S_Bank&Test\S_Templates\Phieu_cham_bai.docx", _
        S_Drive & "S_Bank&Test\S_Templates\Phieu_cham_bai.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\AnswerSheet_50.docx", _
        S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_50.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\AnswerSheet_120.docx", _
        S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_120.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\AnswerSheet_A5.docx", _
        S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_A5.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\AnswerSheet_NH.docx", _
        S_Drive & "S_Bank&Test\S_Templates\AnswerSheet_NH.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\PPCT10_DS.docx", _
        S_Drive & "S_Bank&Test\S_Templates\PPCT10_DS.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\PPCT11_DS.docx", _
        S_Drive & "S_Bank&Test\S_Templates\PPCT11_DS.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\PPCT12_DS.docx", _
        S_Drive & "S_Bank&Test\S_Templates\PPCT12_DS.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\PPCT10_HH.docx", _
        S_Drive & "S_Bank&Test\S_Templates\PPCT10_HH.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\PPCT11_HH.docx", _
        S_Drive & "S_Bank&Test\S_Templates\PPCT11_HH.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\PPCT12_HH.docx", _
        S_Drive & "S_Bank&Test\S_Templates\PPCT12_HH.docx"
        
        FileCopy path & "S_Bank&Test\S_Templates\Help.doc", _
        S_Drive & "S_Bank&Test\S_Templates\Help.doc"
        
        
        If S_inf.t1 <> "" Then
        Else
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = ChrW(272) & "ã kh" & ChrW(7903) & "i t" & ChrW( _
            7841) & "o th" & ChrW(432) & " m" & ChrW(7909) & _
            "c S_Bank&Test và các th" & ChrW(432) & " m" & ChrW(7909) & _
            "c con trong " & ChrW(7893) & " " & Left(S_Drive, 1) & "."
            Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        End If
        With ActiveDocument.Styles(wdStyleNormal).Font
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
        End With
    End If
End Sub
Sub Taonhom(ByVal sonhom As Byte, ByVal tuluan As Boolean)
    S_QG.Hide
    Dim sodong As Byte
    Selection.HomeKey Unit:=wdStory
    With Selection.Font
                .Name = "Times New Roman"
                .Size = 12
                .Bold = True
                .Color = 13382400
            End With
    If tuluan Then
    sodong = sonhom + 2
    Else
    sodong = sonhom + 1
    End If
    Selection.TypeText text:="[<Gr>]"
    Selection.TypeParagraph
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=sodong, NumColumns:=4, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=45, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=320, RulerStyle:=wdAdjustFirstColumn
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=55, RulerStyle:=wdAdjustFirstColumn
    Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=60, RulerStyle:=wdAdjustFirstColumn
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText text:="Nhóm"
    Selection.Shading.BackgroundPatternColor = -603923969
    Selection.MoveRight Unit:=wdCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText text:="Tên nhóm"
    Selection.Shading.BackgroundPatternColor = -603923969
    Selection.MoveRight Unit:=wdCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText text:="T" & ChrW(7915) & " câu "
    Selection.Shading.BackgroundPatternColor = -603923969
    Selection.MoveRight Unit:=wdCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText text:="Ð" & ChrW(7871) & "n" & " câu "
    Selection.Shading.BackgroundPatternColor = -603923969
    
    If tuluan Then
        For i = 2 To sodong - 1
        Selection.Tables(1).Cell(i, 1).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=i - 1
        Selection.Shading.BackgroundPatternColor = -603923969
        Selection.Tables(1).Cell(i, 3).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Tables(1).Cell(i, 4).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next i
        Selection.Tables(1).Cell(sodong, 1).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:="TL"
        Selection.Shading.BackgroundPatternColor = -603923969
        
        Selection.Tables(1).Cell(2, 2).Select
        Selection.TypeText text:="PH" & ChrW(7846) & "N I: TR" & ChrW(7854) & _
        "C NGHI" & ChrW(7878) & "M KHÁCH QUAN"
        
        Selection.Tables(1).Cell(sodong, 2).Select
        Selection.TypeText text:="PH" & ChrW(7846) & "N II: T" & ChrW(7920) & _
        " LU" & ChrW(7852) & "N"
    Else
        For i = 2 To sodong
        Selection.Tables(1).Cell(i, 1).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.TypeText text:=i - 1
        Selection.Shading.BackgroundPatternColor = -603923969
        Selection.Tables(1).Cell(i, 3).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Tables(1).Cell(i, 4).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next i
        
    End If
    Selection.Tables(1).Select
    Call FontFormat
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.HomeKey Unit:=wdLine
    Call FontFormat
    Selection.TypeText text:="[</Gr>]"
    Selection.TypeParagraph
    Unload S_QG
End Sub
Sub S_Messages()
S_MsgBox.Label1.Top = 200
S_MsgBox.Show
End Sub
Sub S_PageSetup()
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(0.8)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1.9)
        .RightMargin = CentimetersToPoints(0.9)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.8)
        .FooterDistance = CentimetersToPoints(0.7)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .VerticalAlignment = wdAlignVerticalTop
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = True
    End With
End Sub
Sub QuestionTable()
On Error GoTo Thoat
    If S_QTable.OptionButton1 Then
        ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=3, NumColumns:=3, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
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
        End With
        'Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=100, RulerStyle:=wdAdjustNone
        'Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=200, RulerStyle:=wdAdjustNone
        Selection.Tables(1).Cell(1, 3).Select
        Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
        Selection.Cells.Merge
        Selection.Tables(1).Cell(1, 1).Select
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.Cells.Merge
        Selection.Tables(1).Cell(2, 1).Select
        Selection.TypeText text:="A."
        Selection.Tables(1).Cell(2, 2).Select
        Selection.TypeText text:="B."
        Selection.Tables(1).Cell(3, 1).Select
        Selection.TypeText text:="C."
        Selection.Tables(1).Cell(3, 2).Select
        Selection.TypeText text:="D."
    End If
    If S_QTable.OptionButton2 Then
        ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:=4, _
             DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
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
        End With
        Selection.Tables(1).Cell(1, 1).Select
        Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
        Selection.Cells.Merge
        Selection.Tables(1).Cell(2, 1).Select
        Selection.TypeText text:="A."
        Selection.Tables(1).Cell(2, 2).Select
        Selection.TypeText text:="B."
        Selection.Tables(1).Cell(2, 3).Select
        Selection.TypeText text:="C."
        Selection.Tables(1).Cell(2, 4).Select
        Selection.TypeText text:="D."
    End If
Exit Sub
Thoat:
MsgBox "Khoi tao duoc cau hoi bi lôi. Xem lai vi tri con tro."
End Sub
Sub delTable()
For i = ActiveDocument.Tables.Count To 1 Step -1
ActiveDocument.Tables(i).Select
Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=False
Next i
MsgBox "Xong!"
End Sub
Sub ConvertTestPro()
    On Error GoTo S_Quit
    n = ActiveDocument.Tables(1).Rows.Count
    Selection.Tables(1).Columns(1).Select
        With Selection.Font
            .Name = "Times New Roman"
            .Size = 12
            .Bold = True
            .Italic = False
        End With
    For i = n To 6 Step -6
    Selection.Tables(1).Cell(i, 2).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Select Case Trim(Selection)
    Case "A"
    Selection.Tables(1).Cell(i - 4, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "B"
    Selection.Tables(1).Cell(i - 3, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "C"
    Selection.Tables(1).Cell(i - 2, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "D"
    Selection.Tables(1).Cell(i - 1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "-A"
    Selection.Tables(1).Cell(i - 1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Italic = True
        End With
    Selection.Tables(1).Cell(i - 4, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "-B"
    Selection.Tables(1).Cell(i - 1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Italic = True
        End With
    Selection.Tables(1).Cell(i - 3, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "-C"
    Selection.Tables(1).Cell(i - 1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Italic = True
        End With
    Selection.Tables(1).Cell(i - 2, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Underline = wdUnderlineSingle
        End With
    Case "-D"
    Selection.Tables(1).Cell(i - 1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
        With Selection.Font
            .Italic = True
            .Underline = wdUnderlineSingle
        End With
    End Select
    
    Selection.Tables(1).Cell(i, 1).Select
    Selection.TypeText text:="[<Br>]"
    Selection.Tables(1).Cell(i, 2).Delete
    Selection.Tables(1).Cell(i - 5, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    Selection.TypeText text:=". "
    Next i
    
    Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=False
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^t"
        .Replacement.text = ""
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
Exit Sub
S_Quit:
MsgBox "Không chuyên duoc!"
End Sub
Sub Chuan_hoa_1()
    On Error Resume Next
    Application.ScreenUpdating = True
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.WholeStory
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0)
        .LeftIndent = CentimetersToPoints(0.5)
        .RightIndent = CentimetersToPoints(0)
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Font.ColorIndex = wdRed
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
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
    With Selection.Find
        .text = "^p^t"
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
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "([.:,\)])( )"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "([^13^32^9])([Aa])([.:\)])"
        .Replacement.text = "#A."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Bb])([.:\)])"
        .Replacement.text = "#B."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Cc])([.:\)])"
        .Replacement.text = "#C."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Dd])([.:\)])"
        .Replacement.text = "#D."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "C©u"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
        
    With Selection.Find
        .text = "Caâu"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = ""
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
    
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find.Replacement.Font
       .Bold = True
       .Color = wdColorBlue
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
       .Bold = True
       .Color = wdColorBlue
       .Italic = False
    End With
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "(^13)([ABCD])"
        .Replacement.text = "\1\2"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "#A."
        .Replacement.text = "^pA. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#B."
        .Replacement.text = "^tB. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
        .text = "^pB."
        .Replacement.text = "^pB. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#C."
        .Replacement.text = "^tC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^pC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#D."
        .Replacement.text = "^tD. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pD."
        .Replacement.text = "^pD. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
       .Bold = False
       .Color = wdColorBlack
    End With
    Selection.Find.ClearFormatting
   
    With Selection.Find
        .text = ".^t"
        .Replacement.text = ".^t"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([ABCD])(. .)"
        .Replacement.text = "\1" & ". "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
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
        .MatchCase = True
        .Format = True
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    
        Selection.WholeStory
        Selection.ParagraphFormat.TabStops.ClearAll
        ActiveDocument.DefaultTabStop = CentimetersToPoints(0.5)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(0.5) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(9.5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(14), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
 
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}.)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}:)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    
    Set danhsach = ActiveDocument.Content
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
    danhsach.Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.ParagraphFormat.TabStops.ClearAll
    'ActiveDocument.DefaultTabStop = CentimetersToPoints(1.27)
    'Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
    .NumberFormat = "Câu " & "%1."
    .TrailingCharacter = wdTrailingTab
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = CentimetersToPoints(0)
    .Alignment = wdListLevelAlignLeft
    .TextPosition = CentimetersToPoints(0)
    .TabPosition = wdUndefined
    .ResetOnHigher = 0
    .StartAt = 1
    .LinkedStyle = ""
    .Font.Bold = True
    .Font.Color = wdColorBlue
    .Font.Italic = False
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
    True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
    wdWord10ListBehavior
    Selection.Delete Unit:=wdCharacter, Count:=1
    GoTo Tiep
    End If
   
    Selection.WholeStory
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.6)
        .FooterDistance = CentimetersToPoints(0.6)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
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
    Selection.HomeKey Unit:=wdStory
    'If ktBanQuyen = False Then Call S_SerialHDD
    'If ktBanQuyen = False Then S_NoteRig.Show
End Sub
Sub Chuan_hoa_2()
    On Error Resume Next
    Application.ScreenUpdating = True
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Font.ColorIndex = wdRed
    With Selection.Find
        .text = "([ABCD])"
        .Replacement.text = "\1" & "."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
'Exit Sub
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
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
    With Selection.Find
        .text = "^p^t"
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
    'Selection.Find.Replacement.ClearFormatting
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
        .text = "([.:,\)])( )"
        .Replacement.text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "([^13^32^9])([Aa])([.:\)])"
        .Replacement.text = "#A."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        '.Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Bb])([.:\)])"
        .Replacement.text = "#B."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Cc])([.:\)])"
        .Replacement.text = "#C."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "([^32^9])([Dd])([.:\)])"
        .Replacement.text = "#D."
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "C©u"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
        
    With Selection.Find
        .text = "Caâu"
        .Replacement.text = "Câu"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "^t"
        .Replacement.text = ""
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
    
   
    With Selection.Find.Replacement.Font
       .Bold = True
       .Color = wdColorBlue
    End With
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
       .Bold = True
       .Color = wdColorBlue
       .Italic = False
    End With
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "(^13)([ABCD])"
        .Replacement.text = "\1\2"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("Normal")
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0)
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
    End With
'Exit Sub
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "#A."
        .Replacement.text = "^pA. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
       
    With Selection.Find
        .text = "#B."
        .Replacement.text = "^tB. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pB."
        .Replacement.text = "^pB. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#C."
        .Replacement.text = "^tC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^pC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pC."
        .Replacement.text = "^pC. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "#D."
        .Replacement.text = "^tD. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .text = "^pD."
        .Replacement.text = "^pD. "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
       .Bold = False
       .Color = wdColorBlack
    End With
    Selection.Find.ClearFormatting
    
    With Selection.Find
        .text = ".^t"
        .Replacement.text = ".^t"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .text = "([ABCD])(. .)"
        .Replacement.text = "\1" & ". "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
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
        .MatchCase = True
        .Format = True
        .MatchWildcards = False
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
        Selection.WholeStory
        
        Selection.ParagraphFormat.TabStops.ClearAll
        ActiveDocument.DefaultTabStop = CentimetersToPoints(1.75)
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(1.75) _
        , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(6), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(10), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.add Position:=CentimetersToPoints(14), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
 
    With Selection.Find
        .text = "(Câu [0-9]{1,4}.)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}:)"
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    
    Set danhsach = ActiveDocument.Content
Tiep:
    danhsach.Find.Execute FindText:="#", Forward:=True
    If danhsach.Find.Found = True Then
    danhsach.Select
    Selection.ParagraphFormat.TabStops.ClearAll
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphJustify
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-1.75)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
            
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
    .NumberFormat = "Câu " & "%1."
    .TrailingCharacter = wdTrailingTab
    .NumberStyle = wdListNumberStyleArabic
    .NumberPosition = CentimetersToPoints(0)
    .Alignment = wdListLevelAlignLeft
    .TextPosition = CentimetersToPoints(0)
    .TabPosition = wdUndefined
    .ResetOnHigher = 0
    .StartAt = 1
    .LinkedStyle = ""
    .Font.Bold = True
    .Font.Color = wdColorBlue
    .Font.Italic = False
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
    True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
    wdWord10ListBehavior
    Selection.Delete Unit:=wdCharacter, Count:=1
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.75)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(-1.75)
    End With
    GoTo Tiep
    End If
    
    Selection.WholeStory
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0.6)
        .FooterDistance = CentimetersToPoints(0.6)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .text = ". "
        .Replacement.text = ". "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
        Selection.WholeStory
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 12
        Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        Dim i As Integer
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
    Selection.HomeKey Unit:=wdStory
    If ktBanQuyen = False Then Call S_SerialHDD
    If ktBanQuyen = False Then S_NoteRig.Show
End Sub

Sub ConvertText2Auto()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}.)"
        .Replacement.text = "$"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,4}:)"
        .Replacement.text = "$"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "$^t"
        .Replacement.text = "$"
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
        .text = "$ "
        .Replacement.text = "$"
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
    danhsach.Find.Execute FindText:="$", Forward:=True
    If danhsach.Find.Found = True Then
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
    .TextPosition = CentimetersToPoints(0)
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
    End If
    Selection.HomeKey Unit:=wdStory
MsgBox "Done"
Call S_SerialHDD
If ktBanQuyen = False Then S_NoteRig.Show
End Sub
Sub ConvertText2Auto2()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    With Selection.Find.Replacement.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
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
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(^13)([0-9]{1,4})([.:)])"
        .Replacement.text = "^13" & "Câu " & "\2\3"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .text = "(Câu )([0-9]{1,3})(.)"
        .Replacement.text = "\1\2\3"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
MsgBox "Done"
End Sub
Sub ConvertAuto2Text()
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    MsgBox "Done"
End Sub
Sub ConvertABCD()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    With Selection.Find
        .text = "^pa" & S_QTable.TextBox1
        .Replacement.text = "^pA" & S_QTable.TextBox2
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
        .text = " a" & S_QTable.TextBox1
        .Replacement.text = " A" & S_QTable.TextBox2
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
        .text = "^ta" & S_QTable.TextBox1
        .Replacement.text = "^tA" & S_QTable.TextBox2
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
        .text = "^pb" & S_QTable.TextBox1
        .Replacement.text = "^pB" & S_QTable.TextBox2
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
        .text = " b" & S_QTable.TextBox1
        .Replacement.text = " B" & S_QTable.TextBox2
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
        .text = "^tb" & S_QTable.TextBox1
        .Replacement.text = "^tB" & S_QTable.TextBox2
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
        .text = "^pc" & S_QTable.TextBox1
        .Replacement.text = "^pC" & S_QTable.TextBox2
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
        .text = " c" & S_QTable.TextBox1
        .Replacement.text = " C" & S_QTable.TextBox2
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
        .text = "^tc" & S_QTable.TextBox1
        .Replacement.text = "^tC" & S_QTable.TextBox2
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
        .text = "^pd" & S_QTable.TextBox1
        .Replacement.text = "^pD" & S_QTable.TextBox2
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
        .text = " d" & S_QTable.TextBox1
        .Replacement.text = " D" & S_QTable.TextBox2
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
        .text = "^td" & S_QTable.TextBox1
        .Replacement.text = "^tD" & S_QTable.TextBox2
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
     MsgBox "Done"
End Sub
Function docIsOpen(DocName As String) As Boolean
    docIsOpen = False
    Dim wkbk As Document
    Dim opened As Boolean
    For Each wkbk In Application.Documents
        opened = UCase(wkbk.Name) = UCase(DocName)
        If opened Then
            docIsOpen = True
        End If
    Next wkbk
End Function
