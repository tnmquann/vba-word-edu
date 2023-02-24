Attribute VB_Name = "S_Test"
Option Explicit
Dim a() As Integer
Dim b() As Integer
Dim socau As String
Dim i, tucau As Integer
Dim S_sode As Integer
Dim dapanmoi() As Integer
Dim d_a1() As String
Dim S_error() As String
Dim demkt As Byte
Public ktlop As Integer
Public ktexist1, ktexist2, ktMark, ktOpen, ktMix, ktCD, ktInTheoBai, ktInTheoCD As Boolean
Public ale, ten_file_nguon, S_Drive As String
Public S_Khode As String

Private Sub XaoSo(ByRef so1 As Integer, ByRef so2 As Integer)
Dim tg As Integer
    tg = so1
    so1 = so2
    so2 = tg
End Sub

Private Sub RandNum3(ByRef dau As Integer, ByRef cuoi As Integer)
Dim iR3 As Integer
Randomize
    For iR3 = dau To cuoi
        a(iR3) = iR3
        b(iR3) = iR3
    Next
    If S_mf.CheckBox1 Then Exit Sub
    For iR3 = dau To cuoi
        Call XaoSo(a(iR3), a(Int(Rnd * (cuoi - dau)) + dau))
    Next
    'Dim tex As String
    'tex = ""
    'For i = 1 To cuoi
    'tex = tex & a(i) & "_"
    'Next i
    'MsgBox tex
End Sub
Private Sub RandNum4(ByRef dau As Integer, ByRef cuoi As Integer)
If S_mf.CheckBox1 Then Exit Sub
Dim iR4 As Integer
Randomize
    For iR4 = dau To cuoi
        b(iR4) = iR4
    Next
    For iR4 = dau To cuoi
        Call XaoSo(b(iR4), b(Int(Rnd * (cuoi - dau)) + dau))
    Next
    'Dim tex As String
    'tex = ""
    'For i = 1 To cuoi
    'tex = tex & b(i) & "_"
    'Next i
    'MsgBox tex
End Sub
Sub S_Mix(ByRef in_Name As String, ByRef out_Name As String, ByRef header_Name As String, _
ByRef footer_Name As String, ByRef Answer_Name As String, ByRef socau_Name As Integer, ByRef hoanvi_Name As Byte)
    Dim i_new, i1, i2, i3, i4, tam As String
    Dim t1, t2, t3, t4 As String
    Dim j, gr As Integer
    Dim tmax, sodong, socot As Byte
    Dim S_data, S_Header, S_Footer As New Word.Document
    Dim chondoan As Range
    Dim ten As String
    Dim f_nguon As String
    Dim f_dich As String
    Dim Dapan() As Integer
    Dim InAns() As String
    Dim myRange As Range
    Dim Title, msg As String
    Dim docOpener As Document
    On Error GoTo S_Quit
    ktOpen = False
    ktMix = False
    tucau = Val(S_mf.mf_t5)
    Select Case ktlop
        Case 13
            f_nguon = S_Drive & "S_Bank&Test\S_Data\Other\"
            f_dich = S_Drive & "S_Bank&Test\S_Test\Other\" & out_Name & "\"
        Case 12
            f_nguon = S_Drive & "S_Bank&Test\S_Data\Lop 12\"
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 12\" & out_Name & "\"
        Case 11
            f_nguon = S_Drive & "S_Bank&Test\S_Data\Lop 11\"
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 11\" & out_Name & "\"
        Case 10
            f_nguon = S_Drive & "S_Bank&Test\S_Data\Lop 10\"
            f_dich = S_Drive & "S_Bank&Test\S_Test\Lop 10\" & out_Name & "\"
    End Select
   
    Dim www As New Word.Application
    Set S_data = www.Documents.Open(f_nguon & in_Name & ".dat", PasswordDocument:="159")
    Set chondoan = www.ActiveDocument.Range( _
                    Start:=www.ActiveDocument.Bookmarks("c1a").Range.Start, _
                    End:=www.ActiveDocument.Bookmarks("c1b").Range.End)
    ktOpen = True
    'Exit Sub
    ten = www.Selection.Bookmarks(1).Name
    Dim grsocau() As String
    grsocau = Split(ten, "G")
    socau = Val(socau_Name)
    ReDim Dapan(Val(grsocau(1))) As Integer
    ReDim dapanmoi(Val(grsocau(1))) As Integer
    ReDim a(Val(grsocau(1))) As Integer
    ReDim b(Val(grsocau(1))) As Integer
    If socau > Val(grsocau(1)) Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "S" & ChrW(7889) & " câu " & ChrW(273) & "ã ch" & _
         ChrW(7885) & "n v" & ChrW(432) & ChrW(7907) & "t quá s" & ChrW(7889) & _
        " câu có trong ngân hàng câu h" & ChrW(7887) & "i." & Chr(13) & "S" & ChrW(7889) _
        & " câu trong ngân hàng là " & grsocau(1) & "."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        If ktOpen = True Then
            S_data.Close
            www.Quit
            Set www = Nothing
        End If
        Exit Sub
    End If
    Dim tmp5() As String
    Dim aa As Integer
    Dim ktTL As Boolean
    aa = UBound(grsocau)
    ktTL = False
    If grsocau(aa) = "TL" Then
    aa = aa - 1
    ktTL = True
    End If
    Call S_SerialHDD
    If ktTL And socau > 18 And ktBanQuyen = False Then
        S_Free.Show
        S_data.Close
        www.Quit
        Set www = Nothing
        Exit Sub
    End If
    
    ReDim tmp5(aa) As String
    If socau < Val(grsocau(1)) Then
    tmp5(2) = grsocau(2)
    For i = 3 To aa - 1
    tmp5(i) = tmp5(i - 1) + Round(((grsocau(i) - grsocau(i - 1))) * (socau / grsocau(1)))
    Next i
    tmp5(2) = grsocau(2)
    tmp5(aa) = socau + 1
    Else
    For i = 1 To aa
    tmp5(i) = grsocau(i)
    Next i
    End If
    'Tao thu muc luu de
    Select Case ktlop
        Case 13
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Other\" & out_Name & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Other\" & out_Name & "\")
            End If
        Case 12
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 12\" & out_Name & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 12\" & out_Name & "\")
            End If
        Case 11
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 11\" & out_Name & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 11\" & out_Name & "\")
            End If
        Case 10
            If DirExists(S_Drive & "S_Bank&Test\S_Test\Lop 10\" & out_Name & "\") = False Then
                MkDir (S_Drive & "S_Bank&Test\S_Test\Lop 10\" & out_Name & "\")
            End If
    End Select
    ReDim InAns(hoanvi_Name) As String
    'Kiem tra ton tai Header
    If FExists(S_Drive & "S_Bank&Test\S_Templates\default_Header_" & Right(header_Name, 1) & ".docx") = False Then
            Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "Header " & ChrW(273) & "ã ch" & ChrW(7885) & "n không có trong th" & _
            ChrW(432) & " m" & ChrW(7909) & "c S_Templates. B" & ChrW(7841) & "n ki" _
            & ChrW(7875) & "m tra và th" & ChrW(7921) & "c hi" & ChrW(7879) & "n l" & _
            ChrW(7841) & "i."
            Application.Assistant.DoAlert Title, msg, 0, 3, 0, 0, 0
            www.Quit
            Set www = Nothing
            Exit Sub
        End If
        'Kiem tre Header dang mo thi dong lai
        If docIsOpen("default_Header_" & Right(header_Name, 1) & ".docx") Then
            Set docOpener = Application.Documents("default_Header_" & Right(header_Name, 1) & ".docx")
            docOpener.Close
            Set docOpener = Nothing
        End If
    For S_sode = 1 To Val(hoanvi_Name)
       
''''''''''''''''''''''
    'IN HEADER
''''''''''''''''''''''
        Documents.add
        Call S_PageSetup
            Select Case header_Name
                Case "Header 1"
                    Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_1.docx")
                Case "Header 2"
                    Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_2.docx")
                Case "Header 3"
                    Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_3.docx")
                Case "Header 4"
                    Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_4.docx")
                Case "Header 5"
                    Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Header_5.docx")
            End Select
            www.Selection.Tables(1).Select
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            S_Header.Close
            Set S_Header = Nothing
        Dim ktAns As Boolean
        ktAns = False
        If Answer_Name = "Before" And Val(socau_Name) <= 50 Then
                ktAns = True
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx")
                Set myRange = www.ActiveDocument.Tables(Int(((socau - 1) / 5)) + 1).Range
                myRange.Copy
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
                Selection.TypeParagraph
        End If
        Dim MadeTmp As String
        MadeTmp = ""
        If S_mf.ListBox3.list(0) <> "" Then
        MadeTmp = S_mf.ListBox3.list(S_sode - 1)
        Else
        MadeTmp = S_sode Mod 10 & Int(89 * Rnd() + 10)
        End If
        If S_mf.CheckBox1 Then MadeTmp = "GOC"
        ActiveDocument.Variables("MADE") = MadeTmp
        ActiveDocument.Variables("<lop>") = ktlop
        ActiveDocument.Fields.Update
        If ktAns = True Then
        S_Header.Close
        ktAns = False
        Set S_Header = Nothing
        End If
''''''''''''''''''''''
        'Bat dau xao de
''''''''''''''''''''''
    For gr = 2 To (aa - 1)
        If www.ActiveDocument.Bookmarks.Exists("gr" & gr - 1) Then
        www.Selection.GoTo what:=wdGoToBookmark, Name:="gr" & gr - 1
        www.Selection.Copy
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Selection.TypeParagraph
        End If
        Call RandNum3(Val(grsocau(gr)), Val(grsocau(gr + 1)) - 1)
        Call RandNum4(Val(tmp5(gr)), Val(tmp5(gr + 1)) - 1)
        
        For i = Val(tmp5(gr)) To Val(tmp5(gr + 1)) - 1
            j = a(i + Val(grsocau(gr)) - Val(tmp5(gr)))
            
            If S_mf.CheckBox3 = False Then
                b(i) = b(i) - 4 * Int((b(i) + 3) / 4) + 4
            Else
                b(i) = a(i) - 4 * Int((a(i) + 3) / 4) + 4
            End If
            
            i1 = "a"
            i2 = "b"
            i3 = "c"
            i4 = "d"
            'Lam dap an
            Set chondoan = www.ActiveDocument.Range( _
                    Start:=www.ActiveDocument.Bookmarks("c" & j & "c").Range.Start, _
                    End:=www.ActiveDocument.Bookmarks("c" & j & "d").Range.End)
            Dim S_len As Byte
            ten = chondoan.Bookmarks(2).Name
            Dapan(j) = Val(Mid(ten, 5, 1))
            S_len = Val(Mid(ten, 6, 2))
            If S_len < 18 Then
                sodong = 1
                socot = 4
            ElseIf S_len >= 18 And S_len < 45 Then
                    sodong = 2
                    socot = 2
            Else
                    sodong = 4
                    socot = 1
            End If
        ''''''''''''''''
        If S_mf.CheckBox1 Then
            dapanmoi(i) = Dapan(i)
            GoTo S_Skip
        End If
        ''''''''
            Select Case Dapan(j) - b(i)
            Case 1
                tam = i1
                i1 = i2
                i2 = i3
                i3 = i4
                i4 = tam
                dapanmoi(i) = b(i)
                If Right(ten, 1) = "N" Then
                Select Case Right(ten, 2)
                Case "4N"
                    tam = i4
                    i4 = i3
                    i3 = tam
                    If b(i) = 4 Then dapanmoi(i) = 3
                    If b(i) = 3 Then dapanmoi(i) = 4
                Case "3N"
                    tam = i3
                    i3 = i2
                    i2 = tam
                    If b(i) = 3 Then dapanmoi(i) = 2
                    If b(i) = 2 Then dapanmoi(i) = 3
                Case "2N"
                    tam = i2
                    i2 = i1
                    i1 = tam
                    If b(i) = 2 Then dapanmoi(i) = 1
                    If b(i) = 1 Then dapanmoi(i) = 2
                Case "1N"
                    tam = i1
                    i1 = i4
                    i4 = tam
                    If b(i) = 1 Then dapanmoi(i) = 4
                    If b(i) = 4 Then dapanmoi(i) = 1
                End Select
                End If
            Case 2
                tam = i1
                i1 = i3
                i3 = tam
                tam = i2
                i2 = i4
                i4 = tam
                dapanmoi(i) = b(i)
                If Right(ten, 1) = "N" Then
                Select Case Right(ten, 2)
                Case "4N"
                    tam = i4
                    i4 = i2
                    i2 = tam
                    If b(i) = 4 Then dapanmoi(i) = 2
                    If b(i) = 2 Then dapanmoi(i) = 4
                Case "3N"
                    tam = i3
                    i3 = i1
                    i1 = tam
                    If b(i) = 3 Then dapanmoi(i) = 1
                    If b(i) = 1 Then dapanmoi(i) = 3
                Case "2N"
                    tam = i2
                    i2 = i4
                    i4 = tam
                    If b(i) = 2 Then dapanmoi(i) = 4
                    If b(i) = 4 Then dapanmoi(i) = 2
                Case "1N"
                    tam = i1
                    i1 = i3
                    i3 = tam
                    If b(i) = 1 Then dapanmoi(i) = 3
                    If b(i) = 3 Then dapanmoi(i) = 1
                End Select
                End If
            Case 3
                tam = i4
                i4 = i1
                i1 = tam
                dapanmoi(i) = b(i)
                If Right(ten, 1) = "N" Then
                Select Case Right(ten, 2)
                Case "4N"
                    tam = i4
                    i4 = i1
                    i1 = tam
                    If b(i) = 4 Then dapanmoi(i) = 1
                    If b(i) = 1 Then dapanmoi(i) = 4
                Case "1N"
                    tam = i1
                    i1 = i4
                    i4 = tam
                    If b(i) = 1 Then dapanmoi(i) = 4
                    If b(i) = 4 Then dapanmoi(i) = 1
                End Select
                End If
            Case -1
                tam = i4
                i4 = i3
                i3 = i2
                i2 = i1
                i1 = tam
                dapanmoi(i) = b(i)
                If Right(ten, 1) = "N" Then
                Select Case Right(ten, 2)
                Case "4N"
                    tam = i4
                    i4 = i1
                    i1 = tam
                    If b(i) = 4 Then dapanmoi(i) = 1
                    If b(i) = 1 Then dapanmoi(i) = 4
                Case "3N"
                    tam = i3
                    i3 = i4
                    i4 = tam
                    If b(i) = 3 Then dapanmoi(i) = 4
                    If b(i) = 4 Then dapanmoi(i) = 3
                Case "2N"
                    tam = i2
                    i2 = i3
                    i3 = tam
                    If b(i) = 2 Then dapanmoi(i) = 3
                    If b(i) = 3 Then dapanmoi(i) = 2
                Case "1N"
                    tam = i1
                    i1 = i2
                    i2 = tam
                    If b(i) = 1 Then dapanmoi(i) = 2
                    If b(i) = 2 Then dapanmoi(i) = 1
                End Select
                End If
            Case -2
                tam = i1
                i1 = i3
                i3 = tam
                tam = i2
                i2 = i4
                i4 = tam
                
                dapanmoi(i) = b(i)
                If Right(ten, 1) = "N" Then
                Select Case Right(ten, 2)
                Case "4N"
                    tam = i4
                    i4 = i2
                    i2 = tam
                    If b(i) = 4 Then dapanmoi(i) = 2
                    If b(i) = 2 Then dapanmoi(i) = 4
                Case "3N"
                    tam = i3
                    i3 = i1
                    i1 = tam
                    If b(i) = 3 Then dapanmoi(i) = 1
                    If b(i) = 1 Then dapanmoi(i) = 3
                Case "2N"
                    tam = i2
                    i2 = i4
                    i4 = tam
                    If b(i) = 2 Then dapanmoi(i) = 4
                    If b(i) = 4 Then dapanmoi(i) = 2
                Case "1N"
                    tam = i1
                    i1 = i3
                    i3 = tam
                    If b(i) = 1 Then dapanmoi(i) = 3
                    If b(i) = 3 Then dapanmoi(i) = 1
                End Select
                End If
            Case -3
                tam = i4
                i4 = i1
                i1 = tam
                dapanmoi(i) = b(i)
                If Right(ten, 1) = "N" Then
                Select Case Right(ten, 2)
                Case "4N"
                    tam = i4
                    i4 = i1
                    i1 = tam
                    If b(i) = 4 Then dapanmoi(i) = 1
                    If b(i) = 1 Then dapanmoi(i) = 4
                Case "1N"
                    tam = i1
                    i1 = i4
                    i4 = tam
                    If b(i) = 1 Then dapanmoi(i) = 4
                    If b(i) = 4 Then dapanmoi(i) = 1
                End Select
                End If
            Case 0
                dapanmoi(i) = b(i)
            End Select
S_Skip:
            www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & "q"
            www.Selection.Copy
            
            Selection.TypeText text:="Câu " & i + tucau - 1 & "."
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            With Selection.Font
                .Name = "Times New Roman"
                .Size = 12
                .Bold = True
                .Color = 13382400
            End With
            
            Selection.EndKey Unit:=wdLine, Extend:=wdMove
            Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
            If Selection.Tables.Count > 0 Then
                Select Case Mid(ten, 8, 2)
                Case "t3"
                    ''''''''''
                    Selection.Tables(1).Select
                    Selection.MoveUp Unit:=wdLine, Count:=1
                    Selection.HomeKey Unit:=wdLine
                    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
                    Selection.TypeBackspace
                    Selection.TypeBackspace
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.Tables(1).Cell(1, 1).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                    End With
                    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(-0.2)
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1 ', Extend:=wdExtend
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="Câu " & i + tucau - 1 & ". "
                    ''''''''''
                    Selection.Tables(1).Borders.InsideLineStyle = wdLineStyleNone
                    Selection.Tables(1).Borders.OutsideLineStyle = wdLineStyleNone
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i1
                    www.Selection.Copy
                    Selection.Tables(1).Cell(2, 1).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0.3)
                    Selection.TypeText text:="A."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i2
                    www.Selection.Copy
                    Selection.Tables(1).Cell(2, 2).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="B."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i3
                    www.Selection.Copy
                    Selection.Tables(1).Cell(3, 1).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0.3)
                    Selection.TypeText text:="C."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i4
                    www.Selection.Copy
                    Selection.Tables(1).Cell(3, 2).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="D."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    Selection.MoveDown Unit:=wdLine, Count:=2
                    GoTo S_skip2
                Case "t4"
                    Selection.Tables(1).Select
                    Selection.MoveUp Unit:=wdLine, Count:=1
                    Selection.HomeKey Unit:=wdLine
                    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
                    Selection.TypeBackspace
                    Selection.TypeBackspace
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.Tables(1).Cell(1, 1).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                    End With
                    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(-0.2)
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1 ', Extend:=wdExtend
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="Câu " & i + tucau - 1 & ". "
                    Selection.Tables(1).Borders.InsideLineStyle = wdLineStyleNone
                    Selection.Tables(1).Borders.OutsideLineStyle = wdLineStyleNone
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i1
                    www.Selection.Copy
                    Selection.Tables(1).Cell(2, 1).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="   A."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i2
                    www.Selection.Copy
                    Selection.Tables(1).Cell(2, 2).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="B."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i3
                    www.Selection.Copy
                    Selection.Tables(1).Cell(2, 3).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="C."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i4
                    www.Selection.Copy
                    Selection.Tables(1).Cell(2, 4).Select
                    With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        .Bold = True
                        .Color = 13382400
                    End With
                    Selection.TypeText text:="D."
                    Selection.PasteAndFormat (wdFormatOriginalFormatting)
                    Selection.MoveDown Unit:=wdLine, Count:=2
                    GoTo S_skip2
                Case Else
                    Selection.Tables(1).Select
                    Selection.MoveUp Unit:=wdLine, Count:=1
                    Selection.HomeKey Unit:=wdLine
                    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
                    If Left(Selection, 4) = "Câu " And Len(Selection) <= 11 Then
                        Selection.TypeBackspace
                        Selection.TypeBackspace
                        Selection.MoveDown Unit:=wdLine, Count:=1
                        Selection.Tables(1).Cell(1, 1).Select
                        With Selection.Font
                        .Name = "Times New Roman"
                        .Size = 12
                        End With
                        Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(-0.2)
                        Selection.MoveLeft Unit:=wdCharacter, Count:=1 ', Extend:=wdExtend
                        With Selection.Font
                            .Name = "Times New Roman"
                            .Size = 12
                            .Bold = True
                            .Color = 13382400
                        End With
                        Selection.TypeText text:="Câu " & i + tucau - 1 & ". "
                    End If
                    Selection.EndKey Unit:=wdStory
                    Selection.TypeParagraph
                End Select
            Else
                Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            End If
            
            Dim myTable As Table
            Set myTable = ActiveDocument.Tables.add(Range:=Selection.Range, _
            NumRows:=sodong, NumColumns:=socot)
           
            Select Case socot
            Case 4
            myTable.Columns(1).SetWidth ColumnWidth:=131, RulerStyle:=wdAdjustNone
            myTable.Columns(2).SetWidth ColumnWidth:=120, RulerStyle:=wdAdjustNone
            myTable.Columns(3).SetWidth ColumnWidth:=120, RulerStyle:=wdAdjustNone
            Case 2
            myTable.Columns(1).SetWidth ColumnWidth:=251, RulerStyle:=wdAdjustNone
            End Select
            www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i1
            www.Selection.Copy
            Application.Keyboard (1033)
            Selection.TypeText text:="A. "
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i2
            www.Selection.Copy
            
            Selection.MoveRight Unit:=wdCell
            Selection.TypeText text:="B. "
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i3
            www.Selection.Copy
            
            Selection.MoveRight Unit:=wdCell
            Selection.TypeText text:="C. "
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            www.Selection.GoTo what:=wdGoToBookmark, Name:="c" & j & i4
            www.Selection.Copy
            
            Selection.MoveRight Unit:=wdCell
            Selection.TypeText text:="D. "
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            Selection.Rows.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=False
            Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
            Call S_ParagaphFormat
            
            Selection.MoveDown Unit:=wdLine, Count:=1
            Set myTable = Nothing
S_skip2:
        Next i
     Next gr
        If ktTL Then
            Selection.TypeParagraph
            Selection.MoveUp Unit:=wdLine, Count:=1
            If www.ActiveDocument.Bookmarks.Exists("grTL") Then
            www.Selection.GoTo what:=wdGoToBookmark, Name:="grTL"
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            Selection.TypeParagraph
            End If
            www.Selection.GoTo what:=wdGoToBookmark, Name:="cTL"
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End If
''''''''''''''''''
        'Kiem tre Footer dang mo thi dong lai
        If docIsOpen("default_Footer_" & Right(footer_Name, 1) & ".docx") Then
            Set docOpener = Application.Documents("default_Footer_" & Right(footer_Name, 1) & ".docx")
            docOpener.Close
            Set docOpener = Nothing
        End If
        Selection.TypeParagraph
        Dim ktFooter As Boolean
        ktFooter = False
        Select Case footer_Name
        Case "Footer 1"
            ktFooter = True
            Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Footer_1.docx")
            www.Selection.WholeStory
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Case "Footer 2"
            ktFooter = True
            Set S_Footer = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Footer_2.docx")
            www.Selection.WholeStory
            www.Selection.Copy
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
        Case "Default"
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.TypeText text:="---------- " & "H" & ChrW(7870) & "T" & " ----------"
        End Select
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Call FontFormat2
        Selection.TypeText text:="Trang "
        Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            "PAGE  ", PreserveFormatting:=True
        Selection.TypeText text:="/"
        Selection.Fields.add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
            "NUMPAGES  ", PreserveFormatting:=True
        Selection.TypeText text:=" - M" & ChrW(227) & " " & ChrW(273) & ChrW(7873) & " thi "
        Selection.TypeText text:=MadeTmp
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        If ktFooter = True Then S_Footer.Close
        
        If Answer_Name = "After" And Val(socau_Name) <= 50 Then
                ktAns = True
                Set S_Header = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\default_Answer.docx")
                Set myRange = www.ActiveDocument.Tables(Int(((socau - 1) / 5)) + 1).Range
                myRange.Copy
                Selection.TypeParagraph
                Selection.TypeParagraph
                Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End If
        If ktAns = True Then S_Header.Close
'''''''''''''''''''
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Dim ktIn As Boolean
        ktIn = False
        If S_mf.ComboBox4 = "In chung voi de" Then
            Selection.TypeParagraph
            Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            Selection.TypeParagraph
            Selection.TypeText text:=ChrW(272) & "áp án:"
            Call in_dapan(S_mf.ListBox3.list(S_sode - 1))
            ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_mf.mf_t1 & "] Made " & MadeTmp & ".docx", FileFormat:= _
            wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False, CompatibilityMode:=15
            ktIn = True
        Else
            ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_mf.mf_t1 & "] Made " & MadeTmp & ".docx", FileFormat:= _
            wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False, CompatibilityMode:=15
            InAns(S_sode) = MadeTmp
            For i = 1 To socau
            InAns(S_sode) = InAns(S_sode) & dapanmoi(i)
            Next i
        End If
        ktMix = True
        If S_mf.CheckBox2 Or S_sode > 8 Then ActiveDocument.Close
    Next S_sode
    If ktOpen = True Then
    S_data.Close
    www.Quit
    Set S_data = Nothing
    Set www = Nothing
    End If
    If ktIn = False Then
            Documents.add
            Call S_PageSetup
            Selection.TypeText text:=ChrW(272) & "ÁP ÁN [" & S_mf.mf_t1 & "]:"
            ActiveDocument.SaveAs2 FileName:=f_dich & "\[" & S_mf.mf_t1 & "] Dapan" & ".docx", FileFormat:= _
                wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
                :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
            For i = 1 To Val(hoanvi_Name)
                For j = 1 To socau
                dapanmoi(j) = Mid(InAns(i), j + 3, 1)
                Next j
                Call in_dapan(Left(InAns(i), 3))
            Selection.MoveDown Unit:=wdLine, Count:=1
            Next i
            ActiveDocument.Save
        'End If
    End If
    Unload S_mf
    If ktBanQuyen = False Then S_Free.Show
Exit Sub
S_Quit:
    
    'Dim Title, msg As String
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Quá trình in " & ChrW(273) & ChrW(7873) & " x" & _
        ChrW(7843) & "y ra l" & ChrW(7895) & "i. B" & ChrW(7841) & "n có th" & _
        ChrW(7875) & " xem xét m" & ChrW(7897) & "t s" & ChrW(7889) & " g" & ChrW _
        (7907) & "i ý sau:" & Chr(13) & _
        "- Th" & ChrW(7921) & "c hi" & ChrW(7879) & "n l" & ChrW(7841) & "i thêm m" & ChrW(7897) & "t l" & ChrW(7847) & "n n" & _
        ChrW(7919) & "a. N" & ChrW(7871) & "u v" & ChrW(7851) & "n còn l" & ChrW( _
        7895) & "i thì hãy " & ChrW(273) & ChrW(7885) & "c l" & ChrW(7841) & "i h" & ChrW(432) & ChrW( _
        7899) & "ng d" & ChrW(7851) & "n nh" & ChrW(7853) & "p câu h" & ChrW(7887 _
        ) & "i (Câu h" & ChrW(7887) & "i ch" & ChrW(7913) & "a b" & ChrW(7843) & _
        "ng, câu h" & ChrW(7887) & "i t" & ChrW(7921) & " lu" & ChrW(7853) & "n,...)" & Chr(13) & _
        "- Ch" & ChrW(7885) & "n file ngu" & ChrW(7891) & _
        "n khác và in th" & ChrW(7917) & "." & Chr(13) & _
        "- Kh" & ChrW(7903) & "i " & ChrW(273) & ChrW( _
        7897) & "ng l" & ChrW(7841) & "i máy." & Chr(13) & _
        "N" & ChrW(7871) & "u v" & ChrW(7851) & _
        "n không kh" & ChrW(7855) & "c ph" & ChrW(7909) & "c " & ChrW(273) & ChrW _
        (432) & ChrW(7907) & "c vui lòng liên h" & ChrW(7879) & " tác gi" & ChrW( _
        7843) & ". Xin l" & ChrW(7895) & "i b" & ChrW(7841) & "n vì s" & ChrW _
        (7921) & " b" & ChrW(7845) & "t ti" & ChrW(7879) & "n này."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    www.Quit
    Set www = Nothing
End Sub
Sub S_Mark(ByRef in_Name As String)
        Dim er As String
        Dim chondoan As Range
        Dim L1, L2, L3, L4, d_a As Byte
        Dim C As Integer
        Dim lmax As String
        Dim Shape1, Shape2, Shape3, Shape4 As Byte
        Dim title2, msg As String
        Dim ktBr As Boolean
        Dim ktMsg As Byte
        'On Error GoTo S_Quit
        ktBr = True
        If DirExists(S_Drive & "S_Bank&Test\S_Data") = False Then
            Call MadeDir
        End If
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "(\[\<)([Bb])([Rr])(\>\])"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        If Selection.Find.Execute = False Then
            title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = "D" & ChrW(7919) & " li" & ChrW(7879) & "u c" & _
            ChrW(7911) & "a b" & ChrW(7841) & "n ch" & ChrW(432) & "a chèn ký hi" & _
            ChrW(7879) & "u [<Br>]." & Chr(13) & "N" & ChrW(7871) & "u d" & ChrW(7919) & " li" & _
            ChrW(7879) & "u c" & ChrW(7911) & "a b" & ChrW(7841) & "n " & ChrW(273) & _
             "ã " & ChrW(273) & "ánh th" & ChrW(7913) & " t" & ChrW(7921) & _
            " ""Câu x."" ho" & ChrW(7863) & "c ""Câu x:""" & Chr(13) & "ho" & ChrW(7863) & "c ""Câu x)"" và không có ph" & ChrW(7847) & "n t" & _
            ChrW(7921) & " lu" & ChrW(7853) & "n thì ch" & ChrW(432) & ChrW(417) & _
            "ng trình có th" & ChrW(7875) & " nh" & ChrW(7853) & "n di" & ChrW(7879) & "n theo ký hi" & ChrW(7879) _
             & "u Câu. B" & ChrW(7841) & "n có ti" & ChrW(7871) & "p t" & ChrW(7909) _
            & "c không?"
            ktMsg = Application.Assistant.DoAlert(title2, msg, 4, 2, 0, 0, 1)
            If ktMsg = 6 Then
                ktBr = False
            Else
                Exit Sub
            End If
        End If
        ktMix = False
        ActiveDocument.SaveAs2 FileName:=S_Drive & "S_Bank&Test\S_Data\tmp.docx", _
            FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False, CompatibilityMode:=15
        Call RemoveMarks
        If ActiveDocument.Tables.Count > 0 Then
            For i = 1 To ActiveDocument.Tables.Count
                ActiveDocument.Tables(i).Select
                Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
                Selection.HomeKey Unit:=wdLine, Extend:=wdMove
                Selection.TypeParagraph
            Next i
        End If
        Selection.WholeStory
        Selection.Range.HighlightColorIndex = wdNoHighlight
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Selection.TypeParagraph
        Selection.TypeText text:="Please wait...........B&T Program"
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
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
        
        If ktBr Then
            Selection.HomeKey Unit:=wdStory, Extend:=wdMove
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "(\[\<)([Bb])([Rr])(\>\])"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            
            Do While Selection.Find.Execute = True
                ktMix = True
                Selection.Collapse Direction:=wdCollapseEnd
                Selection.TypeText text:=" "
                C = C + 1
                Call ClearBlankAfBreak
                Call Check_Br
                With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & C & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
                End With
            Loop
        Else
            Selection.HomeKey Unit:=wdStory, Extend:=wdMove
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            For i = 1 To ActiveDocument.Tables.Count
                ActiveDocument.Tables(i).Cell(1, 1).Select
                Selection.Find.Execute FindText:="(Câu)(*)([.:\)])", MatchWildcards:=True
                If Selection.Find.Found = True Then
                    Selection.TypeBackspace
                    ActiveDocument.Tables(i).Select
                    Selection.MoveLeft Unit:=wdCharacter, Count:=2
                    Selection.TypeParagraph
                    Selection.TypeText text:="$000#"
                End If
            Next i
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(Câu[ ]{1,3}[0-9]{1,3}[.:\)])"
                .Replacement.text = "$000#"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
         'Exit Sub
            Dim stt As String
            Selection.HomeKey Unit:=wdStory, Extend:=wdMove
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "$???#"
                .MatchWildcards = True
            End With
            Do While Selection.Find.Execute = True
                ktMix = True
                Selection.Collapse Direction:=wdCollapseStart
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
                If C < 10 Then
                    stt = "00" & C
                ElseIf C < 100 Then
                    stt = "0" & C
                Else
                    stt = "000"
                End If
                Selection.TypeText text:=stt
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Call ClearBlankAfBreak
                'Call Check_Br
                With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & C & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
                End With
                C = C + 1
            Loop
        End If
        Dim Title, dem As String
        Dim donggr As Byte
        Dim ktgroup, ktTL As Boolean
        ktgroup = False
        ktTL = False
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        Selection.Find.Execute FindText:="(\[\<)([Gg])([Rr])(\>\])(*)(\[\<\/)([Gg])([Rr])(\>\])", _
            MatchWildcards:=True
        If Selection.Find.Found = True Then
            ActiveDocument.Tables(1).Select
            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdMove
            Selection.HomeKey Unit:=wdLine, Extend:=wdMove
            Selection.Delete
            ktgroup = True
            donggr = ActiveDocument.Tables(1).Rows.Count
            ActiveDocument.Tables(1).Cell(donggr, 1).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            dem = Selection
            If dem = "TL" Then
                donggr = donggr - 1
                ktTL = True
                C = C - 1
                ActiveDocument.Tables(1).Cell(donggr + 1, 2).Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="grTL"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
                End With
            End If
            For i = 2 To donggr
                ActiveDocument.Tables(1).Cell(i, 2).Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                dem = Selection
                If Len(dem) > 2 Then
                With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="gr" & i - 1
                .DefaultSorting = wdSortByName
                .ShowHidden = True
                End With
                End If
                ActiveDocument.Tables(1).Cell(i, 3).Select
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                dem = Selection
                Title = Title & "G" & dem
            Next i
            Else
            Title = Title & "G1"
        End If
        Title = Title & "G" & C
        If ktTL Then Title = Title & "GTL"
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="socauG" & C - 1 & Title
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        If ktBr Then
            If ktgroup Then
                ActiveDocument.Tables(1).Select
                Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdMove
                Selection.HomeKey Unit:=wdLine, Extend:=wdMove
                Call ClearBlankAfBreak
                Call Check_Br
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c1q"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="s1"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
            Else
                Call ClearBlankAfBreak
                Call Check_Br
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c1q"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="s1"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
            End If
        Else
            Selection.EndKey Unit:=wdStory
            With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & C & "q"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
            End With
            Selection.HomeKey Unit:=wdLine
            Selection.TypeText text:="$000# "
            Selection.GoTo what:=wdGoToBookmark, Name:="c1q"
            With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="s1"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
            End With
        End If
        'MsgBox c - 1
        ''''
        'If ktBr = False Then c = c + 1
        ReDim S_error(C)
        For i = 1 To C - 1
            S_error(i) = ""
        Next i
        Dim NhomTL As String
        If ktgroup Then NhomTL = NhomTL & "[Nhom]"
        If ktTL Then NhomTL = NhomTL & "[TL]"
        ''''
        Dim ktA, ktB, ktC, ktD, ktd_a As Integer
        Dim ktItalic As String
        Dim ktTab As String
        Dim ktGood As Integer
        Dim S_text As String
        Dim choiceA, choiceB, choiceC, choiceD As String
        Dim myRange As Range
        
        If S_inf.CheckBox1 Then
            choiceA = "([^13^32^9])([A])(.)(*)([^13^32^9])([B])(.)"
            choiceB = "([^13^32^9])([B])(.)(*)([^13^32^9])([C])(.)"
            choiceC = "([^13^32^9])([C])(.)(*)([^13^32^9])([D])(.)"
            If ktBr Then
                choiceD = "([^13^32^9])([D])(.)(*)(\[\<)([Bb])([Rr]\>\])"
            Else
                choiceD = "([^13^32^9])([D])(.)(*)($???#)"
            End If
        Else
            choiceA = "([^13^32^9])([Aa])(.)(*)([^13^32^9])([Bb])(.)"
            choiceB = "([^13^32^9])([Bb])(.)(*)([^13^32^9])([Cc])(.)"
            choiceC = "([^13^32^9])([Cc])(.)(*)([^13^32^9])([Dd])(.)"
            If ktBr Then
                choiceD = "([^13^32^9])([Dd])(.)(*)(\[\<)([Bb])([Rr]\>\])"
            Else
                choiceD = "([^13^32^9])([Dd])(.)(*)($???#)"
            End If
            
        End If
        
        If ktTL Then
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & C & "q"
            Selection.HomeKey Unit:=wdLine
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & C & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & C & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & C + 1 & "q").Range.End)
            myRange.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.MoveLeft Unit:=wdCharacter, Count:=8, Extend:=wdExtend
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & "TL"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
        er = 1
        For i = 1 To C - 1
            S_text = ""
            ktA = 0
            ktB = 0
            ktC = 0
            ktD = 0
            ktd_a = 0
            ktGood = True
            ktItalic = "00"
            ktTab = ""
            Selection.Find.ClearFormatting
            
            'Danh dau phuong an A
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
        '''''''''''''''
            If myRange.Tables.Count > 0 Then
            Select Case myRange.Tables(1).Columns.Count
            Case 3
                If myRange.Tables(1).Rows.Count <> 3 Then
                    myRange.Tables(1).Range.HighlightColorIndex = wdYellow
                    myRange.Tables(1).Borders.InsideLineStyle = wdLineStyleSingle
                    myRange.Tables(1).Borders.OutsideLineStyle = wdLineStyleSingle
                    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
                    msg = "Ch" & ChrW(432) & ChrW(417) & "ng trình phát hi" _
                     & ChrW(7879) & "n câu h" & ChrW(7887) & "i th" & ChrW(7913) & " " & i & " ch" & _
                    ChrW(7913) & "a Table không " & ChrW(273) & "úng " & ChrW(273) & ChrW( _
                    7883) & "nh d" & ChrW(7841) & "ng c" & ChrW(7911) & "a ch" _
                     & ChrW(432) & ChrW(417) & "ng trình." & Chr(13) & _
                     "Hãy " & ChrW(273) & ChrW(7885) & "c k" & ChrW( _
                    7929) & " h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n cách t" _
                    & ChrW(7841) & "o câu h" & ChrW(7887) & "i ch" & ChrW(7913) & "a Table."
                    Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
                    
                    Exit Sub
                End If
                myRange.Tables(1).Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "q"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Selection.Tables(1).Cell(2, 1).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "a"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Selection.Tables(1).Cell(2, 2).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "b"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Selection.Tables(1).Cell(3, 1).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="da00110t3" & i
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "c"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                
                Selection.Tables(1).Cell(3, 2).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "d"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                d_a = 1
                ktA = 1
                ktB = 1
                ktC = 1
                ktD = 1
                ktd_a = 1
                lmax = 10
                ktTab = "t3"
                Selection.MoveDown Unit:=wdLine, Count:=1
                GoTo S_Skip1
                
            Case 4
                If myRange.Tables(1).Rows.Count <> 2 Then
                    myRange.Tables(1).Range.HighlightColorIndex = wdYellow
                    myRange.Tables(1).Borders.InsideLineStyle = wdLineStyleSingle
                    myRange.Tables(1).Borders.OutsideLineStyle = wdLineStyleSingle
                    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
                    msg = "Ch" & ChrW(432) & ChrW(417) & "ng trình phát hi" _
                     & ChrW(7879) & "n câu h" & ChrW(7887) & "i th" & ChrW(7913) & " " & i & " ch" & _
                    ChrW(7913) & "a Table không " & ChrW(273) & "úng " & ChrW(273) & ChrW( _
                    7883) & "nh d" & ChrW(7841) & "ng c" & ChrW(7911) & "a ch" _
                     & ChrW(432) & ChrW(417) & "ng trình." & Chr(13) & _
                     "Hãy " & ChrW(273) & ChrW(7885) & "c k" & ChrW( _
                    7929) & " h" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n cách t" _
                    & ChrW(7841) & "o câu h" & ChrW(7887) & "i ch" & ChrW(7913) & "a Table."
                    Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
                    
                    Exit Sub
                End If
                myRange.Tables(1).Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "q"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Selection.Tables(1).Cell(2, 1).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="da00110t4" & i
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "a"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                
                Selection.Tables(1).Cell(2, 2).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "b"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Selection.Tables(1).Cell(2, 3).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "c"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Selection.Tables(1).Cell(2, 4).Select
                Set myRange = Selection.Range
                myRange.MoveEnd Unit:=wdCharacter, Count:=-1
                myRange.MoveStart Unit:=wdCharacter, Count:=3
                myRange.Select
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "d"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                ktA = 1
                ktB = 1
                ktC = 1
                ktD = 1
                ktd_a = 1
                lmax = 10
                d_a = 1
                ktTab = "t4"
                Selection.MoveDown Unit:=wdLine, Count:=1
                GoTo S_Skip1
            End Select
            End If
        '''''''''''''''
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
            Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
            Call ClearBlankBf
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
                End:=ActiveDocument.Bookmarks("s2").Range.End)
            myRange.Select
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Call ClearBlankBf
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
            Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                    d_a = 1
                    ktd_a = ktd_a + 1
            End If
            If Selection.Font.Italic Then
                    ktItalic = "1N"
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            Call ClearBlankAfABCD
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
                End:=ActiveDocument.Bookmarks("s2").Range.End)
            myRange.Select
            L1 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
                Case 1
                L1 = L1 + Round(myRange.InlineShapes(1).Width / 5.8)
                Case 2
                L1 = L1 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            If L1 > 0 Then
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "a"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktA = ktA + 1
            Selection.Find.ClearFormatting
            Selection.Find.Execute FindText:="([^13^32^9])([Aa])(.)", MatchWildcards:=True
                If Selection.Find.Found = True Then
                    ktA = ktA + 1
                    Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
                    Selection.Range.HighlightColorIndex = wdTurquoise
                End If
            End If
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
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Call ClearBlankBf
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
            Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                    d_a = 2
                    ktd_a = ktd_a + 1
            End If
            If Selection.Font.Italic Then
                    ktItalic = "2N"
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            Call ClearBlankAfABCD
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
                End:=ActiveDocument.Bookmarks("s2").Range.End)
            myRange.Select
            L2 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
                Case 1
                L2 = L2 + Round(myRange.InlineShapes(1).Width / 6)
                Case 2
                L2 = L2 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            If L2 > 0 Then
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "b"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktB = ktB + 1
            
            Selection.Find.ClearFormatting
            Selection.Find.Execute FindText:="([^13^32^9])([Bb])(.)", MatchWildcards:=True
            If Selection.Find.Found = True Then
                ktB = ktB + 1
                Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
                Selection.Range.HighlightColorIndex = wdTurquoise
            End If
            End If
        End If
           'Danh dau phuong an C
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Find.Execute FindText:=choiceC, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "c"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Call ClearBlankBf
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
            Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                    d_a = 3
                    ktd_a = ktd_a + 1
            End If
             If Selection.Font.Italic Then
                    ktItalic = "3N"
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            Call ClearBlankAfABCD
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
                End:=ActiveDocument.Bookmarks("s2").Range.End)
            myRange.Select
            L3 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
               Case 1
                L3 = L3 + Round(myRange.InlineShapes(1).Width / 6)
                Case 2
                L3 = L3 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            If L3 > 0 Then
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "c"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktC = ktC + 1
            
            Selection.Find.ClearFormatting
            Selection.Find.Execute FindText:="([^13^32^9])([Cc])(.)", MatchWildcards:=True
            If Selection.Find.Found = True Then
                ktC = ktC + 1
                Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
                Selection.Range.HighlightColorIndex = wdTurquoise
            End If
            End If
        End If
            'Danh dau phuong an D
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Find.Execute FindText:=choiceD, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            If ktBr Then
                myRange.MoveEnd Unit:=wdCharacter, Count:=-7
                myRange.Select
            Else
                myRange.MoveEnd Unit:=wdCharacter, Count:=-5
                myRange.Select
            End If
           
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "d"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Call ClearBlankBf
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s2"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
            Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                    d_a = 4
                    ktd_a = ktd_a + 1
            End If
            If Selection.Font.Italic Then
                    ktItalic = "4N"
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            Call ClearBlankAfABCD
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
                End:=ActiveDocument.Bookmarks("s2").Range.End)
            myRange.Select
            L4 = myRange.Characters.Count
            Select Case myRange.InlineShapes.Count
                Case 1
                L4 = L4 + Round(myRange.InlineShapes(1).Width / 6)
                Case 2
                L4 = L4 + Round((myRange.InlineShapes(1).Width + myRange.InlineShapes(2).Width) / 6.3)
            End Select
            If L4 > 0 Then
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "d"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktD = ktD + 1
            End If
        End If
            lmax = L1
            If Val(lmax) < L2 Then lmax = L2
            If Val(lmax) < L3 Then lmax = L3
            If Val(lmax) < L4 Then lmax = L4
            If Val(lmax) < 10 Then lmax = "0" & lmax
            If Val(lmax) > 60 Then lmax = 60
            If Val(lmax) = 0 Then GoTo S_Quit
S_Skip1:
            If ActiveDocument.Bookmarks.Exists("c" & i & "d") Then
                Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
                Selection.MoveLeft Unit:=wdCharacter, Count:=2
                If i < 10 Then
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="da" & "0" & i & d_a & lmax & ktTab & ktItalic
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                Else
                With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="da" & i & d_a & lmax & ktTab & ktItalic
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
                End With
                End If
            End If
        '''''''''
        If S_inf.CheckBox2 Then
            d_a = 1
            ktd_a = 1
        End If
        If ktA <> 1 And ktB = 1 Then
            S_text = S_text & "A, "
            ktGood = False
            Selection.MoveUp Unit:=wdLine, Extend:=wdExtend
            Selection.Range.HighlightColorIndex = wdYellow
        End If
        If ktB <> 1 And ktC = 1 Then
            S_text = S_text & "B, "
            ktGood = False
            Selection.MoveUp Unit:=wdLine, Extend:=wdExtend
            Selection.Range.HighlightColorIndex = wdYellow
        End If
        If ktC <> 1 And ktD = 1 Then
            S_text = S_text & "C, "
            ktGood = False
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Range.HighlightColorIndex = wdYellow
        End If
        If ktD <> 1 Then
            S_text = S_text & "D, "
            ktGood = False
            Selection.MoveUp Unit:=wdLine, Extend:=wdExtend
            Selection.Range.HighlightColorIndex = wdYellow
        End If
        If ktd_a <> 1 Then
            S_text = S_text & "có " & ktd_a & " " & ChrW(273) & "áp án. "
            ktGood = False
            Selection.MoveUp Unit:=wdLine, Extend:=wdExtend
            Selection.Range.HighlightColorIndex = wdYellow
        End If
        If ktA <> 1 And ktB <> 1 Then
            S_text = S_text & "L" & ChrW(7895) & "i n" & ChrW(7863) & "ng!"
            ktGood = False
        End If
        If ktGood = False And S_text <> "" Then
            If i < 10 Then i = "0" & i
            If ktA = 1 And ktB = 1 And ktC = 1 And ktD = 1 Then
                S_error(er) = "Câu " & i & " : " & S_text
            Else
                S_error(er) = "Câu " & i & ": Không xác " & ChrW(273) & ChrW(7883) & "nh " & _
                ChrW(273) & ChrW(432) & ChrW(7907) & "c ph" & ChrW(432) & ChrW(417) & _
                "ng án " & S_text
            End If
            er = er + 1
        End If
        Set myRange = Nothing
        Next i
        If S_error(1) = "" And ktGood Then
            ActiveDocument.Bookmarks("s1").Delete
            ActiveDocument.Bookmarks("s2").Delete
            Selection.EndKey Unit:=wdStory, Extend:=wdMove
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.TypeBackspace
            Selection.TypeBackspace
            '''''''''''
            Select Case ktlop
            Case 13
                ActiveDocument.SaveAs2 FileName:=S_Drive & "S_Bank&Test\S_Data\Other\" & in_Name & " [" & C - 1 & "]" & NhomTL & ".dat", _
                FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="159", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
            Case 12
                ActiveDocument.SaveAs2 FileName:=S_Drive & "S_Bank&Test\S_Data\Lop 12\" & in_Name & " [" & C - 1 & "]" & NhomTL & ".dat", _
                FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="159", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
            Case 11
                ActiveDocument.SaveAs2 FileName:=S_Drive & "S_Bank&Test\S_Data\Lop 11\" & in_Name & " [" & C - 1 & "]" & NhomTL & ".dat", _
                FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="159", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
            Case 10
                ActiveDocument.SaveAs2 FileName:=S_Drive & "S_Bank&Test\S_Data\Lop 10\" & in_Name & " [" & C - 1 & "]" & NhomTL & ".dat", _
                FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="159", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
            End Select
            ten_file_nguon = in_Name & " [" & C - 1 & "]" & NhomTL
            Kill (S_Drive & "S_Bank&Test\S_Data\tmp.docx")
            '''''''''''
            ActiveDocument.Close
            Load S_mf
            S_mf.mf_t2 = ten_file_nguon
            ktMark = True
            Unload S_inf
            title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
            msg = ChrW(272) & "ã nh" & ChrW(7853) & "p câu h" & ChrW(7887) & "i thành công. N" & ChrW(7871) & "u mu" & ChrW(7889) & "n làm " _
         & ChrW(273) & ChrW(7873) & " ngay bây gi" & ChrW(7901) & " b" & ChrW( _
        7841) & "n ch" & ChrW(7885) & "n ""MixTest"" trên thanh công c" & ChrW(7909) & "."
            Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
        Else
            If ktMix Then
            Selection.EndKey Unit:=wdStory, Extend:=wdMove
            Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
            Selection.TypeBackspace
            Load S_ErrorF
            S_ErrorF.ListBox1.Clear
            For i = 1 To UBound(S_error)
            S_ErrorF.ListBox1.AddItem S_error(i)
            Next i
            '''
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = "  "
                .Replacement.text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            ''''
            S_ErrorF.Show
            End If
            ktMark = False
        End If
    'Dim msg As String
   
Exit Sub
S_Quit:
        Selection.HomeKey Unit:=wdLine
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Range.HighlightColorIndex = wdYellow
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="loi"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        Set myRange = Nothing
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        Selection.TypeBackspace
        Selection.GoTo what:=wdGoToBookmark, Name:="loi"
        'Dim msg As String
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = ChrW(272) & "ã x" & ChrW(7843) & "y ra l" & ChrW( _
        7895) & "i " & ChrW(7903) & " câu có tô vàng." & Chr(13) & "B" & ChrW(7841) & "n nên s" _
         & ChrW(7917) & " d" & ChrW(7909) & "ng ch" & ChrW(7913) & "c n" & ChrW( _
        259) & "ng chu" & ChrW(7849) & "n hóa d" & ChrW(7919) & " li" _
         & ChrW(7879) & "u tr" & ChrW(432) & ChrW(7899) & "c khi nh" & ChrW(7853) _
         & "p và " & ChrW(273) & ChrW(7885) & "c k" & ChrW(7929) & " h" & ChrW( _
        432) & ChrW(7899) & "ng d" & ChrW(7851) & "n cách " & ChrW(273) & ChrW( _
        7883) & "nh d" & ChrW(7841) & "ng các câu h" & ChrW(7887) & "i."
        Application.Assistant.DoAlert title2, msg, 0, 3, 0, 0, 0
        ktMark = False
        
End Sub
Sub Luu_dap_an()
'On Error GoTo S_Quit
Dim cot, dong, si, sj, sn As Integer
If S_Ans.OptionButton1 = False And S_Ans.OptionButton2 = False And S_Ans.OptionButton3 = False Then
    MsgBox "Chua chon dang Table"
    Exit Sub
End If

cot = Selection.Tables(1).Columns.Count
dong = Selection.Tables(1).Rows.Count
If S_Ans.OptionButton1 Then
    sn = cot * dong / 2
    ReDim d_a1(sn) As String
    For sj = 1 To cot / 2
    For si = 1 To dong
    Selection.Tables(1).Cell(si, 2 * sj).Select
    d_a1(si + sj * dong - dong) = Left(Selection.text, 1)
    
    Next si
    Next sj
End If
If S_Ans.OptionButton2 Then
    sn = cot * dong / 2
    ReDim d_a1(sn) As String
    For sj = 1 To dong / 2
    For si = 1 To cot
    Selection.Tables(1).Cell(2 * sj, si).Select
    d_a1(sj * cot + si - cot) = Left(Selection.text, 1)
    
    Next si
    Next sj
End If
If S_Ans.OptionButton3 Then
    sn = cot * dong
    ReDim d_a1(sn) As String
    For sj = 1 To dong
    For si = 1 To cot
    Selection.Tables(1).Cell(sj, si).Select
    d_a1(sj * cot + si - cot) = Left(Right(Selection.text, 3), 1)
    
    Next si
    Next sj
End If
'Dim T_ans As String
'T_ans = ""
S_Ans.ListBox1.Clear
For si = 1 To sn
    If d_a1(si) = "A" Or d_a1(si) = "B" Or d_a1(si) = "C" Or d_a1(si) = "D" Or _
    d_a1(si) = "a" Or d_a1(si) = "b" Or d_a1(si) = "c" Or d_a1(si) = "d" _
    Then S_Ans.ListBox1.AddItem si & ": " & d_a1(si)
    'T_ans = T_ans & d_a1(si)
Next si
    'Dim MyDataObj As New DataObject
    'MyDataObj.SetText T_ans
    'MyDataObj.PutInClipboard
    'MyDataObj.GetFromClipboard
Exit Sub
S_Quit:
MsgBox "Chua chon bang dap an"
End Sub
Sub S_Chamdiem(ByVal control As Office.IRibbonControl)
    On Error GoTo S_Quit
    Call CheckDrive
    Dim S_data As New Word.Document
    Dim cot, dong, si, sj, tongsocau, sn As Integer
    cot = Selection.Tables(1).Columns.Count
    dong = Selection.Tables(1).Rows.Count
    Dim www As New Word.Application
    tongsocau = 0
    sn = cot * dong
    ReDim d_a1(sn) As String
    For sj = 1 To dong
    For si = 1 To cot
    Selection.Tables(1).Cell(sj, si).Select
    d_a1(sj * cot + si - cot) = Left(Right(Selection.text, 3), 1)
    If d_a1(sj * cot + si - cot) = "A" Or d_a1(sj * cot + si - cot) = "B" Or _
    d_a1(sj * cot + si - cot) = "C" Or d_a1(sj * cot + si - cot) = "D" Then
    tongsocau = tongsocau + 1
    End If
    Next si
    Next sj
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\Phieu_cham_bai.docx", ConfirmConversions:=False, ReadOnly _
                :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
                :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
                , Format:=wdOpenFormatAuto, XMLTransform:=""
    Selection.WholeStory
    Selection.Delete
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=Int((tongsocau - 1) / 5) + 1, NumColumns:=10, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Set S_data = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\Mark_Printer.docx")
    dong = Int((tongsocau - 1) / 5) + 1
    For sj = 2 To 10 Step 2
    For si = 1 To dong
    Selection.Tables(1).Cell(si, sj).Select
    Select Case d_a1(si + sj * dong / 2 - dong)
    Case "A"
    www.Selection.Tables(1).Cell(1, 2).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    www.Selection.Copy
    Selection.Paste
    Case "B"
    www.Selection.Tables(1).Cell(1, 4).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    www.Selection.Copy
    Selection.Paste
    Case "C"
    www.Selection.Tables(1).Cell(1, 6).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    www.Selection.Copy
    Selection.Paste
    Case "D"
    www.Selection.Tables(1).Cell(1, 8).Select
    www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    www.Selection.Copy
    Selection.Paste
    End Select
    If (si + sj * dong / 2 - dong) < tongsocau Then _
    Selection.MoveRight Unit:=wdCell, Count:=2
    Next si
    Next sj
    S_data.Close
    www.Quit
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=21, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=21, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=21, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(7).SetWidth ColumnWidth:=21, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(9).SetWidth ColumnWidth:=21, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=76, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=76, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(6).SetWidth ColumnWidth:=76, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(8).SetWidth ColumnWidth:=76, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Columns(10).SetWidth ColumnWidth:=76, RulerStyle:=wdAdjustNone
    Selection.Tables(1).Select
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 3
        .SpaceBeforeAuto = False
        .SpaceAfter = 3
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
Exit Sub
S_Quit:
        Dim title2, msg As String
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n b" & ChrW(7843) & "ng " & ChrW(273) & "áp án"
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
End Sub
Sub S_Chamdiem2(ByVal control As Office.IRibbonControl)
On Error GoTo Thoat
Dim st As String
Dim made() As String
Dim Dapan() As String
Dim title2, msg As String
Dim cot, dong, i, j, sn, si, sj, socot, sodong As Byte
Call S_SerialHDD
If ActiveDocument.Tables.Count = 0 Then
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n b" & ChrW(7843) & "ng " & ChrW(273) & "áp án"
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
socot = (ActiveDocument.Tables.Count) * 2
cot = ActiveDocument.Tables(1).Columns.Count
dong = ActiveDocument.Tables(1).Rows.Count
sn = cot * dong
ReDim made(ActiveDocument.Tables.Count + 1) As String
For i = 1 To ActiveDocument.Tables.Count
    ActiveDocument.Tables(i).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    made(i) = Selection
Next i
ReDim Dapan(ActiveDocument.Tables.Count + 1) As String
For i = 1 To ActiveDocument.Tables.Count
    st = ""
    For sj = 1 To dong
    For si = 1 To cot
    ActiveDocument.Tables(i).Cell(sj, si).Select
    If Left(Right(Selection.text, 3), 1) = "A" Or Left(Right(Selection.text, 3), 1) = "B" _
    Or Left(Right(Selection.text, 3), 1) = "C" Or Left(Right(Selection.text, 3), 1) = "D" _
    Then st = st & Left(Right(Selection.text, 3), 1)
    Next si
    Next sj
    Dapan(i) = st
Next i
    Documents.add
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=socot, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Application.Keyboard (1033)
   
For i = 2 To socot Step 2
    Selection.TypeText text:="Câu"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText text:=made(i / 2)
    Selection.MoveRight Unit:=wdCell
Next i
st = Trim(Dapan(1))
For i = 1 To Len(st)
    For j = 1 To socot / 2
    Selection.TypeText text:=i
    Selection.Shading.BackgroundPatternColor = wdColorTurquoise
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText text:=Mid(Dapan(j), i, 1)
    Selection.MoveRight Unit:=wdCell
    Next j
Next i
    ActiveDocument.Tables(1).Select
    Call FontFormat
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ActiveDocument.Tables(1).Rows(ActiveDocument.Tables(1).Rows.Count).Delete
    ActiveDocument.Tables(1).Rows(1).Select
    Selection.Shading.BackgroundPatternColor = wdColorTurquoise
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
Exit Sub

Thoat:
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = ChrW(272) & ChrW(7883) & "nh d" & ChrW(7841) & "ng " & ChrW(273) & "áp án không " & ChrW(273) & "úng"
        Application.Assistant.DoAlert title2, msg, 0, 4, 0, 0, 0
End Sub
Sub S_Chamdiem3(ByVal control As Office.IRibbonControl)
    S_Pages.Show
End Sub
Sub S_MakeAns()
        Dim C As Integer
        Call RemoveMarks
        Selection.WholeStory
        Selection.Range.HighlightColorIndex = wdNoHighlight
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Selection.TypeParagraph
        Selection.TypeText text:="Please wait.........B&T Program"
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        C = 1
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "([^13^32^9])([AaBbCcDd])([.\)])"
            .Replacement.text = "\1\2\3" & " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
        With Selection.Find
            .text = "(\[\<)([Bb])([Rr])(\>\])(?)"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Do While Selection.Find.Execute = True
            Selection.Collapse Direction:=wdCollapseEnd
            C = C + 1
            Call ClearBlankAfBreak
            Call Check_Br
            With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
        Loop
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        Call ClearBlankAfBreak
        Call Check_Br
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c1q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        
        Dim i As Integer
        For i = 1 To C - 1
           
            Selection.Find.ClearFormatting
            Dim myRange As Range
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Select
        Select Case d_a1(i)
        Case "A"
            myRange.Find.Execute FindText:="([^13^32^9])([A])([.:\)])", _
            MatchWildcards:=True
        Case "B"
            myRange.Find.Execute FindText:="([^13^32^9])([B])([.:\)])", _
            MatchWildcards:=True
        Case "C"
            myRange.Find.Execute FindText:="([^13^32^9])([C])([.:\)])", _
            MatchWildcards:=True
        Case "D"
            myRange.Find.Execute FindText:="([^13^32^9])([D])([.:\)])", _
            MatchWildcards:=True
        End Select
        If myRange.Find.Found = True Then
            myRange.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            With Selection.Font
                .Name = "Times New Roman"
                .Size = 12
                .Bold = True
                .Underline = wdUnderlineSingle
                .Color = &HFF&
            End With
        End If
        Next i
        
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        Selection.TypeBackspace
        Selection.TypeBackspace
End Sub
Sub S_ExportAns()
        On Error GoTo s
        Dim C As Integer
        Call RemoveMarks
        Selection.WholeStory
        'Selection.Range.HighlightColorIndex = wdNoHighlight
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Selection.TypeParagraph
        Selection.TypeText text:="Please wait.........B&T Program"
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        C = 1
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "([^13^32^9])([AaBbCcDd])([.\)])"
            .Replacement.text = "\1\2\3" & " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
        Selection.Find.Highlight = True
        With Selection.Find
            .text = "([ABCD].)"
            .Replacement.text = "\1"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .Format = True
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
        With Selection.Find
            .text = "[<br>]"
            .Replacement.text = "[<Br>]$"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
        With Selection.Find
            .text = "[<Br>]$"
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
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.TypeBackspace
            C = C + 1
            With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
            
        Loop
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c1q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
        End With
        Dim i As Integer
        Dim myRange As Range
        Dim d_a, ktd_a, ktd_b, ktd_c, ktd_d As Byte
        Dim choiceA, choiceB, choiceC, choiceD, txt As String
            choiceA = "([^13^32^9])([Aa])([.:\)])(*)([^13^32^9])([Bb])([.:\)])"
            choiceB = "([^13^32^9])([Bb])([.:\)])(*)([^13^32^9])([Cc])([.:\)])"
            choiceC = "([^13^32^9])([Cc])([.:\)])(*)([^13^32^9])([Dd])([.:\)])"
            choiceD = "([^13^32^9])([Dd])([.:\)])(*)(\[\<)([Bb])([Rr]\>\])"
       txt = ""
    For i = 1 To C - 1
        d_a = 0
        ktd_a = 0
        ktd_b = 0
        ktd_c = 0
        ktd_d = 0
        Set myRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
        Selection.Find.ClearFormatting
        myRange.Find.Execute FindText:=choiceA, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
                Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                d_a = 1
                ktd_a = ktd_a + 1
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
        Selection.Find.ClearFormatting
        Set myRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
        myRange.Find.Execute FindText:=choiceB, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
                Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                d_a = 2
                ktd_b = ktd_b + 1
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
        Selection.Find.ClearFormatting
        Set myRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
        myRange.Find.Execute FindText:=choiceC, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
                Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                d_a = 3
                ktd_c = ktd_c + 1
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
        Selection.Find.ClearFormatting
        Set myRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Bookmarks("s1").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
        myRange.Find.Execute FindText:=choiceD, MatchWildcards:=True
        If myRange.Find.Found = True Then
             myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-2
            myRange.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            If Selection.Font.Underline = wdUnderlineSingle Or _
                Selection.Font.Underline = wdUnderlineDouble Or Selection.Font.ColorIndex = wdRed Then
                d_a = 4
                ktd_d = ktd_d + 1
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
        End If
        If ktd_a + ktd_b + ktd_c + ktd_d = 1 Then
            txt = txt & d_a
        Else
            txt = txt & "0"
        End If
    Next i
    Selection.EndKey Unit:=wdStory
    Selection.MoveUp Unit:=wdLine, Extend:=wdExtend
    Selection.TypeParagraph
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=15, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Application.Keyboard (1033)
    Dim T As String
    For i = 1 To Len(txt)
        Select Case Mid(txt, i, 1)
            Case 1
                T = "A"
            Case 2
                T = "B"
            Case 3
                T = "C"
            Case 4
                T = "D"
            Case Else
                T = "_"
        End Select
        Selection.TypeText text:=i & T
        If i < Len(txt) Then Selection.MoveRight Unit:=wdCell
    Next i
    Selection.Tables(1).Select
    Call FontFormat
Exit Sub
s:

 Selection.HomeKey Unit:=wdLine
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Range.HighlightColorIndex = wdYellow
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="loi"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        Set myRange = Nothing
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        Selection.TypeBackspace
        Selection.GoTo what:=wdGoToBookmark, Name:="loi"
        
        Dim msg, title2 As String
        title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = ChrW(272) & "ã x" & ChrW(7843) & "y ra l" & ChrW( _
        7895) & "i " & ChrW(7903) & " câu có tô vàng." & Chr(13) & "B" & ChrW(7841) & "n nên s" _
         & ChrW(7917) & " d" & ChrW(7909) & "ng ch" & ChrW(7913) & "c n" & ChrW( _
        259) & "ng chu" & ChrW(7849) & "n hóa d" & ChrW(7919) & " li" _
         & ChrW(7879) & "u tr" & ChrW(432) & ChrW(7899) & "c khi nh" & ChrW(7853) _
         & "p và " & ChrW(273) & ChrW(7885) & "c k" & ChrW(7929) & " h" & ChrW( _
        432) & ChrW(7899) & "ng d" & ChrW(7851) & "n cách " & ChrW(273) & ChrW( _
        7883) & "nh d" & ChrW(7841) & "ng các câu h" & ChrW(7887) & "i."
        Application.Assistant.DoAlert title2, msg, 0, 3, 0, 0, 0
End Sub
Private Sub in_dapan(ByRef md As String)
    Dim T As String
    Dim idx As Integer
    Selection.TypeParagraph
    Selection.TypeText text:="M" & ChrW(227) & " " & ChrW(273) & ChrW(7873) & " [" & md & "]"
    Selection.TypeParagraph
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=15, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Application.Keyboard (1033)
    For idx = 1 To socau
        Select Case dapanmoi(idx)
            Case 1
                T = "A"
            Case 2
                T = "B"
            Case 3
                T = "C"
            Case 4
                T = "D"
            Case Else
                T = "_"
        End Select
        Selection.TypeText text:=idx + tucau - 1 & T
        If idx < socau Then Selection.MoveRight Unit:=wdCell
    Next idx
    If S_mf.ComboBox4 = "Default" Then
    Selection.WholeStory
    Else
    Selection.Tables(1).Select
    End If
    Call FontFormat
    'Selection.HomeKey unit:=wdStory
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdRed
    With Selection.Find
        .text = "(\[)(*)(\])"
        .Replacement.text = "\1\2\3"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.EndKey Unit:=wdLine
End Sub

Sub S_FormatHeader1(ByVal control As Office.IRibbonControl)
Call CheckDrive
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Header_1.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub
Sub S_FormatHeader2(ByVal control As Office.IRibbonControl)
Call CheckDrive
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Header_2.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub
Sub S_FormatHeader3(ByVal control As Office.IRibbonControl)
Call CheckDrive
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Header_3.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub
Sub S_FormatHeader4(ByVal control As Office.IRibbonControl)
Call CheckDrive
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Header_4.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub
Sub S_FormatHeader5(ByVal control As Office.IRibbonControl)
Call CheckDrive
    Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Header_5.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub
Sub S_FormatFooter1(ByVal control As Office.IRibbonControl)
Call CheckDrive
Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Footer_1.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub
Sub S_FormatFooter2(ByVal control As Office.IRibbonControl)
Documents.Open FileName:=S_Drive & "S_Bank&Test\S_Templates\default_Footer_2.docx", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""
End Sub

Sub Help(ByVal control As Office.IRibbonControl)
Dim path As String
path = "C:\Program Files\"
If FExists(S_Drive & "S_Bank&Test\S_Templates\Help.doc") Then
path = S_Drive & ""
Else
If DirExists("C:\Program Files (x86)\") Then path = "C:\Program Files (x86)\"
End If
Documents.Open FileName:=path & "S_Bank&Test\S_Templates\Help.doc", ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:=""

End Sub
Sub ClearBlankBf()
    Dim Tb As String
    demkt = 0
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Tb = Selection
    Do While Tb = Chr(13) Or Tb = Chr(9) Or Tb = " " Or Tb = Chr(11)
        demkt = demkt + 1
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Tb = Selection
        If demkt > 100 Then Exit Do
    Loop
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
End Sub
Sub ClearBlankAfABCD()
    Dim Tb As String
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Tb = Selection
    Do While Tb = " " Or Tb = Chr(13) Or Tb = Chr(9) Or Tb = Chr(11)
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    'Selection.TypeBackspace
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Tb = Selection
    Loop
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
    'Selection.TypeText text:=" "
End Sub
Sub ClearBlankAfBreak()
    Dim Tb As String
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Tb = Selection
    Do While Tb = " " Or Tb = Chr(13) Or Tb = Chr(9) Or Tb = Chr(11)
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Tb = Selection
    Loop
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
End Sub
Private Sub Check_Br()
    Dim Tb, tb2 As String
    demkt = 0
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Tb = Selection
    If Tb = "Câu" Then
        demkt = 4
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        tb2 = Selection
        Do
            demkt = demkt + 1
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            tb2 = Selection
        Loop Until tb2 = ":" Or tb2 = "." Or tb2 = ")" Or demkt = 10
        If demkt = 10 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=6, Extend:=wdExtend
            Selection.Range.HighlightColorIndex = wdYellow
        End If
        Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    Else
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
    End If
    If demkt = 10 Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=demkt + 4, Extend:=wdMove
    End If
End Sub
Sub MainForm_load(ByVal control As Office.IRibbonControl)
    S_mf.Show
End Sub
Sub InputForm_load(ByVal control As Office.IRibbonControl)
    S_inf.Show
End Sub
Sub ImBankForm_load(ByVal control As Office.IRibbonControl)
    S_ImBank.Show
End Sub
Sub ReBankForm_load(ByVal control As Office.IRibbonControl)
    S_ReBank.Show
End Sub
Sub CreBankForm_load(ByVal control As Office.IRibbonControl)
    S_PPCT.Show
End Sub
Sub S_Matrix(ByVal control As Office.IRibbonControl)
    S_matran.Show
End Sub
Sub S_Lamdapan(ByVal control As Office.IRibbonControl)
    S_Ans.Show
End Sub
Sub S_seri(ByVal control As Office.IRibbonControl)
    S_Serial.Show
End Sub

Sub Standar(ByVal control As Office.IRibbonControl)
    S_Standar.Show
End Sub
Sub BangBT(ByVal control As Office.IRibbonControl)
    S_bbtF.Show
End Sub
Sub Model(ByVal control As Office.IRibbonControl)
    S_ModelF.Show
End Sub

