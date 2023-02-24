VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_ImBank 
   Caption         =   "Import-Split"
   ClientHeight    =   8010
   ClientLeft      =   135
   ClientTop       =   2760
   ClientWidth     =   3825
   OleObjectBlob   =   "S_ImBank.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "S_ImBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub butChuanhoa_Click()
    Call ChuanDATA
End Sub
Private Sub Save_File_CB()
    Call CheckDrive
    Dim www As New Word.Application
    Dim myDoc As Document
    Dim myRange As Range
    Dim tam() As String
    Dim S_path, S_pathROOT, tenF As String
    '''''''''
    If DirExists(S_Drive & "S_Bank&Test\S_Split") = False Then
        MkDir (S_Drive & "S_Bank&Test\S_Split")
    End If
    If DirExists(S_Drive & "S_Bank&Test\S_Split\Tach theo Chuong_Bai") = False Then
        MkDir (S_Drive & "S_Bank&Test\S_Split\Tach theo Chuong_Bai")
    End If
    '''''''''
    S_pathROOT = S_Drive & "S_Bank&Test\S_Split\Tach theo Chuong_Bai\"
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeText text:="NHOM CAU DANG "
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "(NHOM CAU DANG )(\[*.[abcd]\])(*)(NHOM CAU DANG)"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
   
    Do While Selection.Find.Execute = True
        Set myRange = Selection.Range
        myRange.MoveStart Unit:=wdCharacter, Count:=28
        myRange.MoveEnd Unit:=wdCharacter, Count:=-14
        myRange.Select
        With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="Smark"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
        End With
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.EndKey Unit:=wdLine
        Selection.MoveLeft Unit:=wdCharacter, Count:=13, Extend:=wdExtend
        tam = Split(Trim(Selection), ".")
        S_path = Mid(tam(0), 2, 4) & "-Chuong " & Right(tam(1), 1)
    'MsgBox S_path
        If DirExists(S_pathROOT & S_path) = False Then
            MkDir (S_pathROOT & S_path)
        End If
        S_path = S_path & "\Bai " & tam(2)
        Select Case Left(tam(3), 1)
        Case "a"
            tenF = "_Muc do 1.docx"
        Case "b"
            tenF = "_Muc do 2.docx"
        Case "c"
            tenF = "_Muc do 3.docx"
        Case "d"
            tenF = "_Muc do 4.docx"
        End Select
        If FExists(S_pathROOT & S_path & tenF) = False Then
            Set myDoc = www.Documents.add
            Call S_PageSetup
        Else
            Set myDoc = www.Documents.Open(S_pathROOT & S_path & tenF, PasswordDocument:="")
            www.Selection.EndKey Unit:=wdStory
            www.Selection.TypeParagraph
        End If
        Selection.GoTo what:=wdGoToBookmark, Name:="Smark"
        Selection.Copy
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        www.Selection.Paste
        myDoc.SaveAs2 FileName:=S_pathROOT & S_path & tenF, FileFormat:= _
                wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
        myDoc.Close
    Loop
    www.Quit
End Sub
Private Sub Save_File_CD()
    Call CheckDrive
    Dim www As New Word.Application
    Dim myDoc As Document
    Dim myRange As Range
    Dim tam() As String
    Dim S_path, S_pathROOT, tenF As String
    '''''''''
    If DirExists(S_Drive & "S_Bank&Test\S_Split") = False Then
        MkDir (S_Drive & "S_Bank&Test\S_Split")
    End If
    If DirExists(S_Drive & "S_Bank&Test\S_Split\Tach theo Chuyen de") = False Then
        MkDir (S_Drive & "S_Bank&Test\S_Split\Tach theo Chuyen de")
    End If
    '''''''''
    If tachChuongBai Then
        S_pathROOT = S_Drive & "S_Bank&Test\S_Split\Tach theo Chuong_Bai\"
    Else
        S_pathROOT = S_Drive & "S_Bank&Test\S_Split\Tach theo Chuyen de\"
    End If
    Selection.EndKey Unit:=wdStory
    Selection.TypeText text:="NHOM CAU DANG "
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "(NHOM CAU DANG )(\[*[abcd]\])(*)(NHOM CAU DANG)"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
   
    Do While Selection.Find.Execute = True
        Set myRange = Selection.Range
        myRange.MoveStart Unit:=wdCharacter, Count:=32
        myRange.MoveEnd Unit:=wdCharacter, Count:=-14
        myRange.Select
        With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="Smark"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
        End With
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.EndKey Unit:=wdLine
        Selection.MoveLeft Unit:=wdCharacter, Count:=17, Extend:=wdExtend
        tam = Split(Trim(Selection), ".")
        S_path = Mid(tam(0), 2, 4) & "-Chuong " & Right(tam(1), 1)
        If DirExists(S_pathROOT & S_path) = False Then
            MkDir (S_pathROOT & S_path)
        End If
        S_path = S_path & "\CD " & tam(2)
        If DirExists(S_pathROOT & S_path) = False Then
            MkDir (S_pathROOT & S_path)
        End If

        tenF = Right(tam(3), 2)

        Select Case Left(tam(4), 1)
        Case "a"
            tenF = tenF & "_Muc do 1.docx"
        Case "b"
            tenF = tenF & "_Muc do 2.docx"
        Case "c"
            tenF = tenF & "_Muc do 3.docx"
        Case "d"
            tenF = tenF & "_Muc do 4.docx"
        End Select
        If FExists(S_pathROOT & S_path & "\Dang " & tenF) = False Then
            Set myDoc = www.Documents.add
            Call S_PageSetup
        Else
            Set myDoc = www.Documents.Open(S_pathROOT & S_path & "\Dang " & tenF, PasswordDocument:="")
            www.Selection.EndKey Unit:=wdStory
            www.Selection.TypeParagraph
        End If
        Selection.GoTo what:=wdGoToBookmark, Name:="Smark"
        Selection.Copy
        www.Selection.Paste

        www.ActiveDocument.SaveAs2 FileName:=S_pathROOT & S_path & "\Dang " & tenF, FileFormat:= _
                wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
                :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
        myDoc.Close
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop
    www.Quit
End Sub
Private Sub butXuatfile_Click()
    If ktBanQuyen = False And S_ImBank.tachChuyenDe Then Call S_SerialHDD
    If ktBanQuyen = False Then
        S_NoteRig.Show
        Exit Sub
    End If
    
        Dim ktTB As Byte
        If boxSave Then
        ktTB = Application.Assistant.DoAlert("Th" & ChrW(244) & "ng b" & ChrW(225) & "o", _
                "B" & ChrW(7841) & "n hãy xem k" & ChrW(7929) & _
                " m" & ChrW(7897) & "t l" & ChrW(7847) & "n n" & ChrW(7919) & _
                "a các thông tin ""Mã câu h" & ChrW(7887) & _
                "i"" hay ""Tách theo ch" & ChrW(432) & ChrW(417) & "ng bài hay chuyên " _
                & ChrW(273) & ChrW(7873) & """ vì khi " & ChrW(273) & "ã tách và xu" & ChrW(7845) _
                & "t thành t" & ChrW(7915) & "ng m" & ChrW(7913) & "c " & ChrW(273) & _
                ChrW(7897) & " r" & ChrW(7891) & "i thì r" & ChrW(7845) & "t khó s" & _
                ChrW(7917) & "a ch" & ChrW(7919) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u." _
                , 1, 3, 0, 0, 0)
        If ktTB = 2 Then Exit Sub
        End If
        Dim C, i As Integer
        Dim Title, msg As String
        Dim DS_10 As String
        Dim HH_10 As String
        Dim DS_11 As String
        Dim HH_11 As String
        Dim DS_12 As String
        Dim HH_12 As String
        Dim ID, findTxt As String
        On Error Resume Next
        Dim docThis, docThat As Document
        Call RemoveMarks
        S_ImBank.Hide
        Set docThis = ActiveDocument
        C = 0
        DS_10 = "STRART"
        HH_10 = "STRART"
        DS_11 = "STRART"
        HH_11 = "STRART"
        DS_12 = "STRART"
        HH_12 = "STRART"
        Selection.HomeKey Unit:=wdStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        If tachChuyenDe Then
            findTxt = "(\[)(??1)([012])(*)(.D??.)([abcd])(\])"
        Else
            findTxt = "(\[)(??1)([012])(*)(.)([abcd])(\])"
        End If
        With Selection.Find
            .text = findTxt
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Do While Selection.Find.Execute = True
            'Selection.Select
            If tachChuyenDe Then
                ID = Trim(Selection)
            Else
                ID = Left(Trim(Selection), 10) & Right(Trim(Selection), 3)
            End If
            Selection.HomeKey Unit:=wdLine
            C = C + 1
            With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
            Selection.EndKey Unit:=wdLine
            
            Select Case Mid(ID, 2, 4)
            Case "DS10"
                If InStr(DS_10, ID) = 0 Then DS_10 = DS_10 & "," & ID
            Case "HH10"
                If InStr(HH_10, ID) = 0 Then HH_10 = HH_10 & "," & ID
            Case "DS11"
                If InStr(DS_11, ID) = 0 Then DS_11 = DS_11 & "," & ID
            Case "HH11"
                If InStr(HH_11, ID) = 0 Then HH_11 = HH_11 & "," & ID
            Case "DS12"
                If InStr(DS_12, ID) = 0 Then DS_12 = DS_12 & "," & ID
            Case "HH12"
                If InStr(HH_12, ID) = 0 Then HH_12 = HH_12 & "," & ID
            End Select
        Loop
        If C = 0 Then
            Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", _
            "Không tìm th" & ChrW(7845) & "y mã câu h" & ChrW(7887) & "i phù h" & ChrW(7907) & "p. Ki" & _
            ChrW(7875) & "m tra l" & ChrW(7841) & "i ""Tách theo ch" & ChrW(432) & ChrW(417) & _
            "ng bài"" hay ""Tách theo chuyên " & ChrW(273) & ChrW(7873) & """" & "." _
            , 0, 4, 0, 0, 0
            Exit Sub
        End If
        'MsgBox DS_12
    'Exit Sub
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C + 1 & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
    'Exit Sub
        Dim myRange As Range
        Dim tam() As String
        Dim tam_MARK() As String
        Set docThat = Documents.add
        Call S_PageSetup
        Dim i1, i2 As Integer
        Dim tg As String
        If tachChuongBai Then
            tam_MARK = Split(DS_10, ",")
  
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
        
            tam_MARK = Split(HH_10, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(DS_11, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(HH_11, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(DS_12, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(HH_12, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
        Else
            tam_MARK = Split(DS_10, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & Mid(tam_MARK(i), 12, 3) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(HH_10, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & Mid(tam_MARK(i), 12, 3) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(DS_11, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & Mid(tam_MARK(i), 12, 3) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(HH_11, ",")
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & Mid(tam_MARK(i), 12, 3) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(DS_12, ",")
            
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & Mid(tam_MARK(i), 12, 3) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
            tam_MARK = Split(HH_12, ",")
            
            If UBound(tam_MARK) > 0 Then
                For i1 = 1 To UBound(tam_MARK) - 1
                For i2 = i1 + 1 To UBound(tam_MARK)
                    If StrComp(tam_MARK(i1), tam_MARK(i2)) = 1 Then
                        tg = tam_MARK(i1)
                        tam_MARK(i1) = tam_MARK(i2)
                        tam_MARK(i2) = tg
                    End If
                Next i2
                Next i1
                For i = 1 To UBound(tam_MARK)
                    Selection.TypeText text:="NHOM CAU DANG " & tam_MARK(i)
                    With ActiveDocument.Bookmarks
                        .add Range:=Selection.Range, Name:=Mid(tam_MARK(i), 2, 4) & "C" & Mid(tam_MARK(i), 8, 1) & _
                                    "B" & Mid(tam_MARK(i), 10, 1) & Mid(tam_MARK(i), 12, 3) & "MD" & Mid(tam_MARK(i), Len(tam_MARK(i)) - 1, 1)
                        .DefaultSorting = wdSortByName
                        .ShowHidden = True
                    End With
                    Selection.TypeParagraph
                Next i
            End If
        End If
'Exit Sub
        
    For i = 1 To C
        Set myRange = docThis.Range( _
            Start:=docThis.Bookmarks("c" & i & "q").Range.Start, _
            End:=docThis.Bookmarks("c" & i + 1 & "q").Range.End)
        myRange.Select
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="cautam"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        myRange.Find.Execute FindText:="(\[)(??1)([012])(*)(.)([abcd])(\])", MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.Select
            ID = Trim(Selection)
            tam = Split(ID, ".")
            Selection.GoTo what:=wdGoToBookmark, Name:="cautam"
            Selection.Copy
            docThat.Activate
            
            If tachChuongBai Then
                If docThat.Bookmarks.Exists(Mid(tam(0), 2, 4) & tam(1) & "B" & tam(2) & "MD" & Left(tam(4), 1)) Then
                    Selection.GoTo what:=wdGoToBookmark, Name:=Mid(tam(0), 2, 4) & tam(1) & "B" & tam(2) & "MD" & Left(tam(4), 1)
                Else
                    'MsgBox "Cau " & i & ": Chuong hoac, Bai hoac Muc do khong dung"
                    'Exit Sub
                End If
                Selection.TypeParagraph
                Selection.Paste
            Else
                If docThat.Bookmarks.Exists(Mid(tam(0), 2, 4) & tam(1) & "B" & tam(2) & tam(3) & "MD" & Left(tam(4), 1)) Then
                    Selection.GoTo what:=wdGoToBookmark, Name:=Mid(tam(0), 2, 4) & tam(1) & "B" & tam(2) & tam(3) & "MD" & Left(tam(4), 1)
                Else
                    'MsgBox "Cau " & i & ": Chuong hoac, Bai hoac Muc do khong dung"
                    'Exit Sub
                End If
                Selection.TypeParagraph
                Selection.Paste
            End If
            
        End If
    Next i
    docThat.Activate
    'Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
    
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    'Selection.Find.Replacement.Highlight
    With Selection.Find.Replacement.Font
        .Underline = wdUnderlineDouble
        .Color = wdColorRed
        .Bold = True
    End With
    'If tachChuyenDe Then
        With Selection.Find
            .text = "(NHOM CAU DANG )(\[*.[abcd]\])"
            .Replacement.text = "\1\2"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
   
    Selection.HomeKey Unit:=wdStory
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
    'docThis.Close (False)
    Selection.HomeKey Unit:=wdStory
    If boxSave Then
        If tachChuongBai Then
            Call Save_File_CB
        Else
            Call Save_File_CD
        End If
    End If
    
    MsgBox "Done"
End Sub
        
Private Sub butNhapcauhoi_Click()
    Dim ktTB As Byte
    ktTB = Application.Assistant.DoAlert("Th" & ChrW(244) & "ng b" & ChrW(225) & "o", _
            "B" & ChrW(7841) & "n hãy xem k" & ChrW(7929) & " m" & ChrW(7897) & "t l" & ChrW( _
            7847) & "n n" & ChrW(7919) & "a các thông tin ""Mã câu h" & ChrW(7887) & _
            "i"" hay ""Nh" & ChrW(7853) & "p vào ngân hàng theo ch" & ChrW(432) & ChrW(417) & _
            "ng bài hay chuyên " & ChrW(273) & ChrW(7873) & """ vì khi d" & ChrW(7919 _
            ) & " li" & ChrW(7879) & "u " & ChrW(273) & "ã nh" & ChrW(7853) & "p vào ngân hàng r" & ChrW(7891) & "i thì r" & ChrW( _
            7845) & "t khó s" & ChrW(7917) & "a ch" & ChrW(7919) & "a. " & _
            "B" & ChrW(7841) & "n c" & ChrW(7847) & _
            "n tách nh" & ChrW(7887) & " file d" & ChrW(7919) & " li" & ChrW(7879) & _
            "u c" & ChrW(7847) & "n nh" & ChrW(7853) & "p vào (kho" & ChrW(7843) & _
            "ng 20 trang A4 ho" & ChrW(7863) & "c 50 câu) " & ChrW(273) & _
             ChrW(7875) & " tránh l" & ChrW(7895) & "i khi nh" & ChrW(7853) & "p. " & _
            "B" & ChrW(7841) & "n có ti" & ChrW(7871) & "p t" & ChrW(7909) & "c nh" & ChrW(7853) & "p không ?" _
            , 1, 3, 0, 0, 0)
    If ktTB = 2 Then Exit Sub
    
    If Theo_Bai = False And Theo_CD = False Then
        MsgBox "Chua chon theo bai hay theo chuyen de"
        Exit Sub
    End If
    If CoHDgiai = False And KhongHDgiai = False Then
        MsgBox "Chua chon co HDG hoac khong HDG"
        Exit Sub
    End If
        Dim i As Integer
        Call CheckDrive
        Dim C As Integer
        Dim d_a As String
        Dim www As New Word.Application
        Dim S_Bank As Word.Document
        Dim tam() As String
        Dim Tnumber As Byte
        Dim Title, msg As String
        Dim ktMsg As Byte
        On Error GoTo S_Quit
        Call RemoveMarks
    S_ImBank.Hide
        'ActiveDocument.SaveAs2 FileName:="tmp.docx", FileFormat:= _
                wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
                :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
        C = 0
        Selection.HomeKey Unit:=wdStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        If Theo_Bai Then
            With Selection.Find
                .text = "(\[)(??1)([012])(*)([LB])(T.)([abcd])(\])"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
        End If
        If Theo_CD Then
            With Selection.Find
                .text = "(\[)(??1)([012])(*)(??)(.)([abcd])(\])"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
        End If
        Do While Selection.Find.Execute = True
            Selection.Collapse Direction:=wdCollapseStart
            C = C + 1
            With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
            End With
            Selection.EndKey Unit:=wdLine
        Loop
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        With ActiveDocument.Bookmarks
            .add Range:=Selection.Range, Name:="c" & C + 1 & "q"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        
        ''''
        Dim ktHDG As Boolean
        Dim ktA, ktB, ktC, ktD, ktd_a As Integer
        Dim choiceA, choiceB, choiceC, choiceD, choiceID, choiceHDG, ID As String
        Dim myRange As Range
            If Theo_Bai Then
                choiceID = "(\[)(??1)([012])(*)([LB])(T.)([abcd])(\])"
            Else
                choiceID = "(\[)(??1)([012])(*)(??)(.)([abcd])(\])"
            End If
            choiceA = "([^13^32^9])(A)([.:\)])(*)([^13^32^9])(B)([.:\)])"
            choiceB = "([^13^32^9])(B)([.:\)])(*)([^13^32^9])(C)([.:\)])"
            choiceC = "([^13^32^9])(C)([.:\)])(*)([^13^32^9])(D)([.:\)])"
            choiceD = "([^13^32^9])(D)([.:\)])"
        'www.Application.ScreenUpdating = False
        For i = 1 To C
            ktA = 0
            ktB = 0
            ktC = 0
            ktD = 0
            ktd_a = 0
            
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            'Danh dau ID
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            myRange.Find.Execute FindText:=choiceID, MatchWildcards:=True
        If myRange.Find.Found = True Then
            myRange.Select
            ID = Selection
            
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            Call ClearBlankBf
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "q"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        End If
            'Danh dau phuong an A
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "q").Range.Start, _
            End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
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
                    d_a = "A"
                    ktd_a = ktd_a + 1
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
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "a"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktA = ktA + 1
            
        End If
            'Danh dau phuong an B
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
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
                    d_a = "B"
                    ktd_a = ktd_a + 1
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
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "b"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktB = ktB + 1
        End If
           'Danh dau phuong an C
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
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
                    d_a = "C"
                    ktd_a = ktd_a + 1
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
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "c"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            ktC = ktC + 1
        End If
            'Danh dau phuong an D
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("s2").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            myRange.Find.Execute FindText:=choiceD, MatchWildcards:=True
        If myRange.Find.Found = True Then
                        
            myRange.MoveStart Unit:=wdCharacter, Count:=1
            myRange.MoveEnd Unit:=wdCharacter, Count:=-1
            myRange.Select
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "d"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i + 1 & "q"
            Selection.HomeKey Unit:=wdLine
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
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
                    d_a = "D"
                    ktd_a = ktd_a + 1
            End If
            Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
            Call ClearBlankAfABCD
            Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="c" & i & "d"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        If CoHDgiai = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            With ActiveDocument.Bookmarks
                .add Range:=Selection.Range, Name:="s1"
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
            
            Set myRange = ActiveDocument.Range( _
                Start:=ActiveDocument.Bookmarks("c" & i & "d").Range.Start, _
                End:=ActiveDocument.Bookmarks("c" & i + 1 & "q").Range.End)
            myRange.Select
            
            Set myRange = ActiveDocument.Range( _
                Start:=Selection.Bookmarks(2).Range.Start, _
                End:=Selection.Bookmarks(3).Range.End)
            myRange.Select
            
            With ActiveDocument.Bookmarks
                    .add Range:=Selection.Range, Name:="c" & i & "HD"
                    .DefaultSorting = wdSortByName
                    .ShowHidden = True
            End With
        End If
       'Exit Sub
            ktD = ktD + 1
        End If
        ''''''''''
        If d_a = "" Or ktd_a <> 1 Or ktA <> 1 Or ktB <> 1 Or ktC <> 1 Or ktD <> 1 Then
             GoTo Tiep
        End If
        ''''''''''
        tam = Split(ID, ".")
        
        If Theo_Bai Then
            
            If FExists(S_Drive & "S_Bank&Test\S_Bank\Lop " & Right(tam(0), 2) & "\" & Left(ID, 8) & "].dat") Then
                Set S_Bank = www.Documents.Open(S_Drive & "S_Bank&Test\S_Bank\Lop " & _
                Right(tam(0), 2) & "\" & Left(ID, 8) & "].dat", PasswordDocument:="159")
            Else
                Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
                msg = "C" & ChrW(417) & " s" & ChrW(7903) & " d" & ChrW _
                        (7919) & " li" & ChrW(7879) & "u ch" & ChrW(432) & "a có file " & Left(ID, 8) & "].dat." & Chr(13) _
                        & "B" & ChrW(7887) & " qua câu này r" & ChrW(7891) _
                        & "i nh" & ChrW(7853) & "p ti" & ChrW(7871) & "p?"
                ktMsg = Application.Assistant.DoAlert(Title, msg, 4, 2, 0, 0, 1)
                If ktMsg = 6 Then
                    GoTo Tiep
                Else
                    Exit Sub
                End If
            End If
        Else
            If FExists(S_Drive & "S_Bank&Test\S_Bank\Lop " & Right(tam(0), 2) & "\Chuyen de\" & Mid(ID, 2, 7) & "\" & Left(ID, 14) & "].dat") Then
                Set S_Bank = www.Documents.Open(S_Drive & "S_Bank&Test\S_Bank\Lop " & _
                Right(tam(0), 2) & "\Chuyen de\" & Mid(ID, 2, 7) & "\" & Left(ID, 14) & "].dat", PasswordDocument:="159")
            Else
                Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
                msg = "C" & ChrW(417) & " s" & ChrW(7903) & " d" & ChrW _
                        (7919) & " li" & ChrW(7879) & "u ch" & ChrW(432) & "a có file " & Left(ID, 10) & "].dat." & Chr(13) _
                        & "B" & ChrW(7887) & " qua câu này r" & ChrW(7891) _
                        & "i nh" & ChrW(7853) & "p ti" & ChrW(7871) & "p?"
                ktMsg = Application.Assistant.DoAlert(Title, msg, 4, 2, 0, 0, 1)
                If ktMsg = 6 Then
                    GoTo Tiep
                Else
                    Exit Sub
                End If
            End If
        End If
        If Theo_Bai Then
            Tnumber = 8 * (tam(2) - 1) + 1
            If tam(3) = "BT" Then Tnumber = Tnumber + 4
            Select Case tam(4)
            Case "b]"
            Tnumber = Tnumber + 1
            Case "c]"
            Tnumber = Tnumber + 2
            Case "d]"
            Tnumber = Tnumber + 3
            End Select
        Else
            'Tnumber = 4 * (Right(tam(3), 2) - 1) + 1
            Select Case tam(4)
            Case "a]"
            Tnumber = 1
            Case "b]"
            Tnumber = 2
            Case "c]"
            Tnumber = 3
            Case "d]"
            Tnumber = 4
            End Select
        End If
        If S_Bank.Tables.Count < Tnumber Then
            'S_bank.Close
            GoTo Tiep
        End If
        S_Bank.Tables(Tnumber).Rows(S_Bank.Tables(Tnumber).Rows.Count).Select
        www.Selection.InsertRowsBelow 6
        www.Selection.TypeText text:="Câu " & (Int(S_Bank.Tables(Tnumber).Rows.Count / 6) + 1) - 1 & ":"
        www.Selection.MoveRight Unit:=wdCell
        Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "q"
        Selection.Cut
        www.Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
        www.Selection.MoveRight Unit:=wdCell
        www.Selection.TypeText text:="A."
        www.Selection.MoveRight Unit:=wdCell
        Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "a"
        Selection.Cut
        www.Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
        www.Selection.MoveRight Unit:=wdCell
        www.Selection.TypeText text:="B."
        www.Selection.MoveRight Unit:=wdCell
        Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "b"
        Selection.Cut
        www.Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
        www.Selection.MoveRight Unit:=wdCell
        www.Selection.TypeText text:="C."
        www.Selection.MoveRight Unit:=wdCell
        Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "c"
        Selection.Cut
        www.Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
        www.Selection.MoveRight Unit:=wdCell
        www.Selection.TypeText text:="D."
        www.Selection.MoveRight Unit:=wdCell
        Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "d"
        Selection.Cut
        www.Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
        ktHDG = False
        If CoHDgiai Then
            Selection.GoTo what:=wdGoToBookmark, Name:="c" & i & "HD"
            If Len(Selection) > 1 And CoHDgiai Then
                Selection.Cut
                ktHDG = True
            End If
        End If
        www.Selection.MoveRight Unit:=wdCell
        www.Selection.TypeText text:=ChrW(272) & "áp án:"
        www.Selection.MoveRight Unit:=wdCell
        www.Selection.TypeText text:=d_a
        If ktHDG Then
            www.Selection.Paste
        End If
        If www.Documents.Count > 1 Then
        www.Documents(2).Save
        www.Documents(2).Close (False)
        End If
Tiep:
    Next i
    www.Documents(1).Close (True)
    www.Quit
    MsgBox "Xong!"
    'Set S_bank = Nothing
    'Set www = Nothing
Exit Sub
S_Quit:
    
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "D" & ChrW(7919) & " li" & ChrW(7879) & "u nh" & _
        ChrW(7853) & "p vào b" & ChrW(7883) & " l" & ChrW(7895) & "i."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    If www.Documents.Count > 0 Then www.Documents(1).Close (False)
    www.Quit (False)
    'Set S_bank = Nothing
    'Set www = Nothing
End Sub

Private Sub buttaoID_Click()
    Dim i As Integer
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "([CcBb])([âaà©])([ui])( [0-9]{1,4})([.:\)])"
        .Replacement.text = "$#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    'Selection.Font.ColorIndex = wdBlue
    'Selection.Find.Replacement.ClearFormatting
    Dim txt As String
    txt = "[" & IDmon & IDclass & ".C" & IDchuong & "." & IDbai & "." & IDdang & "." & IDmucdo & "]"
    With Selection.Find
        .text = "$#"
        .MatchWildcards = False
    End With
    i = 1
    Do While Selection.Find.Execute = True
        Selection.TypeText text:="Câu " & i & ": " & txt
        Selection.EndKey Unit:=wdLine, Extend:=wdMove
        i = i + 1
    Loop
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    'Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdBlue
    With Selection.Find
        .text = "(^13)([CcBb])([âaà©])([ui])( [0-9]{1,4})([.:\)])"
        .Replacement.text = "\1\2\3\4\5\6"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    'Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdPink
    With Selection.Find
        .text = "(\[)([DGH])([STH])(*)(.)([abcd])(\])"
        .Replacement.text = "\1\2\3\4\5\6\7"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.ColorIndex = wdPink
    With Selection.Find
        .text = "(\[)([DGH])([STH])(*)(.)([abcd])(\])(^13)"
        .Replacement.text = "\1\2\3\4\5\6\7"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
End Sub

Private Sub butxoaID_Click()
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "(\[)([DGH])([STH])(*)(.)([abcd])(\])"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.HomeKey Unit:=wdStory
End Sub

Private Sub CommandButton1_Click()
    Dim www As New Word.Application
    Dim docGoc As Document
    Dim docShare As Document
    Dim path As String
    Dim selectedFilename() As String
    Dim item, i, j As Integer
    Dim tenTam() As String
    Dim Title, msg As String
    Unload S_ImBank
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
        item = .SelectedItems.Count
        If item = 0 Then
            ReDim selectedFilename(1) As String
            selectedFilename(1) = ""
        Else
            ReDim selectedFilename(item) As String
            For i = 1 To item
                selectedFilename(i) = .SelectedItems(i)
            Next
        End If
    End With
    'MsgBox item
    'MsgBox selectedFilename(2)
    Call CheckDrive
    If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de") = False Then
    MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de")
    End If
    If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de") = False Then
        MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de")
    End If
    If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de") = False Then
        MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de")
    End If
    '''''''''
    If selectedFilename(1) = "" Then
        Exit Sub
    End If
    For j = 1 To item
        tenTam = Split(selectedFilename(j), "\")
        If S_ImBank.Theo_Bai Then
             If FExists(S_Drive & "S_Bank&Test\S_Bank\Lop 10\" & tenTam(UBound(tenTam))) Then
                 path = S_Drive & "S_Bank&Test\S_Bank\Lop 10\"
             ElseIf FExists(S_Drive & "S_Bank&Test\S_Bank\Lop 11\" & tenTam(UBound(tenTam))) Then
                 path = S_Drive & "S_Bank&Test\S_Bank\Lop 11\"
             ElseIf FExists(S_Drive & "S_Bank&Test\S_Bank\Lop 12\" & tenTam(UBound(tenTam))) Then
                 path = S_Drive & "S_Bank&Test\S_Bank\Lop 12\"
             Else
                 Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
                 msg = "Data b" & ChrW(7841) & "n ch" & ChrW(7885) & _
                 "n không tìm th" & ChrW(7845) & "y trong ngân hàng câu h" & ChrW(7887) & _
                 "i nên không th" & ChrW(7875) & " ghép n" & ChrW(7889) & "i."
                 Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 1
                 Exit Sub
             End If
         Else
             If FExists(S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de\" & Mid(tenTam(UBound(tenTam)), 2, 7) & "\" & tenTam(UBound(tenTam))) Then
                 path = S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de\" & Mid(tenTam(UBound(tenTam)), 2, 7) & "\"
             ElseIf FExists(S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de\" & Mid(tenTam(UBound(tenTam)), 2, 7) & "\" & tenTam(UBound(tenTam))) Then
                 path = S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de\" & Mid(tenTam(UBound(tenTam)), 2, 7) & "\"
             ElseIf FExists(S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de\" & Mid(tenTam(UBound(tenTam)), 2, 7) & "\" & tenTam(UBound(tenTam))) Then
                 path = S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de\" & Mid(tenTam(UBound(tenTam)), 2, 7) & "\"
             Else
                 Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
                 msg = "Data b" & ChrW(7841) & "n ch" & ChrW(7885) & _
                 "n không tìm th" & ChrW(7845) & "y trong ngân hàng câu h" & ChrW(7887) & _
                 "i nên không th" & ChrW(7875) & " ghép n" & ChrW(7889) & "i."
                 Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 1
                 GoTo Tieptuc
                 'Exit Sub
             End If
         End If
         Set docGoc = www.Documents.Open(path & tenTam(UBound(tenTam)), PasswordDocument:="159")
         Set docShare = Documents.Open(selectedFilename(j), PasswordDocument:="159")
         Dim myRange As Range
         
         For i = 1 To docShare.Tables.Count
             If docShare.Tables(i).Rows.Count > 6 Then
                 Set myRange = docShare.Range( _
                             Start:=docShare.Tables(i).Cell(2, 1).Range.Start, _
                             End:=docShare.Tables(i).Cell(docShare.Tables(i).Rows.Count, 2).Range.End)
                 myRange.Select
                 Selection.Copy
                 docGoc.Tables(i).Rows(docGoc.Tables(i).Rows.Count).Select
                 www.Selection.InsertRowsBelow docShare.Tables(i).Rows.Count - 1
                 Set myRange = docGoc.Range( _
                             Start:=docGoc.Tables(i).Cell(docGoc.Tables(i).Rows.Count - docShare.Tables(i).Rows.Count + 2, 1).Range.Start, _
                             End:=docGoc.Tables(i).Cell(docGoc.Tables(i).Rows.Count + 2, 2).Range.End)
                 myRange.Select
                 www.Selection.Paste
             End If
         Next i
         docShare.Close
         docGoc.Close (True)
Tieptuc:
    Next j
    www.Quit
    Set docShare = Nothing
    Set docGoc = Nothing
    Set www = Nothing
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "B" & ChrW(7841) & "n " & ChrW(273) & "ã ghép n" & ChrW(7889) & "i thành công."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 1
End Sub

Private Sub CommandButton2_Click()
    S_ImBank.Width = 289
End Sub


Private Sub Label13_Click()
        On Error Resume Next
        Dim tam() As String
        S_ImBank.Width = 204
        Dim Chuong As String
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "(H)(.)(*)(.[0-9]{1,4}.)([abcd1234])([.: ^13])"
            .Replacement.text = "HH" & "\2\3\4\5"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .text = "HHH"
            .Replacement.text = "HH"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .text = "(G)(.)(*)(.[0-9]{1,4}.)([abcd1234])([.: ^13])"
            .Replacement.text = "DS" & "\2\3\4\5"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "([GD])([TS])(.)(*)(.[0-9]{1,4}.)([abcd1234])([.: ^13])"
            .Replacement.text = "DS" & "\3\4\5\6"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "([DHG])([SHT])(.)(*)(.[0-9]{1,4}.)([abcd1234])"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Do While Selection.Find.Execute = True
            tam = Split(Selection, ".")
            Select Case tam(1)
                Case "I"
                Chuong = "C1"
                Case "II"
                Chuong = "C2"
                Case "III"
                Chuong = "C3"
                Case "IV"
                Chuong = "C4"
                Case "V"
                Chuong = "C5"
                Case "VI"
                Chuong = "C6"
                Case "VII"
                Chuong = "C7"
                Case "VIII"
                Chuong = "C8"
                Case "IX"
                Chuong = "C9"
                Case "X"
                Chuong = "C10"
                Case "1"
                Chuong = "C1"
                Case "2"
                Chuong = "C2"
                Case "3"
                Chuong = "C3"
                Case "4"
                Chuong = "C4"
                Case "5"
                Chuong = "C5"
                Case "6"
                Chuong = "C6"
                Case "7"
                Chuong = "C7"
                Case "8"
                Chuong = "C8"
                Case "9"
                Chuong = "C9"
                Case "10"
                Chuong = "C10"
            End Select
            
            Select Case tam(5)
                Case "1"
                tam(5) = "a"
                Case "2"
                tam(5) = "b"
                Case "3"
                tam(5) = "c"
                Case "4"
                tam(5) = "d"
            End Select
            If tam(0) = "GT" Then tam(0) = "DS"
            If UBound(tam) >= 5 Then Selection.TypeText text:="[" & tam(0) & _
            IDclass.text & "." & Chuong & "." & tam(2) & "." & tam(3) & "." & tam(5) & "]"
            Selection.EndKey Unit:=wdLine
        Loop
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "H[HH"
            .Replacement.text = "[HH"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        Do While .Execute
            .Execute Replace:=wdReplaceAll
        Loop
        End With
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.ColorIndex = wdPink
        With Selection.Find
            .text = "(\[)([DGH])([STH])(?????????)(T)(.)([abcd])(\])"
            .Replacement.text = "\1\2\3\4\5\6\7\8"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Font.ColorIndex = wdPink
        With Selection.Find
            .text = "(\[)([DGH])([STH])(?????????)(T)(.)([abcd])(\])(^13)"
            .Replacement.text = "\1\2\3\4\5\6\7\8"
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
End Sub
Private Sub Label16_Click()
        On Error Resume Next
        Dim tam As String
        S_ImBank.Width = 204
        Dim Chuong As String
        Selection.HomeKey Unit:=wdStory
        If S_ImBank.Convert.Value Then
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([DH])([SH]1)"
                .Forward = True
                .Replacement.text = "[" & "\2"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(a)(\])"
                .Forward = True
                .Replacement.text = "1"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(b)(\])"
                .Forward = True
                .Replacement.text = "2"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(c)(\])"
                .Forward = True
                .Replacement.text = "3"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(d)(\])"
                .Forward = True
                .Replacement.text = "4"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([DH])([012])(.C)([0-9]{1})(.)([0-9]{1})(.[BL]T.)([1234])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "-" & "\9" & "]"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "-0-"
                .Forward = True
                .Replacement.text = "-"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        Else

            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([012])([HD])([0-9]{1})(-)([0-9]{1})(\])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "1" & "\2" & ".C" & "\4" & ".0.BT." & "\6" & "]"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([012])([HD])([0-9]{1})(-)([0-9]{1})(-)([0-9]{1})(\])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "1" & "\2" & ".C" & "\4" & "." & "\6" & ".BT." & "\8" & "]"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "[D1"
                .Forward = True
                .Replacement.text = "[DS1"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
          
            With Selection.Find
                .text = "[H"
                .Forward = True
                .Replacement.text = "[HH"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = "[HHH"
                .Forward = True
                .Replacement.text = "[HH"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".1]"
                .Forward = True
                .Replacement.text = ".a]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".2]"
                .Forward = True
                .Replacement.text = ".b]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".3]"
                .Forward = True
                .Replacement.text = ".c]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".4]"
                .Forward = True
                .Replacement.text = ".d]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdPink
            With Selection.Find
                .text = "(\[)([DGH])([STH])(?????????)(T)(.)([abcd])(\])"
                .Replacement.text = "\1\2\3\4\5\6\7\8"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdPink
            With Selection.Find
                .text = "(\[)([DGH])([STH])(?????????)(T)(.)([abcd])(\])(^13)"
                .Replacement.text = "\1\2\3\4\5\6\7\8"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
                .Format = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        End If
End Sub

Private Sub Label21_Click()
On Error Resume Next
        Dim tam As String
        S_ImBank.Width = 204
        Dim Chuong As String
        Selection.HomeKey Unit:=wdStory
        
        If S_ImBank.Convert.Value Then
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([DH])([SH]1)(?.C?.?.???.)"
                .Forward = True
                .Replacement.text = "[" & "\2\4"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(.???.)(a)(\])"
                .Forward = True
                .Replacement.text = "\1" & "1"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(.???.)(b)(\])"
                .Forward = True
                .Replacement.text = "\1" & "2"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(.???.)(c)(\])"
                .Forward = True
                .Replacement.text = "\1" & "3"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(.???.)(d)(\])"
                .Forward = True
                .Replacement.text = "\1" & "4"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([DH])([012])(.C)([0-9]{1})(.)([0-9]{1}.D)([0-9]{2}.)([1234])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "\8" & "-" & "\9" & "]"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([DH])([012])(.C)([0-9]{1})(.)([0-9]{1}.D0)([0-9]{1}.)([1234])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "\2" & "\5" & "-" & "\7" & "\8" & "-" & "\9" & "]"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = ".D0"
                .Forward = True
                .Replacement.text = "."
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = ".D"
                .Forward = True
                .Replacement.text = "."
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = ".-"
                .Forward = True
                .Replacement.text = "-"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        Else
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([012])([HD])([0-9]{1})(-)([0-9]{1}.)([0-9]{2})(-)([0-9]{1}\])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "1" & "\2" & ".C" & "\4" & "." & "\6" & "D" & "\7" & "." & "\9"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "(\[)([012])([HD])([0-9]{1})(-)([0-9]{1}.)([0-9]{1})(-)([0-9]{1}\])"
                .Forward = True
                .Replacement.text = "[" & "\3" & "1" & "\2" & ".C" & "\4" & "." & "\6" & "D0" & "\7" & "." & "\9"
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .text = "[D1"
                .Forward = True
                .Replacement.text = "[DS1"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = "[H1"
                .Forward = True
                .Replacement.text = "[HH1"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            With Selection.Find
                .text = ".1]"
                .Forward = True
                .Replacement.text = ".a]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".2]"
                .Forward = True
                .Replacement.text = ".b]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".3]"
                .Forward = True
                .Replacement.text = ".c]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            With Selection.Find
                .text = ".4]"
                .Forward = True
                .Replacement.text = ".d]"
                .Wrap = wdFindContinue
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdPink
            With Selection.Find
                .text = "(\[)([DGH])([STH])(*)([abcd])(\])"
                .Replacement.text = "\1\2\3\4\5\6"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Font.ColorIndex = wdPink
            With Selection.Find
                .text = "(\[)([DGH])([STH])(*)([abcd])(\])(^13)"
                .Replacement.text = "\1\2\3\4\5\6\7"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchWildcards = True
                .Format = True
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        End If
End Sub

Private Sub Label25_Click()
    Call CheckDrive
    ChangeFileOpenDirectory (S_Drive & "S_Bank&Test\S_Split\")
    If Dialogs(wdDialogFileOpen).Show = -1 Then
    ActiveWindow.View = wdPrintView
    End If
End Sub

Private Sub tachChuongBai_Click()
Call S_SerialHDD
End Sub

Private Sub tachChuyenDe_Click()
Call S_SerialHDD
End Sub

Private Sub Theo_Bai_Click()
    IDdang.list = Array("LT", "BT")
    IDdang.text = "BT"
    CoHDgiai.Enabled = False
    KhongHDgiai.Value = True
    ktCD = False
End Sub

Private Sub Theo_CD_Click()
    IDdang.list = Array("D01", "D02", "D03", "D04", "D05", "D06", "D07", "D08", "D09", "D10")
    IDdang.text = "D01"
    CoHDgiai.Enabled = True
    CoHDgiai.Value = False
    KhongHDgiai.Value = False
    Label5 = "Chuyên dê"
    ktCD = True
End Sub

Private Sub UserForm_Initialize()
Call CheckDrive
IDclass.list = Array("10", "11", "12")
IDmon.list = Array("DS", "HH", "LY", "HO", "SI", "SU", "DI", "CD", "TI", "CN")
IDchuong.list = Array("1", "2", "3", "4", "5", "6", "7", "8", "9")
IDbai.list = Array("1", "2", "3", "4", "5", "6", "7", "8", "9")
IDdang.list = Array("LT", "BT")
IDmucdo.list = Array("a", "b", "c", "d")
If ktCD Then
    'Theo_Bai.Enabled = False
    'Theo_CD.Enabled = True
    Theo_CD.Value = True
Else
    'Theo_CD.Enabled = False
    'Theo_Bai.Enabled = True
    Theo_Bai.Value = True
End If
tachChuongBai = True
End Sub
