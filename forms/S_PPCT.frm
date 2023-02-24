VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} S_PPCT 
   Caption         =   "Create Data"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   2865
   ClientWidth     =   7320
   OleObjectBlob   =   "S_PPCT.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "S_PPCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Title, msg As String
Dim i As Integer
Dim idx As Byte
Dim Chuong As Byte
Dim tabNum() As Byte
Dim SoBai() As Byte
Dim titleCD() As String

Private Sub load_CT()
    Dim www As New Word.Application
    Dim myDoc As New Word.Document
    
    Dim add As String
    add = "ChDe" & S_PPCT.ComboBox2 & "_" & S_PPCT.ComboBox1 & ".docx"
'Kiem tre Header dang mo thi dong lai
    Dim docOpener As Document
        If docIsOpen(add) Then
            Set docOpener = Application.Documents(add)
            docOpener.Close
            Set docOpener = Nothing
        End If
    If FExists(S_Drive & "S_Bank&Test\S_Templates\" & add) = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Chuyên dê này ch" & ChrW(432) & "a có."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        www.Quit (False)
        Exit Sub
    End If
    Set myDoc = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\" & add, PasswordDocument:="")
    Chuong = myDoc.Tables(1).Rows.Count - 1
    ReDim tabNum(Chuong) As Byte
    ReDim SoBai(Chuong) As Byte
    tabNum(1) = 2
    For i = 2 To Chuong
        tabNum(i) = tabNum(i - 1) + Val(myDoc.Tables(1).Cell(i, 4).Range.text)
    Next i
    For i = 1 To Chuong
        SoBai(i) = Val(myDoc.Tables(1).Cell(i + 1, 4).Range.text)
    Next i
    Dim TSB, j As Byte
    ReDim titleCD(Chuong) As String
    TSB = 1
    For j = 1 To Chuong
        titleCD(j) = "ST"
        For i = 1 To SoBai(j)
        titleCD(j) = titleCD(j) & "," & myDoc.Tables(TSB + i).Cell(1, 3).Range.text
        Next i
        TSB = TSB + SoBai(j)
    Next j
    myDoc.Close
    www.Quit (False)
    boxChuong.Clear
    boxChuong.text = "CHUONG"
    boxBai.text = "CHUYEN DE"
End Sub
Private Sub PPCT_Browers()
    On Error Resume Next
    ListBox1.Clear
    If boxBai.text = "Chuyen de" Or boxBai.text = "" Then Exit Sub
    Dim www As New Word.Application
    Dim myDoc As New Word.Document
    If FExists(S_Drive & "S_Bank&Test\S_Templates\ChDe" & S_PPCT.ComboBox2 & "_" & S_PPCT.ComboBox1 & ".docx") = False Then
        Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        msg = "Chuyên dê này ch" & ChrW(432) & "a có."
        Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
        www.Quit (False)
        Exit Sub
    End If
    Set myDoc = www.Documents.Open(S_Drive & "S_Bank&Test\S_Templates\ChDe" & S_PPCT.ComboBox2 & "_" & S_PPCT.ComboBox1 & ".docx", PasswordDocument:="")
   
    For i = 2 To myDoc.Tables(tabNum(Right(boxChuong.Value, 1)) + Left(boxBai.Value, 1) - 1).Rows.Count
        myDoc.Tables(tabNum(Right(boxChuong.Value, 1)) + Left(boxBai.Value, 1) - 1).Cell(i, 2).Select
        www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        ListBox1.AddItem www.Selection
        myDoc.Tables(tabNum(Right(boxChuong.Value, 1)) + Left(boxBai.Value, 1) - 1).Cell(i, 3).Select
        www.Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        ListBox1.list(i - 2, 1) = www.Selection
    Next i
    www.Quit (False)
End Sub

Private Sub boxBai_Change()
    If boxBai.Value = "CHUYEN DE" Then Exit Sub
    Call PPCT_Browers
End Sub

Private Sub boxChuong_DropButtonClick()
    boxBai.Clear
    If boxChuong.ListCount = 0 Then
        For i = 2 To Chuong + 1
            boxChuong.AddItem "Chuong " & i - 1
        Next i
    End If
End Sub
Private Sub boxChuong_Change()
    idx = boxChuong.listIndex + 2
End Sub
Private Sub boxBai_DropButtonClick()
    On Error Resume Next
    Dim temp() As String
    temp = Split(titleCD(Right(boxChuong.text, 1)), ",")
    If boxBai.ListCount = 0 Then
        For i = 1 To SoBai(Right(boxChuong.text, 1))
            boxBai.AddItem i & ". " & Left(temp(i), Len(temp(i)) - 2)
        Next i
    End If
End Sub


Private Sub TaoDL()
On Error GoTo S_Quit
Call CheckDrive
If S_PPCT.ComboBox1.Value = "" Or S_PPCT.ComboBox2.Value = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Thông tin ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7847) & "y " & ChrW(273) & ChrW(7911) & "."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
    Dim i, j As Byte
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Set ThisDoc = ActiveDocument
    S_PPCT.Hide
For j = 1 To ThisDoc.Tables.Count
    Set ThatDoc = Documents.add
    Call S_PageSetup
    For i = 1 To ThisDoc.Tables(j).Rows.Count - 1
    With ThisDoc.Range
        .Tables(j).Cell(i + 1, 3).Select
        Selection.MoveLeft Count:=1, Extend:=wdExtend
        Selection.Copy
    End With
       
    ThatDoc.Activate
    Selection.Font.Size = 12
    ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Selection.Tables(1).Style = "Colorful Grid - Accent 1"
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
    'Selection.InsertRows 1
    Selection.Collapse Direction:=wdCollapseStart
    
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "1"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - LÝ THUY" & ChrW(7870) & "T"
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    
    ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Selection.Tables(1).Style = "Colorful Grid - Accent 2"
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
    'Selection.InsertRows 1
    Selection.Collapse Direction:=wdCollapseStart
   
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "2"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - LÝ THUY" & ChrW(7870) & "T"
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    
    ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Selection.Tables(1).Style = "Colorful Grid - Accent 3"
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
    'Selection.InsertRows 1
    Selection.Collapse Direction:=wdCollapseStart
   
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "3"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - LÝ THUY" & ChrW(7870) & "T"
   
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    
    ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    Selection.Tables(1).Style = "Colorful Grid - Accent 4"
    Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
    'Selection.InsertRows 1
    Selection.Collapse Direction:=wdCollapseStart
    
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "4"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - LÝ THUY" & ChrW(7870) & "T"
 
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    'Bai tap
    'If S_PPCT.ComboBox3.Value = "Bai" Then
        ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
        Selection.Tables(1).Style = "Colorful Grid - Accent 1"
        Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
        'Selection.InsertRows 1
        Selection.Collapse Direction:=wdCollapseStart
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "1"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - BÀI T" & ChrW(7852) & "P"
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        
        ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
        Selection.Tables(1).Style = "Colorful Grid - Accent 2"
        Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
        'Selection.InsertRows 1
        Selection.Collapse Direction:=wdCollapseStart
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "2"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - BÀI T" & ChrW(7852) & "P"
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        
        ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
        Selection.Tables(1).Style = "Colorful Grid - Accent 3"
        Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
        'Selection.InsertRows 1
        Selection.Collapse Direction:=wdCollapseStart
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "3"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - BÀI T" & ChrW(7852) & "P"
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        
        ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
        Selection.Tables(1).Style = "Colorful Grid - Accent 4"
        Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
        'Selection.InsertRows 1
        Selection.Collapse Direction:=wdCollapseStart
        Selection.TypeText text:="Bài " & i & ".M" & ChrW(272) & "4"
        Selection.MoveRight Unit:=wdCell
        Selection.Paste
        Selection.TypeText text:=" - BÀI T" & ChrW(7852) & "P"
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
    'End If
    Next i
    Dim fname As String
    With ThisDoc.Range
        .Tables(j).Cell(2, 2).Select
        Selection.MoveLeft Count:=1, Extend:=wdExtend
       
            fname = S_Drive & "S_Bank&Test\S_Bank\Lop " & S_PPCT.ComboBox2.Value & "\" & Left(Selection, 7) & j & "].dat"
       
    End With
    If FExists(fname) = False Then
        ThatDoc.SaveAs2 FileName:=fname, _
                FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="159", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
    End If
    ThatDoc.Close (yes)
Next j
Unload S_PPCT
MsgBox "Xong!"
Exit Sub
S_Quit:
MsgBox "Bi lôi!"
End Sub
Private Sub TaoDL_CD()
On Error GoTo S_Quit
Call CheckDrive
If S_PPCT.ComboBox1.Value = "" Or S_PPCT.ComboBox2.Value = "" Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Thông tin ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7847) & "y " & ChrW(273) & ChrW(7911) & "."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
    Dim i, C, b, cd, tabNum As Byte
    Dim Chuong, Bai As Byte
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    Dim SoBai() As Byte
    Set ThisDoc = ActiveDocument
    S_PPCT.Hide
    Chuong = ThisDoc.Tables(1).Rows.Count - 1
    ReDim SoBai(Chuong) As Byte
For i = 1 To Chuong
    SoBai(i) = Val(ThisDoc.Tables(1).Cell(i + 1, 4).Range.text)
Next i
'MsgBox sobai(1) & sobai(2) & sobai(3)
'Exit Sub
    SoBai(0) = 0
    tabNum = 1
    Dim fname As String
    Dim tmdang As String
For C = 1 To Chuong
    If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop " & S_PPCT.ComboBox2.Value & "\Chuyen de\" & ComboBox1 & ComboBox2 & ".C" & C) = False Then _
        MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop " & S_PPCT.ComboBox2.Value & "\Chuyen de\" & ComboBox1 & ComboBox2 & ".C" & C)
    'tabnum = tabnum
    For b = 1 To SoBai(C)
    
        
        'If FExists(fname) = False Then Exit For
        tabNum = tabNum + 1
        
        Call S_PageSetup
        For cd = 1 To ThisDoc.Tables(tabNum).Rows.Count - 1
            Set ThatDoc = Documents.add
            If cd < 10 Then
            tmdang = "D0" & cd
            Else
            tmdang = "D" & cd
            End If
            fname = S_Drive & "S_Bank&Test\S_Bank\Lop " & S_PPCT.ComboBox2.Value & "\Chuyen de\" & ComboBox1 & ComboBox2 & ".C" & C & "\[" & _
            S_PPCT.ComboBox1.Value & S_PPCT.ComboBox2.Value & ".C" & C & "." & b & "." & tmdang & "].dat"
             
             With ThisDoc.Range
                 .Tables(tabNum).Cell(cd + 1, 3).Select
                 Selection.MoveLeft Count:=1, Extend:=wdExtend
                 Selection.Copy
             End With
                
             ThatDoc.Activate
             Selection.Font.Size = 12
             ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
                 DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
             Selection.Tables(1).Style = "Colorful Grid - Accent 1"
             Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
             Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
             Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
             Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
             'Selection.InsertRows 1
             Selection.Collapse Direction:=wdCollapseStart
            
                 Selection.TypeText text:="B" & b & ".D" & cd & ".M" & ChrW(272) & "1"
                 Selection.MoveRight Unit:=wdCell
                 Selection.Paste
             
             Selection.EndKey Unit:=wdStory
             Selection.TypeParagraph
             
             ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
                 DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
             Selection.Tables(1).Style = "Colorful Grid - Accent 2"
             Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
             Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
             Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
             Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
             'Selection.InsertRows 1
             Selection.Collapse Direction:=wdCollapseStart
             
                 Selection.TypeText text:="B" & b & ".D" & cd & ".M" & ChrW(272) & "2"
                 Selection.MoveRight Unit:=wdCell
                 Selection.Paste
             
             Selection.EndKey Unit:=wdStory
             Selection.TypeParagraph
             
             ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
                 DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
             Selection.Tables(1).Style = "Colorful Grid - Accent 3"
             Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
             Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
             Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
             Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
             'Selection.InsertRows 1
             Selection.Collapse Direction:=wdCollapseStart
             
                 Selection.TypeText text:="B" & b & ".D" & cd & ".M" & ChrW(272) & "3"
                 Selection.MoveRight Unit:=wdCell
                 Selection.Paste
             
             Selection.EndKey Unit:=wdStory
             Selection.TypeParagraph
             
             ThatDoc.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
                 DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
             Selection.Tables(1).Style = "Colorful Grid - Accent 4"
             Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
             Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
             Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
             Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=400, RulerStyle:=wdAdjustNone
             'Selection.InsertRows 1
             Selection.Collapse Direction:=wdCollapseStart
             
                 Selection.TypeText text:="B" & b & ".D" & cd & ".M" & ChrW(272) & "4"
                 Selection.MoveRight Unit:=wdCell
                 Selection.Paste
            
             Selection.EndKey Unit:=wdStory
             Selection.TypeParagraph
             
            If FExists(fname) = False Then
                ThatDoc.SaveAs2 FileName:=fname, _
                FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="159", AddToRecentFiles _
                :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
                :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                SaveAsAOCELetter:=False, CompatibilityMode:=15
            End If
            ThatDoc.Close (yes)
        Next cd
    Next b
Next C
Unload S_PPCT
MsgBox "Xong!"
Exit Sub
S_Quit:
MsgBox "Bi lôi!"
End Sub


Private Sub ComboBox1_Change()
    boxChuong.text = ""
    boxBai.text = ""
End Sub

Private Sub ComboBox2_Change()
    boxChuong.text = ""
    boxBai.text = ""
End Sub

Private Sub Label6_Click()
    MsgBox "Se cap nhat sau!"
End Sub

Private Sub Label7_Click()
Dim txt As String
Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
If ListBox1.listIndex = -1 Then
    msg = "Ch" & ChrW(7885) & "n d" & ChrW(7841) & "ng toán trong List."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
txt = Left(S_PPCT.ListBox1.list(ListBox1.listIndex, 0), Len(S_PPCT.ListBox1.list(ListBox1.listIndex)) - 1)
If S_PPCT.OptionButton1 Then
    txt = txt & ".a]"
ElseIf S_PPCT.OptionButton2 Then
    txt = txt & ".b]"
ElseIf S_PPCT.OptionButton3 Then
    txt = txt & ".c]"
ElseIf S_PPCT.OptionButton4 Then
    txt = txt & ".d]"
Else
    msg = "Ch" & ChrW(432) & "a ch" & ChrW(7885) & "n m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7897) & _
        " cho câu h" & ChrW(7887) & "i."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
Selection.Font.ColorIndex = wdPink
Selection.TypeText text:=txt
End Sub



Private Sub MoDLnguon_Click()
Dim ThisDoc As Document
On Error Resume Next
Call CheckDrive
If S_PPCT.ComboBox1.Value = "" Or S_PPCT.ComboBox2.Value = "" Or (QLtheoBai = False And QLtheoCD = False) Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "Thông tin ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7847) & "y " & ChrW(273) & ChrW(7911) & "."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If

Dim add As String
Dim Lop, Mon, cd As String
If TaoData.Enabled = False Then
    Call load_CT
    Exit Sub
End If

If QLtheoBai Then
    add = S_Drive & "S_Bank&Test\S_Templates\PPCT" & S_PPCT.ComboBox2.Value & "_" _
            & S_PPCT.ComboBox1.Value & ".docx"
Else
    add = S_Drive & "S_Bank&Test\S_Templates\ChDe" & S_PPCT.ComboBox2.Value & "_" _
            & S_PPCT.ComboBox1.Value & ".docx"
End If
Lop = S_PPCT.ComboBox2.Value
Mon = S_PPCT.ComboBox1.Value
'cd = S_PPCT.QLtheoCD
If FExists(add) = False Then
    Title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
    msg = "PPCT này ch" & ChrW(432) & "a có."
    Application.Assistant.DoAlert Title, msg, 0, 4, 0, 0, 0
    Exit Sub
End If
Unload S_PPCT

Set ThisDoc = Documents.Open(add, ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:="")
ThisDoc.Activate
Load S_PPCT
S_PPCT.ComboBox2.Value = Lop
S_PPCT.ComboBox1.Value = Mon
'S_PPCT.ComboBox3.Value = cd
S_PPCT.Show
End Sub


Private Sub QLChuyenDe_Click()
    TaoData.Enabled = False
    S_PPCT.Height = 357
End Sub

Private Sub QLtheoBai_Click()
    TaoData.Enabled = True
    S_PPCT.Height = 106
    QLChuyenDe.Enabled = False
    ktCD = False
End Sub

Private Sub QLtheoCD_Click()
    TaoData.Enabled = True
    S_PPCT.Height = 106
    QLChuyenDe.Enabled = True
    ktCD = True
End Sub

Private Sub TaoData_Click()
If ktCD Then
    Call TaoDL_CD
Else
    Call TaoDL
End If
End Sub

Private Sub ThemCD_Click()
S_PPCT.Height = 380
End Sub

Private Sub UserForm_Initialize()
Dim path As String
Call CheckDrive
If DirExists("C:\Program Files (x86)\") Then
    path = "C:\Program Files (x86)\"
Else
    path = "C:\Program Files\"
End If
If DirExists(S_Drive & "S_Bank&Test\S_Data") = False Then
    Call MadeDir
End If
If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de") = False Then
    MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 12\Chuyen de")
End If
If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de") = False Then
    MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 11\Chuyen de")
End If
If DirExists(S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de") = False Then
    MkDir (S_Drive & "S_Bank&Test\S_Bank\Lop 10\Chuyen de")
End If
If FExists(S_Drive & "S_Bank&Test\S_Templates\ChDe12_DS.docx") = False Then
    FileCopy path & "S_Bank&Test\S_Templates\ChDe12_DS.docx", S_Drive & "S_Bank&Test\S_Templates\ChDe12_DS.docx"
End If
If FExists(S_Drive & "S_Bank&Test\S_Templates\ChDe12_HH.docx") = False Then
    FileCopy path & "S_Bank&Test\S_Templates\ChDe12_HH.docx", S_Drive & "S_Bank&Test\S_Templates\ChDe12_HH.docx"
End If
ComboBox1.list = Array("DS", "HH", "LY", "HO", "SI", "SU", "DI", "CD", "TI", "CN")
ComboBox2.list = Array("10", "11", "12")
QLChuyenDe.Enabled = False
ListBox1.ColumnWidths = "80"
If ktCD Then
    S_PPCT.QLtheoCD = True
Else
    S_PPCT.QLtheoBai = True
End If
End Sub
