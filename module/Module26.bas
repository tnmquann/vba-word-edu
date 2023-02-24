Attribute VB_Name = "Module26"
Sub taomatran(ByVal control As Office.IRibbonControl)

Dim arrcotdau(8) As String
arrcotdau(1) = "Ch" & ChrW(7911) & " " & ChrW(273) & ChrW(7873)
    
   arrcotdau(2) = "Nh" & ChrW(7853) & "n bi" & ChrW(7871) & "t"
  
   arrcotdau(3) = "Thông hi" & ChrW(7875) & "u"
   
    arrcotdau(4) = "V" & ChrW(7853) & "n d" & ChrW(7909) & "ng th" & _
         ChrW(7845) & "p"
   
    arrcotdau(5) = "V" & ChrW(7853) & "n d" & ChrW(7909) & "ng cao"
    
    arrcotdau(6) = "T" & ChrW(7893) & "ng"
   
    arrcotdau(7) = "s" & ChrW(7889) & " " & ChrW(273) & "i" & ChrW( _
        7875) & "m"


' thong ke'

Selection.HomeKey Unit:=wdStory

Dim arrchude(18) As String
arrchude(1) = ChrW(7912) & "ng d" & ChrW(7909) & "ng c" & ChrW( _
        7911) & "a " & ChrW(273) & ChrW(7841) & "o hàm"
    
arrchude(2) = "M" & ChrW(361) & " -logarit"

 arrchude(3) = "Nguyên hàm tích phân"

arrchude(4) = "S" & ChrW(7889) & " ph" & ChrW(7913) & "c"

   arrchude(5) = "Kh" & ChrW(7889) & "i " & ChrW(273) & "a di" & _
        ChrW(7879) & "n"

   arrchude(6) = "Nón tr" & ChrW(7909) & " tròn xoay"

    arrchude(7) = "PP t" & ChrW(7885) & "a " & ChrW(273) & ChrW( _
        7897) & " trong không gian"

    arrchude(8) = "Ph" & ChrW(432) & ChrW(417) & "ng trình l" & _
        ChrW(432) & ChrW(7907) & "ng giác"

  arrchude(9) = "Xác su" & ChrW(7845) & "t Nh" & ChrW(7883) & _
        " th" & ChrW(7913) & "c Niuton"
arrchude(10) = " Day so "
    arrchude(11) = "Gi" & ChrW(7899) & "i h" & ChrW(7841) & "n"

   arrchude(12) = ChrW(272) & ChrW(7841) & "o hàm"

    arrchude(13) = "Phép bi" & ChrW(7871) & "n hình"

    arrchude(14) = "Quan h" & ChrW(7879) & " song song"

   arrchude(15) = "Quan h" & ChrW(7879) & " vuông góc"
    arrchude(16) = " "
     arrchude(17) = " "

' so diem'

 ' Giai tich 12'
 
 

Dim arrChuong(20, 5) As Integer


For i = 1 To 4
For j = 1 To 4
Selection.HomeKey Unit:=wdStory
T = 0
With Selection.Find
.text = "(\[2D)" & i & "(?[0-9]{1}?[0-9]{1,2}?)" & j & "(\])"
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
Do While .Execute = True
T = T + 1
Loop
End With
arrChuong(i, j) = T
Next j
Next i

' Hinh hoc 12'
Selection.HomeKey Unit:=wdStory
For i = 1 To 3
For j = 1 To 4
Selection.HomeKey Unit:=wdStory
T = 0
With Selection.Find
.text = "(\[2H)" & i & "(?[0-9]{1}?[0-9]{1,2}?)" & j & "(\])"
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
Do While .Execute = True
T = T + 1
Loop
End With
a = i + 4
arrChuong(a, j) = T
Next j
Next i

' dai so 11'

Selection.HomeKey Unit:=wdStory
For i = 1 To 5
For j = 1 To 4
Selection.HomeKey Unit:=wdStory
T = 0
With Selection.Find
.text = "(\[1D)" & i & "(?[0-9]{1}?[0-9]{1,2}?)" & j & "(\])"
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
Do While .Execute = True
T = T + 1
Loop
End With
b = i + 7
arrChuong(b, j) = T
Next j
Next i

' hinh hoc 11'
Selection.HomeKey Unit:=wdStory

For i = 1 To 3
For j = 1 To 4
Selection.HomeKey Unit:=wdStory
T = 0
With Selection.Find
.text = "(\[1H)" & i & "(?[0-9]{1}?[0-9]{1,2}?)" & j & "(\])"
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchWildcards = True
Do While .Execute = True
T = T + 1
Loop
End With
C = i + 12
arrChuong(C, j) = T
Next j
Next i


For j = 1 To 4
d = 0
For i = 1 To 15
d = d + arrChuong(i, j)
Next i
arrChuong(16, j) = d
Next j

For i = 1 To 16
d = 0
For j = 1 To 4
d = d + arrChuong(i, j)
Next j
arrChuong(i, 5) = d
Next i

Dim sodiem(5) As Double
sodiem(1) = Round((arrChuong(16, 1) / arrChuong(16, 5)) * 10, 1)
sodiem(2) = Round((arrChuong(16, 2) / arrChuong(16, 5)) * 10, 1)
sodiem(3) = Round((arrChuong(16, 3) / arrChuong(16, 5)) * 10, 1)
sodiem(4) = Round(10 - (sodiem(1) + sodiem(2) + sodiem(3)), 1)

' diem cot cuoi cung'

Dim diemcuoi(20) As Double
For i = 1 To 14
diemcuoi(i) = Round((arrChuong(i, 5) / arrChuong(16, 5)) * 10, 1)
Next i
X = 0
For i = 1 To 14
X = X + diemcuoi(i)
Next i
diemcuoi(15) = Round(10 - X, 1)
' tao bang'

Documents.add DocumentType:=wdNewBlankDocument
    Selection.WholeStory
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1)
        .RightMargin = CentimetersToPoints(1)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.27)
        .FooterDistance = CentimetersToPoints(1.27)
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
        .LayoutMode = wdLayoutModeDefault
    End With
    




 Selection.EndKey Unit:=wdStory
    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=18, NumColumns _
        :=8, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
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
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
    Selection.HomeKey Unit:=wdColumn
    Selection.TypeText text:="Phân môn"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="Gi" & ChrW(7843) & "i tích 12"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="Hình h" & ChrW(7885) & "c 12"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:=ChrW(272) & ChrW(7841) & "i s" & ChrW(7889) & _
        " 12"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeBackspace
    Selection.TypeText text:="1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="Hình h" & ChrW(7885) & "c 11"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="T" & ChrW(7893) & "ng s" & ChrW(7889) & " câu"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="T" & ChrW(7893) & "ng s" & ChrW(7889) & " " & _
        ChrW(273) & "i" & ChrW(7875) & "m"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#18#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#17#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#16#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#15#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#14#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#13#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#12#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#11#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#10#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Application.Run MacroName:="MathTypeCommands.UIEnableDisable.UIUpdate"
    Selection.TypeText text:="#9#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#8#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#7#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#6#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2#0"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1#0"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#1#1"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#1#2"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#1#3"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#1#4"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#1#5"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#1#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#6#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#7#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#8#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#9#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#10#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#11#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#12#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#13#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#14#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#15#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#16#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#17#6"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#18"
    Selection.MoveLeft Unit:=wdWord, Count:=2
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#17"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#16"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#15"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#14"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#13"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#12"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveLeft Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#11"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#10"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#9"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Application.Run MacroName:="MathTypeCommands.UIEnableDisable.UIUpdate"
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#8"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#7"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#6"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#1"
    Selection.MoveRight Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#2"
    Selection.MoveRight Unit:=wdCharacter, Count:=3
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#3"
    Selection.MoveRight Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText text:="#4"
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText text:="#5"
    Selection.MoveRight Unit:=wdCharacter, Count:=4
    Selection.TypeText text:="#6"
    Selection.MoveUp Unit:=wdLine, Count:=13
    Selection.TypeText text:="#6"
    Selection.Tables(1).Select
    Selection.SelectCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter

' ma hoa bang'

' hang dau'
For j = 1 To 7
With Selection.Find
.text = "#" & "1" & "#" & j - 1
.Replacement.text = arrcotdau(j)
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
Next j

' hang cuoi '
For j = 1 To 4
With Selection.Find
.text = "#" & "18" & "#" & j
.Replacement.text = sodiem(j)
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
Next j

For i = 2 To 18
With Selection.Find
.text = "#" & i & "#" & "0"
.Replacement.text = arrchude(i - 1)
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
Next i

For i = 2 To 18
For j = 1 To 5
With Selection.Find
.text = "#" & i & "#" & j
.Replacement.text = arrChuong(i - 1, j)
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
Next j
Next i

For i = 1 To 15
With Selection.Find
.text = "#" & i + 1 & "#6"
.Replacement.text = diemcuoi(i)
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
Next i


With Selection.Find
.text = "#17#6"
.Replacement.text = " "
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With

With Selection.Find
.text = "#18#6"
.Replacement.text = "10"
.Replacement.Font.Size = 12
.Replacement.Font.Color = wdColorbulue
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = False
.Execute Replace:=wdReplaceOne
End With
 
    
    Selection.EndKey Unit:=wdStory
  End Sub





