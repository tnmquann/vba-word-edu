Attribute VB_Name = "Module11"
Sub Tach_theo_chuong(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
    If DirName("D:\" & "Tach Theo Chuong" & "\") = False Then
        MkDir ("D:\" & "Tach Theo Chuong" & "\")
    End If
    Dim FileName, DocName
    FileName = "D:\" & "Tach Theo Chuong" & "\" & ActiveDocument.Name
    ActiveDocument.SaveAs FileName
    DocName = ActiveDocument.Name
    Call Add_End_Cau
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "["
        .Replacement.text = "#"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "]"
        .Replacement.text = "~"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    For i = 0 To 2
        For m = 1 To 2
        If m = 1 Then
            Mn = "D"
            If i = 2 Then
                    Mon = "DaiSo"
                Else
                If i = 1 Then
                    Mon = "DaiSo"
                Else
                    Mon = "DaiSo"
                End If
            End If
        Else
            Mn = "H"
            Mon = "HinhHoc"
        End If
            For j = 1 To 6
                For k = 1 To 32
                    If k = 1 Then
                    md = "Y1"
                    Mucdo = "Biet"
                       Else
                       If k = 2 Then
                       md = "Y2"
                       Mucdo = "Biet"
                    Else
                      If k = 3 Then
                       md = "Y3"
                       Mucdo = "Biet"
            Else
                      If k = 4 Then
                       md = "Y4"
                       Mucdo = "Biet"
Else
                      If k = 5 Then
                       md = "Y5"
                       Mucdo = "Biet"
Else
                      If k = 6 Then
                       md = "Y6"
                       Mucdo = "Biet"
Else
                      If k = 7 Then
                       md = "Y7"
                       Mucdo = "Biet"
Else
                      If k = 8 Then
                       md = "Y8"
                       Mucdo = "Biet"
Else
                      If k = 9 Then
                      md = "B1"
                      Mucdo = "Hieu"
Else
                      If k = 10 Then
                      md = "B2"
                      Mucdo = "Hieu"
Else
                      If k = 11 Then
                      md = "B3"
                      Mucdo = "Hieu"
Else
                      If k = 12 Then
                      md = "B4"
                      Mucdo = "Hieu"
Else
                      If k = 13 Then
                      md = "B5"
                      Mucdo = "Hieu"
Else
                      If k = 14 Then
                      md = "B6"
                      Mucdo = "Hieu"
Else
                      If k = 15 Then
                      md = "B7"
                      Mucdo = "Hieu"
Else
                      If k = 16 Then
                      md = "B8"
                      Mucdo = "Hieu"
Else
                      If k = 17 Then
                      md = "K1"
                      Mucdo = "VanDung"
Else
                      If k = 18 Then
                      md = "K2"
                      Mucdo = "VanDung"
Else
                      If k = 19 Then
                      md = "K3"
                      Mucdo = "VanDung"
Else
                      If k = 20 Then
                      md = "K4"
                      Mucdo = "VanDung"
Else
                      If k = 21 Then
                      md = "K5"
                      Mucdo = "VanDung"
Else
                      If k = 22 Then
                      md = "K6"
                      Mucdo = "VanDung"
Else
                      If k = 23 Then
                      md = "K7"
                      Mucdo = "VanDung"
Else
                      If k = 24 Then
                      md = "K8"
                      Mucdo = "VanDung"
Else
                      If k = 25 Then
                      md = "G1"
                      Mucdo = "VDCao"
Else
                      If k = 26 Then
                      md = "G2"
                      Mucdo = "VDCao"
Else
                      If k = 27 Then
                      md = "G3"
                      Mucdo = "VDCao"
Else
                      If k = 28 Then
                      md = "G4"
                      Mucdo = "VDCao"
Else
                      If k = 29 Then
                      md = "G5"
                      Mucdo = "VDCao"
Else
                      If k = 30 Then
                      md = "G6"
                      Mucdo = "VDCao"
Else
                      If k = 31 Then
                      md = "G7"
                      Mucdo = "VDCao"
                                Else
                                md = "G8"
                                Mucdo = "VDCao"
                               End If
                              End If
                            End If
                        End If
                    End If
             End If
           End If
                              End If
                            End If
                        End If
                    End If
             End If
                        End If
                              End If
                            End If
                        End If
                    End If
             End If
                        End If
                              End If
                            End If
                        End If
                    End If
             End If
                        End If
                              End If
                            End If
                        End If
                    End If
             End If
             End If
             
                      Tukhoa = "#" & i & Mn & j & md & "~"
                    NewFileName = "Lop1" & i & "_" & Mon & "_Chuong" & j & "_" & Mucdo & ".doc"
                    With Selection.Find
                        .text = Tukhoa
                        .Replacement.text = "#"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                    If Selection.Find.Execute = True Then
                        Call Tach_Key(Tukhoa, NewFileName)
                    End If
                    End With
                   Next k
            Next j
        Next m
    Next i
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "#"
        .Replacement.text = "["
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "~"
        .Replacement.text = "]"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.end^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
Application.ScreenUpdating = True
    thbao = "Các file câu h" & ChrW(7887) & _
            "i theo m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7897) & " " & ChrW(273) & ChrW(227) & _
             " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c" & " l" & ChrW(432) & "u vào th" & _
             ChrW(432) & " m" & ChrW(7909) & "c" & vbCrLf & ActiveDocument.path
Application.Assistant.DoAlert "Thông báo " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng d" & ChrW(7851) _
         & "n l" & ChrW(432) & "u file", thbao, 0, 4, 0, 0, 0
ActiveDocument.Close (No)
End Sub

Private Sub Tach_Key(ByVal Key As String, ByVal NewFileName As String)
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "(Câu [0-9]{1,2})(*)" & Key
        .Replacement.text = Key & "\1\2" & Key
        .Forward = False
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = True
    If Selection.Find.Execute = False Then Exit Sub
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
    Set ThisDoc = ActiveDocument
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = Key & "(Câu [0-9]{1,2}*)(A.*)(B.*)(C.*)(D.*)(z.end)"
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
    ThatDoc.Activate
    With Selection.Find
        .text = Key & "Câu "
        .Replacement.text = "Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "z.end"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "#"
        .Replacement.text = "["
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "~"
        .Replacement.text = "]"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Dim FileName, DocName
    FileName = ThisDoc.path & "\" & NewFileName
    ActiveDocument.SaveAs FileName
    DocName = ActiveDocument.Name
    ThatDoc.Close (No)
    ThisDoc.Activate
    With Selection.Find
        .text = Key & "Câu "
        .Replacement.text = "Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
End Sub
Private Sub Add_End_Cau()
Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    With Selection.Find
        .text = "z.end^p"
        .Replacement.text = ""
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
        .ClearFormatting
        .text = "^p "
        .Replacement.ClearFormatting
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    Do While .Execute
        .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "(Câu [0-9]{1,2})"
        .Replacement.text = "z.end^p\1"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText text:="z.end"
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.end^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
    For i = 1 To ActiveDocument.Tables.Count
    ActiveDocument.Tables(i).Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "z.end^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceOne
    End With
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
        Selection.TypeText "z.end"
    Next i
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Size = 12
        .Color = wdColorGreen
    End With
    With Selection.Find
        .text = "z.end"
        .Replacement.text = "z.end"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
End Sub
Public Function DirName(Origin As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
DirName = fs.folderexists(Origin)
End Function
Private Sub CopyOld_4Cau(ByVal a As String)
Dim ThisDoc As Document
Dim ThatDoc As Document
Application.ScreenUpdating = False
Selection.HomeKey Unit:=wdStory
Set ThisDoc = ActiveDocument
Call Add_EndCau
Selection.Find.ClearFormatting
With Selection.Find
.text = "(Câu [0-9]{1,2}[.:]*)(\[????\])"
.Replacement.text = "\2\1\2"
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.text = "(\[???)" & a & "(\])(Câu [0-9]{1,2}*)(z.end)(^13)"
.Replacement.text = "\1" & a & "\2\3\4\5"
.MatchWildcards = True
If Selection.Find.Execute = True Then
Documents.Open FileName:=ThisDoc.path & "\loai" & a & ThisDoc.Name
Selection.EndKey Unit:=wdStory
Set ThatDoc = ActiveDocument
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
ThatDoc.Activate
Call Del_EndCau
Selection.HomeKey Unit:=wdStory
ActiveDocument.Save
ActiveWindow.Close
ThisDoc.Activate
Call Del_EndCau
Application.ScreenUpdating = True
End Sub

