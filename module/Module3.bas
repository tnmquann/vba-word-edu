Attribute VB_Name = "Module3"
Sub Ghi_chu_thichBTN(ByVal control As Office.IRibbonControl)
    ChuthichBTN.OptionButton1 = True
    ChuthichBTN.Show
End Sub
Sub Ghi_chu_thichVDC(ByVal control As Office.IRibbonControl)
    ChuthichVDC.OptionButton1 = True
    ChuthichVDC.Show
End Sub
Sub To_mau_chu_thich(ByVal control As Office.IRibbonControl)
Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "\[" & "([0-2][D-H][1-6]-[1-4])" & "\]"
        .Replacement.text = "#" & "\1" & "~"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 192
    End With
    With Selection.Find
        .text = "\[" & "(*)" & "\]"
        .Replacement.text = "[" & "\1" & "]"
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
    With Selection.Find
        .ClearFormatting
        .text = "#" & "([0-2][D-H][1-6]-[1-4])" & "~"
        .Replacement.text = "[" & "\1" & "]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Xoa_chu_thich(ByVal control As Office.IRibbonControl)

    msg1 = "Xo" & ChrW(225) & " ch" & ChrW(250) & " th" & ChrW(237) & "ch ngu" & ChrW(7891) & "n " & ChrW(273) & "" & ChrW(7873) & " r" & ChrW(7891) & "i khi c" & ChrW(7847) & "n ch" & ChrW(250) & " th" & ChrW(237) & "ch l" & ChrW(7841) & "i s" & ChrW(7869) & " g" & ChrW(226) & "y nhi" & ChrW(7873) & "u kh" & ChrW(243) & " kh" & ChrW(259) & "n." & vbCrLf & "Do " & ChrW(273) & "" & ChrW(243) & " thao t" & ChrW(225) & "c n" & ChrW(224) & "y ch" & ChrW(7881) & " th" & ChrW(7921) & "c hi" & ChrW(7879) & "n tr" & ChrW(234) & "n m" & ChrW(7897) & "t file m" & ChrW(7899) & "i (c" & ChrW(249) & "ng th" & ChrW(432) & " m" & ChrW(7909) & "c v" & ChrW(7899) & "I" & vbCrLf & "file hi" & ChrW(7879) & "n h" & ChrW(224) & "nh c" & ChrW(7911) & "a b" & ChrW(7841) & "n). File g" & ChrW(7889) & "c v" & ChrW(7851) & "n c" & ChrW(242) & "n nguy" & ChrW(234) & "n v" & ChrW(7865) & "n."
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg1, 0, 4, 0, 0, 0
    
    Dim FileName, DocName
    If Right(ActiveDocument.Name, 4) = ".doc" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    Else
    If Right(ActiveDocument.Name, 5) = ".docx" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 5)
    Else
        DocName = ActiveDocument.Name
    End If
    End If

    FileName = ActiveDocument.path & "\" & DocName & " (xoa nguon).doc"
    ActiveDocument.SaveAs FileName
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .text = "\[" & "([0-2][D-H][1-6]-[1-4])" & "\]"
        .Replacement.text = "#" & "\1" & "~"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[" & "(*)" & "\] "
        .Replacement.text = "[" & "\1" & "]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute
    .Execute Replace:=wdReplaceAll
    Loop
    End With
    With Selection.Find
        .text = "\[" & "(*)" & "\]^13"
        .Replacement.text = ""
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
    With Selection.Find
        .ClearFormatting
        .text = "#" & "([0-2][D-H][1-6]-[1-4])" & "~"
        .Replacement.text = "[" & "\1" & "]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey Unit:=wdStory
    msg2 = "Thao t" & ChrW(225) & "c " & ChrW(273) & "" & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t. H" & ChrW(227) & "y nh" & ChrW(7845) & "n Ctrl + S " & ChrW(273) & "" & ChrW(7875) & " l" & ChrW(432) & "u file n" & ChrW(224) & "y l" & ChrW(7841) & "i."
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg2, 0, 4, 0, 0, 0
End Sub
Sub huong_dan_nhap_lieu(ByVal control As Office.IRibbonControl)
    Huongdan.Show
End Sub
Sub To_mau_ky_hieu(ByVal control As Office.IRibbonControl)
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = 16711884
    End With
    With Selection.Find
        .text = "\[" & "([0-2][D-H][1-6]-[1-4])" & "\]"
        .Replacement.text = "[" & "\1" & "]"
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
    msg = "C" & ChrW(244) & "ng vi" & ChrW(7879) & "c ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Dem_ky_hieu(ByVal control As Office.IRibbonControl)
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = False
    Dim Cau, msgCau, msgDuongke
    Dim Mucdo(1 To 4)
    Dim msgMucdo(1 To 4)
        For k = 1 To 4
            Mucdo(k) = 0
            With Selection.Find
                .ClearFormatting
                .text = "(\[[0-2][D-H][1-6]-)" & k & "(\])"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWildcards = True
            End With
            If Selection.Find.Execute = False Then
                Mucdo(k) = 0
                msgMucdo(k) = ""
            Else
            Do
                Mucdo(k) = Mucdo(k) + 1
                msgMucdo(k) = "S" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i thu" & ChrW(7897) & "c m" & ChrW(7913) & "c " & ChrW(273) & "" & ChrW(7897) & " " & k & " l" & ChrW(224) & ": " & Mucdo(k) & vbCrLf
            Loop While Selection.Find.Execute = True
            End If
            Selection.HomeKey Unit:=wdStory
        Next k
    Cau = Mucdo(1) + Mucdo(2) + Mucdo(3) + Mucdo(4)
    msgCau = "T" & ChrW(7893) & "ng s" & ChrW(7889) & " c" & ChrW(226) & "u " & ChrW(273) & "" & ChrW(227) & " " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7907) & "c th" & ChrW(234) & "m k" & ChrW(253) & " hi" & ChrW(7879) & "u l" & ChrW(224) & ": " & Cau & vbCrLf
    msgDuongke = "----------------------------------" & vbCrLf
    Dim Daiso(1 To 6)
    Dim msgDaiso(1 To 6)
        For d = 1 To 6
            Daiso(d) = 0
            With Selection.Find
                .ClearFormatting
                .text = "(\[[0-2]D)" & d & "(-[1-4]\])"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWildcards = True
            End With
            If Selection.Find.Execute = False Then
                Daiso(d) = 0
                msgDaiso(d) = ""
            Else
            Do
                Daiso(d) = Daiso(d) + 1
                msgDaiso(d) = "S" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i " & ChrW(272) & "S v" & ChrW(224) & " Gi" & ChrW(7843) & "i t" & ChrW(237) & "ch - ch" & ChrW(432) & "" & ChrW(417) & "ng " & d & " l" & ChrW(224) & ": " & Daiso(d) & vbCrLf
            Loop While Selection.Find.Execute = True
            End If
            Selection.HomeKey Unit:=wdStory
        Next d
    Dim Hinhhoc(1 To 3)
    Dim msgHinhhoc(1 To 3)
        For h = 1 To 3
            Hinhhoc(h) = 0
            With Selection.Find
                .ClearFormatting
                .text = "(\[[0-2]H)" & h & "(-[1-4]\])"
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWildcards = True
            End With
            If Selection.Find.Execute = False Then
                Hinhhoc(h) = 0
                msgHinhhoc(h) = ""
            Else
            Do
                Hinhhoc(h) = Hinhhoc(h) + 1
                msgHinhhoc(h) = "S" & ChrW(7889) & " c" & ChrW(226) & "u h" & ChrW(7887) & "i H" & ChrW(236) & "nh h" & ChrW(7885) & "c - ch" & ChrW(432) & "" & ChrW(417) & "ng " & h & " l" & ChrW(224) & ": " & Hinhhoc(h) & vbCrLf
            Loop While Selection.Find.Execute = True
            End If
            Selection.HomeKey Unit:=wdStory
        Next h
    Application.ScreenUpdating = True
    msg = msgCau & msgDuongke & msgDaiso(1) & msgDaiso(2) & msgDaiso(3) & msgDaiso(4) & msgDaiso(5) & msgDaiso(6) & msgHinhhoc(1) & msgHinhhoc(2) & msgHinhhoc(3) & msgDuongke & msgMucdo(1) & msgMucdo(2) & msgMucdo(3) & msgMucdo(4)
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg, 0, 4, 0, 0, 0
End Sub
Sub Xoa_ky_hieu(ByVal control As Office.IRibbonControl)
    Dim FileName, DocName
    If Right(ActiveDocument.Name, 4) = ".doc" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    Else
    If Right(ActiveDocument.Name, 5) = ".docx" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 5)
    Else
        DocName = ActiveDocument.Name
    End If
    End If
    FileName = ActiveDocument.path & "\" & DocName & " (xoa ky hieu).doc"
    ActiveDocument.SaveAs FileName
    msg1 = "Vi" & ChrW(7879) & "c t" & ChrW(7841) & "o ra c" & ChrW(225) & "c k" & ChrW(253) & " hi" & ChrW(7879) & "u nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng c" & ChrW(226) & "u h" & ChrW(7887) & "i g" & ChrW(226) & "y m" & ChrW(7845) & "t r" & ChrW(7845) & "t nhi" & ChrW(7873) & "u th" & ChrW(7901) & "i gian" & vbCrLf & "Do " & ChrW(273) & "" & ChrW(243) & " thao t" & ChrW(225) & "c xo" & ChrW(225) & " k" & ChrW(253) & " hi" & ChrW(7879) & "u n" & ChrW(224) & "y ch" & ChrW(7881) & " th" & ChrW(7921) & "c hi" & ChrW(7879) & "n tr" & ChrW(234) & "n file m" & ChrW(7899) & "i (b" & ChrW(7843) & "n sao)" & vbCrLf & "File g" & ChrW(7889) & "c c" & ChrW(7911) & "a b" & ChrW(7841) & "n v" & ChrW(7851) & "n c" & ChrW(242) & "n nguy" & ChrW(234) & "n v" & ChrW(7865) & "n."
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg1, 0, 4, 0, 0, 0
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[" & "([0-2][D-H][1-6]-[1-4])" & "\]"
        .Replacement.text = ""
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
    msg2 = "Thao t" & ChrW(225) & "c " & ChrW(273) & "" & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t. H" & ChrW(227) & "y nh" & ChrW(7845) & "n Ctrl + S " & ChrW(273) & "" & ChrW(7875) & " l" & ChrW(432) & "u file n" & ChrW(224) & "y l" & ChrW(7841) & "i."
    Application.Assistant.DoAlert "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", msg2, 0, 4, 0, 0, 0
End Sub
Sub Huong_dan_tach_de(ByVal control As Office.IRibbonControl)
    HD_Tach.Show
End Sub
Sub Tach_de_TN_theo_chuong_muc_do(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
' Kiem tra xem co ky hieu nhan dang trong tai lieu hay chua
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[" & "([0-2][DHL][1-6]-[1-4])" & "\]"
        .MatchWildcards = True
    If Selection.Find.Execute = False Then
        Title1 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg1 = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a th" & ChrW(234) & "m k" & ChrW(253) & " hi" & ChrW(7879) & "u nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng c" & ChrW(226) & "u h" & ChrW(7887) & "i ho" & ChrW(7863) & "c k" & ChrW(253) & " hi" & ChrW(7879) & "u m" & ChrW(224) & vbCrLf & "b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " th" & ChrW(234) & "m ch" & ChrW(432) & "a " & ChrW(273) & "" & ChrW(250) & "ng theo h" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n c" & ChrW(7911) & "a ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh."
        Application.Assistant.DoAlert Title1, msg1, 0, 4, 0, 0, 0
        Huongdan.Show
        Exit Sub
    End If
    End With
' Tao thu muc chua cac file nguon va file moi duoc tach ra
    Dim FileName, DocName
    If Right(ActiveDocument.Name, 4) = ".doc" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    Else
    If Right(ActiveDocument.Name, 5) = ".docx" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 5)
    Else
        DocName = ActiveDocument.Name
    End If
    End If
    If DirExists("D:\" & "Tach chi tiet\") = False Then
        MkDir ("D:\" & "Tach chi tiet\")
    End If
    If DirExists("D:\" & "Tach chi tiet\" & DocName & "\") = False Then
        MkDir ("D:\" & "Tach chi tiet\" & DocName & "\")
    End If
    ' Luu file nguon vao chung thu muc voi cac file se tach ra duoc
    FileName = "D:\" & "Tach chi tiet\" & DocName & "\" & DocName
    ActiveDocument.SaveAs FileName
' Them ky hieu nhan dang het cau
    Call End_cau
' Doi ky hieu nhan dang sang text
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
' Chep tung nhom cau cung muc do cua moi chuong
    For i = 0 To 2
        For m = 1 To 3
        If m = 1 Then
            Kh = "D"
            If i = 2 Then
                    Mon = "GI" & ChrW(7842) & "I TÍCH"
                    Monhoc = "Giai tich"
                Else
                If i = 1 Then
                    Mon = "ÐS & GI" & ChrW(7842) & "I TÍCH"
                    Monhoc = "DS & Giai tich"
                Else
                    Mon = "Ð" & ChrW(7840) & "I S" & ChrW(7888)
                    Monhoc = "Dai so"
                End If
            End If
        Else
        If m = 2 Then
            Kh = "H"
            Mon = "H" & ChrW(204) & "NH H" & ChrW(7884) & "C"
            Monhoc = "Hinh hoc"
        Else
            Kh = "L"
            Mon = "V" & ChrW(7852) & "T L" & ChrW(221)
            Monhoc = "Vat ly"
        End If
        End If
            For j = 1 To 6
                For k = 1 To 4
                    If k = 1 Then
                        Mucdo = "NB"
                        Else
                        If k = 2 Then
                            Mucdo = "TH"
                            Else
                            If k = 3 Then
                                Mucdo = "VDT"
                                Else
                                Mucdo = "VDC"
                            End If
                        End If
                    End If
                    Tukhoa = "#" & i & Kh & j & "-" & k & "~"
                    NewFileName = "Lop 1" & i & " - " & Monhoc & " - Chuong " & j & " - Muc do " & k & " [" & Mucdo & "]"
                    With Selection.Find
                        .text = Tukhoa
                        .Replacement.text = "#"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                    If Selection.Find.Execute = True Then
                        Call Tach_cau_hoi(Tukhoa, NewFileName, "- M" & ChrW(7912) & "C Ð" & ChrW(7896) & " " & k & " - CH" & ChrW(431) & ChrW(416) & "NG " & j & " - " & Mon & " 1" & i)
                    End If
                    End With
                Next k
            Next j
        Next m
    Next i
Application.ScreenUpdating = True
    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7901) & "ng d" & ChrW(7851) & "n file"
    msg2 = "Các file câu h" & ChrW(7887) & _
            "i theo ch" & ChrW(432) & ChrW(417) & "ng, m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7897) & " " & ChrW(273) & ChrW(227) & _
             " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c" & " l" & ChrW(432) & "u vào th" & ChrW(432) & " m" & ChrW(7909) & "c" & vbCrLf & ActiveDocument.path
    Application.Assistant.DoAlert title2, msg2, 0, 4, 0, 0, 0
    ActiveDocument.Close (No)
End Sub
Sub Tach_de_TN_theo_chuong(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
' Kiem tra xem co ky hieu nhan dang trong tai lieu hay chua
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[" & "([0-2][DHL][1-6]-[1-4])" & "\]"
        .MatchWildcards = True
    If Selection.Find.Execute = False Then
        Title1 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg1 = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a th" & ChrW(234) & "m k" & ChrW(253) & " hi" & ChrW(7879) & "u nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng c" & ChrW(226) & "u h" & ChrW(7887) & "i ho" & ChrW(7863) & "c k" & ChrW(253) & " hi" & ChrW(7879) & "u m" & ChrW(224) & vbCrLf & "b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " th" & ChrW(234) & "m ch" & ChrW(432) & "a " & ChrW(273) & "" & ChrW(250) & "ng theo h" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n c" & ChrW(7911) & "a ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh."
        Application.Assistant.DoAlert Title1, msg1, 0, 4, 0, 0, 0
        Huongdan.Show
        Exit Sub
    End If
    End With
' Tao thu muc chua cac file nguon va file moi duoc tach ra
    Dim FileName, DocName
    If Right(ActiveDocument.Name, 4) = ".doc" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    Else
    If Right(ActiveDocument.Name, 5) = ".docx" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 5)
    Else
        DocName = ActiveDocument.Name
    End If
    End If
    If DirExists("D:\" & "Tach theo chuong\") = False Then
        MkDir ("D:\" & "Tach theo chuong\")
    End If
    If DirExists("D:\" & "Tach theo chuong\" & DocName & "\") = False Then
        MkDir ("D:\" & "Tach theo chuong\" & DocName & "\")
    End If
    ' Luu file nguon vao chung thu muc voi cac file se tach ra duoc
    FileName = "D:\" & "Tach theo chuong\" & DocName & "\" & DocName
    ActiveDocument.SaveAs FileName
' Them ky hieu nhan dang het cau
    Call End_cau
' Doi ky hieu nhan dang sang text
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
' Chep tung nhom cau cung muc do cua moi chuong
    For i = 0 To 2
        For m = 1 To 3
        If m = 1 Then
            Kh = "D"
            If i = 2 Then
                    Mon = "GI" & ChrW(7842) & "I TÍCH"
                    Monhoc = "Giai tich"
                Else
                If i = 1 Then
                    Mon = "ÐS & GI" & ChrW(7842) & "I TÍCH"
                    Monhoc = "DS & Giai tich"
                Else
                    Mon = "Ð" & ChrW(7840) & "I S" & ChrW(7888)
                    Monhoc = "Dai so"
                End If
            End If
        Else
        If m = 2 Then
            Kh = "H"
            Mon = "H" & ChrW(204) & "NH H" & ChrW(7884) & "C"
            Monhoc = "Hinh hoc"
        Else
            Kh = "L"
            Mon = "V" & ChrW(7852) & "T L" & ChrW(221)
            Monhoc = "Vat ly"
        End If
        End If
            For j = 1 To 6
                    Tukhoa = "#" & i & Kh & j & "-"
                    NewFileName = "Lop 1" & i & " - " & Monhoc & " - Chuong " & j & " - " & DocName
                    With Selection.Find
                        .text = Tukhoa
                        .Replacement.text = "#"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                    If Selection.Find.Execute = True Then
                        Call Tach_cau_hoi(Tukhoa, NewFileName, " - CH" & ChrW(431) & ChrW(416) & "NG " & j & " - " & Mon & " 1" & i)
                    End If
                    End With
            Next j
        Next m
    Next i
    Application.ScreenUpdating = True
    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7901) & "ng d" & ChrW(7851) & "n file"
    msg2 = "Các file câu h" & ChrW(7887) & _
            "i theo ch" & ChrW(432) & ChrW(417) & "ng" & " " & ChrW(273) & ChrW(227) & _
             " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c" & " l" & ChrW(432) & "u vào th" & ChrW(432) & " m" & ChrW(7909) & "c" & vbCrLf & ActiveDocument.path
    Application.Assistant.DoAlert title2, msg2, 0, 4, 0, 0, 0
    ActiveDocument.Close (No)
End Sub
Sub Tach_de_TN_theo_muc_do(ByVal control As Office.IRibbonControl)
Application.ScreenUpdating = False
ActiveDocument.Range.ListFormat.ConvertNumbersToText
' Kiem tra xem co ky hieu nhan dang trong tai lieu hay chua
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[" & "([0-2][DHL][1-6]-[1-4])" & "\]"
        .MatchWildcards = True
    If Selection.Find.Execute = False Then
        Title1 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i"
        msg1 = "B" & ChrW(7841) & "n ch" & ChrW(432) & "a th" & ChrW(234) & "m k" & ChrW(253) & " hi" & ChrW(7879) & "u nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng c" & ChrW(226) & "u h" & ChrW(7887) & "i ho" & ChrW(7863) & "c k" & ChrW(253) & " hi" & ChrW(7879) & "u m" & ChrW(224) & vbCrLf & "b" & ChrW(7841) & "n " & ChrW(273) & "" & ChrW(227) & " th" & ChrW(234) & "m ch" & ChrW(432) & "a " & ChrW(273) & "" & ChrW(250) & "ng theo h" & ChrW(432) & "" & ChrW(7899) & "ng d" & ChrW(7851) & "n c" & ChrW(7911) & "a ch" & ChrW(432) & "" & ChrW(417) & "ng tr" & ChrW(236) & "nh."
        Application.Assistant.DoAlert Title1, msg1, 0, 4, 0, 0, 0
        Huongdan.Show
        Exit Sub
    End If
    End With
' Tao thu muc chua cac file nguon va file moi duoc tach ra
    Dim FileName, DocName
    If Right(ActiveDocument.Name, 4) = ".doc" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    Else
    If Right(ActiveDocument.Name, 5) = ".docx" Then
        DocName = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 5)
    Else
        DocName = ActiveDocument.Name
    End If
    End If
    If DirExists("D:\" & "Tach theo muc do\") = False Then
        MkDir ("D:\" & "Tach theo muc do\")
    End If
    If DirExists("D:\" & "Tach theo muc do\" & DocName & "\") = False Then
        MkDir ("D:\" & "Tach theo muc do\" & DocName & "\")
    End If
    ' Luu file nguon vao chung thu muc voi cac file se tach ra duoc
    FileName = "D:\" & "Tach theo muc do\" & DocName & "\" & DocName
    ActiveDocument.SaveAs FileName
' Them ky hieu nhan dang het cau
    Call End_cau
' Doi ky hieu nhan dang sang text
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
' Chep tung nhom cau cung muc do cua moi chuong
                For k = 1 To 4
                    If k = 1 Then
                        Mucdo = "NB"
                        Else
                        If k = 2 Then
                            Mucdo = "TH"
                            Else
                            If k = 3 Then
                                Mucdo = "VDT"
                                Else
                                Mucdo = "VDC"
                            End If
                        End If
                    End If
                    Tukhoa = "-" & k & "~"
                    NewFileName = "Muc do " & k & " [" & Mucdo & "]" & " - " & DocName
                    With Selection.Find
                        .text = Tukhoa
                        .Replacement.text = "#"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .MatchWildcards = False
                    If Selection.Find.Execute = True Then
                        Call Tach_cau_hoi(Tukhoa, NewFileName, "- M" & ChrW(7912) & "C Ð" & ChrW(7896) & " " & k & " [" & Mucdo & "]")
                    End If
                    End With
                Next k
Application.ScreenUpdating = True
    title2 = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o " & ChrW(273) & "" & ChrW(432) & "" & ChrW(7901) & "ng d" & ChrW(7851) & "n file"
    msg2 = "Các file câu h" & ChrW(7887) & _
            "i theo m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7897) & " " & ChrW(273) & ChrW(227) & _
             " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c" & " l" & ChrW(432) & "u vào th" & ChrW(432) & " m" & ChrW(7909) & "c" & vbCrLf & ActiveDocument.path
    Application.Assistant.DoAlert title2, msg2, 0, 4, 0, 0, 0
    ActiveDocument.Close (No)
End Sub

Private Sub Tach_cau_hoi(ByVal Key As String, ByVal NewFileName As String, ByVal Titles As String)
    Dim ThisDoc As Document
    Dim ThatDoc As Document
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Application.ScreenUpdating = False
' Nhan dang cau chua tu khoa can copy
    Selection.HomeKey Unit:=wdStory
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
' Them dau cach khoang sau moi tu khoa phuong an
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
        .text = Key & "(Câu [0-9]{1,4}*)(A.*)(B.*)(C.*)(D.*)(z.zz)"
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
        .text = "z.zz"
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
    Dim FileName, DocName
    FileName = ThisDoc.path & "\" & NewFileName & ".docx"
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
Private Sub End_cau()
' Convert_McMix Macro
Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.WholeStory
    With Selection.Find
        .text = "z.zz^p"
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
        .text = " Câu "
        .Replacement.text = " Câu$"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWholeWord = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = "Câu "
        .Replacement.text = "z.zz^pCâu "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .MatchWholeWord = True
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .text = " Câu$"
        .Replacement.text = " Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText text:="z.zz"
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .text = "z.zz^p"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceOne
    End With
    
    For i = 1 To ActiveDocument.Tables.Count
    ActiveDocument.Tables(i).Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "z.zz^p"
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
        Selection.TypeText "z.zz"
    Next i
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Size = 12
        .Color = wdColorGreen
    End With
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = "z.zz"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
Application.ScreenUpdating = True
    Selection.HomeKey Unit:=wdStory
End Sub
Public Function DirExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
DirExists = fs.folderexists(OrigFile)
End Function
