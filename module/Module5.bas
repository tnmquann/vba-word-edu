Attribute VB_Name = "Module5"
Sub Chuyen_SmartTest(ByVal control As Office.IRibbonControl)
ActiveDocument.ConvertNumbersToText
Call ThayThe("Câu ^?^?^?", "# ")
Call ThayThe("Câu ^?^?", "# ")
Call ThayThe("A.", "A. ")
Call ThayThe("B.", "^p B. ")
Call ThayThe("C.", "^p C. ")
Call ThayThe("D.", "^p D. ")
Call ThayThe(" A.", "A.")
Call ThayThe(" B.", "B.")
Call ThayThe(" C.", "C.")
Call ThayThe(" D.", "D.")
Call ThayThe("A.", "A. ")
Call ThayThe("B.", "B. ")
Call ThayThe("C.", "C. ")
Call ThayThe("D.", "D. ")
Call ThayThe("  ", "")
Selection.WholeStory
Selection.ParagraphFormat.TabStops.ClearAll
With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
    End With
Call xoadongtrong
End Sub
Sub ThayThe(ByVal sFind As String, ByVal sReplace As String)
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
.text = sFind
.Replacement.text = sReplace
.Wrap = wdFindContinue
.Execute Replace:=wdReplaceAll
End With
End Sub
Sub xoadongtrong()
   Dim oPara As Word.Paragraph
    Dim var
    Dim SpaceCounter As Long
    Dim oChar As Word.Characters
    For Each oPara In ActiveDocument.Paragraphs
        If Len(oPara.Range) = 1 Then
            oPara.Range.Delete
        Else
            SpaceCounter = 0
            Set oChar = oPara.Range.Characters
            For var = 1 To oChar.Count
                If Asc(oChar(var)) = 32 Then
                    SpaceCounter = SpaceCounter + 1
                End If
            Next
            If SpaceCounter + 1 = Len(oPara.Range) Then
                 ' paragraph contains ONLY spaces
                oPara.Range.Delete
            End If
        End If
    Next
End Sub

