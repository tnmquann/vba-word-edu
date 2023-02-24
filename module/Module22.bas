Attribute VB_Name = "Module22"
Sub sapxepmucdo(ByVal control As Office.IRibbonControl)
Call Mucdo("1")
Call Mucdo("2")
Call Mucdo("3")
Call Mucdo("4")
Documents.Open FileName:="D:\Tachtheomucdo\filedasapxep.doc"
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.EndKey Unit:=wdStory
    ChangeFileOpenDirectory "D:\Tachtheomucdo\"
    Selection.InsertFile FileName:="nhanbiet.doc", Range:="", _
        ConfirmConversions:=False, Link:=False, Attachment:=False
    Selection.InsertFile FileName:="thonghieu.doc", Range:="", _
        ConfirmConversions:=False, Link:=False, Attachment:=False
    Selection.InsertFile FileName:="vandungthap.doc", Range:="", _
        ConfirmConversions:=False, Link:=False, Attachment:=False
    Selection.InsertFile FileName:="vandungcao.doc", Range:="", _
        ConfirmConversions:=False, Link:=False, Attachment:=False
Selection.HomeKey Unit:=wdStory
End Sub
Private Sub Mucdo(ByVal a As String)
Dim ThisDoc As Document
Dim ThatDoc As Document
Application.ScreenUpdating = False
Selection.HomeKey Unit:=wdStory
Set ThisDoc = ActiveDocument
Call Add_EndCau
Selection.Find.ClearFormatting
With Selection.Find
.text = "(Câu [0-9]{1,2}[.:]*)(\[[0-2]{1}?[0-9]{1}?[0-9]{1}?[0-9]{1,2}?[1-4]{1}\])"
.Replacement.text = "\2\1\2"
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
With Selection.Find
.text = "(\[[0-2]{1}?[0-9]{1}?[0-9]{1}?[0-9]{1,2}?)" & a & "(\])(Câu [0-9]{1,2}*)(z.end)(^13)"
.Replacement.text = "\1" & a & "\2\3\4\5"
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
Call Del_EndCau
Selection.HomeKey Unit:=wdStory
Dim FileName, DocName
If a = 1 Then FileName = "D:\Tachtheomucdo\nhanbiet.doc"
If a = 2 Then FileName = "D:\Tachtheomucdo\thonghieu.doc"
If a = 3 Then FileName = "D:\Tachtheomucdo\vandungthap.doc"
If a = 4 Then FileName = "D:\Tachtheomucdo\vandungcao.doc"
ActiveDocument.SaveAs FileName
DocName = ActiveDocument.Name
ThatDoc.Close (yes)
ThisDoc.Activate
Call Del_EndCau
Application.ScreenUpdating = True
End Sub
Private Sub Add_EndCau()
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
.Replacement.text = "z.end^pCâu "
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
Selection.TypeText text:="z.end"
Selection.HomeKey Unit:=wdStory
With Selection.Find
.text = "z.end^p"
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
With Selection.Find
.text = "z.end"
.Replacement.text = "z.end"
.Forward = True
.Wrap = wdFindContinue
.Execute Replace:=wdReplaceAll
End With
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
End Sub
Sub Del_EndCau()
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
Selection.Find.ClearFormatting
With Selection.Find
.text = "(\[[0-2]{1}?[0-9]{1}?[0-9]{1}?[0-9]{1,2}?[1-4]{1}\])(Câu [0-9]{1,2}[.:]*)"
.Replacement.text = "\2"
.Forward = True
.Wrap = wdFindContinue
.MatchCase = True
.MatchWildcards = True
.Execute Replace:=wdReplaceAll
End With
Application.ScreenUpdating = True
Selection.HomeKey Unit:=wdStory
End Sub







