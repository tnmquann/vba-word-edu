VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XoaHighlight 
   Caption         =   " "
   ClientHeight    =   2070
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7155
   OleObjectBlob   =   "XoaHighlight.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XoaHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    XoaHighlight.Hide
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "Câu "
        .Replacement.text = "z.zz^13Câu "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .Execute Replace:=wdReplaceAll
    End With
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText text:="z.zz"
    Selection.TypeParagraph
End Sub

Private Sub CommandButton2_Click()
    XoaHighlight.Hide
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "z.zz"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
