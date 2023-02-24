Attribute VB_Name = "PS_GanID1"
Option Explicit
Public Sub Gan_ID5(ByVal control As Office.IRibbonControl)
    GanID5_Word.Show
End Sub

Function prepare_listVDC(ByVal table_index As Integer, ByVal cot As Long, ByVal cot1 As String, ByVal cot2 As String)
'    ham tra ve mang cac du lieu duy nhat lay tu bang table_index, tu cot col
'    table_index: chi so cua bang trong tap tin - tinh tu 1
'    cot: chi so cot can lay du lieu - tinh tu 1
Dim r As Long, text As String, tabl As Table, dic As Object
    On Error Resume Next
    If cot < 1 Or cot > 3 Then Exit Function
    Set dic = CreateObject("Scripting.Dictionary")
    Set tabl = ThisDocument.Tables.item(table_index)
    For r = 2 To tabl.Rows.Count
        text = Trim(Replace(Replace(Replace(tabl.Cell(r, 1).Range.text, Chr(7), ""), Chr(0), ""), Chr(13), ""))
        If cot > 1 Then
            If LCase(text) = LCase(cot1) Then
                text = Trim(Replace(Replace(Replace(tabl.Cell(r, 2).Range.text, Chr(7), ""), Chr(0), ""), Chr(13), ""))
                If cot > 2 Then
                    If LCase(text) = LCase(cot2) Then
                        text = Trim(Replace(Replace(Replace(tabl.Cell(r, 3).Range.text, Chr(7), ""), Chr(0), ""), Chr(13), ""))
                    Else
                        text = ""
                    End If
                End If
            Else
                text = ""
            End If
        End If
        If text <> "" And Not dic.Exists(text) Then dic.add text, 0
    Next r
    If dic.Count Then prepare_listVDC = dic.keys()
    
    Set tabl = Nothing
    Set dic = Nothing
End Function
