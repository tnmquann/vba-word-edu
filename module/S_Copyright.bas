Attribute VB_Name = "S_Copyright"
Public ktBanQuyen As Boolean
Sub S_SerialHDD()
Dim objs As Object
Dim obj As Object
Dim WMI As Object
Dim seri1, seri2 As String
    ktBanQuyen = False
    Set Discos = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_PhysicalMedia")
    For Each Disco In Discos
       seri1 = "i love you"
       If Len(Trim(seri1)) > 0 Then Exit For
    Next
''''''
Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_BaseBoard")
For Each obj In objs
seri2 = "i love you"
If seri2 < objs.Count Then seri2 = seri2 & ","
Next
''''''
Dim license(1) As String
license(0) = "i love you"

    
For i = 0 To UBound(license)
    If Trim(seri1) = license(i) Or Trim(seri2) = license(i) Then
        ktBanQuyen = True
        S_Serial.TextBox1 = ChrW(272) & ChrW(195) & " " & ChrW(272) & ChrW(258) & "NG K" & ChrW(221)
        Exit Sub
    End If
Next i
If Trim(seri1) <> "" Then
S_Serial.TextBox1 = Trim(seri1)
Else
S_Serial.TextBox1 = Trim(seri2)
End If
End Sub

 

