VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GanIDfrm 
   Caption         =   "GÁN MÃ CÂU"
   ClientHeight    =   1305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495.001
   OleObjectBlob   =   "GanIDfrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GanIDfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim phan, Chuong, Mucdo As Variant
Dim Bai As Integer
Private Sub cbbai_Change()
Bai = cbbai.listIndex + 1
End Sub

Private Sub cbchuong_Change()
cbbai.Clear
Chuong = cbchuong.listIndex + 1
Dim ArrD_0, ArrD_1, ArrD_2, ArrH_0, ArrH_1, ArrH_2 As Variant
ArrD_0 = Array("1.Menh de. Tap hop", "2.Ham so bac nhat va bac hai", "3.Phuong trinh. He phuong trinh", "4.Bat dang thuc. Bat phuong trinh", "5.Thong ke", "6.Cung va goc luong giac. Cong thuc luong giac")
ArrD_1 = Array("1.Ham so luong giac va phuong trinh luong giac", "2.To hop- Xac suat", "3.Day so- Cap so cong va cap so nhan", "4.Gioi han", "5.Dao ham")
ArrD_2 = Array("1.Ung dung dao ham de khao sat va ve do thi cua ham so", "2.Ham so luy thua. Ham so mu va ham so logarit", "3.Nguyen ham-Tich phan va ung dung", "4.So phuc")
ArrH_0 = Array("1.Cung va goc luong giac", "2.Gia tri luong giac cua mot cung", "3.Cong thuc luong giac")
ArrH_1 = Array("1.Phep doi hinh va phep dong dang trong mat phang", "2.Duong thang va mat phang trong khong gian. Quan he song song", "3.Vecto trong khong gian. Quan he vuong goc trong khong gian")
ArrH_2 = Array("1.Khoi da dien va the tich cua chung", "2.Mat cau, mat tru, mat non", "3.Hinh hoc Oxyz")
Select Case cbchuong.text
Case ArrD_0(0)
    cbbai.list = Array("1.Menh de", "2.Tap hop", "3.Cac phep toan tren tap hop", "4.Cac tap hop so", "5.So gan dung, sai so")
Case ArrD_0(1)
    cbbai.list = Array("1.Ham so", "2.Ham so y=ax+b", "3.Ham so bac hai")
Case ArrD_0(2)
    cbbai.list = Array("1.Dai cuong ve phuong trinh", "2.Phuong trinh quy ve bac nhat, bac hai", "3.Phuong trinh va he phuong trinh bac nhat nhieu an")
Case ArrD_0(3)
    cbbai.list = Array("1.Bat dang thuc", "2.Bat phuong trinh va he bat phuong trinh", "3.Dau cua nhi thuc bac nhat", "4.Dau cua tam thuc bac hai")
Case ArrD_0(4)
    cbbai.list = Array("1.Bang phan bo tan suat, tan suat", "2.Bieu do", "3.So trung binh cong. So trung vi. Mot", "4.Phuong sai va do lech chuan")
Case ArrD_0(5)
    cbbai.list = Array("1.Cung va goc luong giac", "2.Gia tri luong giac cua mot cung", "3.Cong thuc luong giac")
Case ArrH_0(0)
    cbbai.list = Array("1.Cac dinh nghia", "2.Tong va hieu hai vecto", "3.Tich cua vecto va mot so", "4.He truc toa do")
Case ArrH_0(1)
    cbbai.list = Array("1.Gia tri luong giac cua mot goc bat ki tu 0 den 180", "2.Tich vo huong cua hai vecto", "3.Cac he thuc luong trong tam giac va giai tam giac")
Case ArrH_0(2)
    cbbai.list = Array("1.Phuong trinh duong thang", "2.Phuong trinh duong tron", "3.Phuong trinh duong elip")
Case ArrD_1(0)
    cbbai.list = Array("1.Ham so luong giac", "2.Phuong trinh luong giac co ban", "3.Mot so phuong trinh luong giac thuong gap", "4.Cac phuong phap dua ve phuong trinh luong giac co ban", "5.Phuong trinh luong giac voi tap nghiem bi chan", "Phuong trinh luong giac chua tham so (khong dung phuong phap ham so)")
Case ArrD_1(1)
    cbbai.list = Array("1.Quy tac dem", "2.Cac bai toan Hoan vi-Chinh hop-To hop", "3.Nhi thuc Newton", "4.Phep thu va bien co", "5.Xac suat cua bien co")
Case ArrD_1(2)
    cbbai.list = Array("1.Phuong phap quy nap toan hoc", "2.Day so", "3.Cap so cong", "4.Cap so nhan")
Case ArrD_1(3)
    cbbai.list = Array("1.Gioi han cua day so", "2.Gioi han cua ham so", "3.Ham so lien tuc")
Case ArrD_1(4)
    cbbai.list = Array("1.Dinh nghia va y nghia cua dao ham", "2.Quy tac tinh dao ham", "3.Dao ham cua ham so luong giac", "4.Vi phan", "5.Dao ham cap 2")
Case ArrH_1(0)
    cbbai.list = Array("1.Phep tinh tien", "2.Phep doi xung truc", "3.Phep doi xung tam", "4.Phep quay", "5.Phep vi tu", "6.Phep dong dang", "7.Ung dung phep bien hinh de giai toan hinh hoc phang")
Case ArrH_1(1)
    cbbai.list = Array("1.Dai cuong ve duong thang va mat phang", "2.Hai duong thang cheo nhau va hai duong thang song song", "3.Duong thang va mat phang song song", "4.Hai mat phang song song", "5.Phep chieu song song. Hinh bieu dien cua mot hinh khong gian")
Case ArrH_1(2)
    cbbai.list = Array("1.Vecto trong khong gian", "2.Hai duong thang vuong goc", "3.Duong thang vuong goc voi mat phang", "4.Hai mat phang vuong goc", "5.Goc", "6.Khoang cach")
Case ArrD_2(0)
    cbbai.list = Array("1.Su dong bien va nghich bien cua ham so", "2.Cuc tri cua ham so", "3.Gia tri lon nhat va gia tri nho nhat cua ham so", "4.Duong tiem can", "5.Khao sat su bien thien va ve do thi cua ham so-Phep bien doi do thi", "6.Su tuong giao cua hai do thi", "7.Su tiep xuc. Tiep tuyen cua do thi ham so", "8.Cac bai toan ve diem, khoang cach, dien tich")
Case ArrD_2(1)
    cbbai.list = Array("1.Luy thua-ham so luy thua", "2.Logarit", "3.Ham so logarit-Ham so mu", "4.Phuong trinh mu", "5.Phuong trinh logarit", "6.Bat phuong trinh mu", "7.Bat phuong trinh logarit", "8.Cac bai toan thuc te")
Case ArrD_2(2)
    cbbai.list = Array("1.Nguyen ham", "2.Tich phan", "3.Ung dung cua tich phan trong tinh dien tich hinh phang", "4.Ung dung cua tich phan trong tinh dien tich khoi tron xoay")
Case ArrD_2(3)
    cbbai.list = Array("1.Cac phep toan tren tap so phuc", "2.Phuong trinh", "3.Tap hop diem bieu dien so phuc", "4.Cac bai toan cuc tri")
Case ArrH_2(0)
    cbbai.list = Array("1.Khai niem ve khoi da dien", "2.The tich khoi chop", "3.The tich khoi lang tru", "4.Cac bai toan ti le", "5.Cac bai toan thuc te")
Case ArrH_2(1)
    cbbai.list = Array("1.Hinh non", "2.Hinh tru", "3.Mat cau", "4.Cac bai toan tong hop hinh non, tru, cau", "5.Cac bai toan thuc te")
Case ArrH_2(2)
    cbbai.list = Array("1.He truc toa do trong khong gian", "2.Phuong trinh mat phang", "3.Phuong trinh duong thang", "4.Mat cau", "5.Cac bai toan tong hop ", "5.Bai toan ve cuc tri hinh hoc")
End Select
End Sub



Private Sub cbmucdo_Change()
Select Case cbmucdo.text
    Case cbmucdo.list(0)
        Mucdo = "1"
    Case cbmucdo.list(1)
        Mucdo = "2"
    Case cbmucdo.list(2)
        Mucdo = "3"
    Case cbmucdo.list(3)
        Mucdo = "4"
End Select
End Sub

Private Sub CommandButton1_Click()
Dim ID As String
ID = "[" & phan & Chuong & "-" & Bai & "-" & Mucdo & "] "
Selection.text = ID
Selection.Font.ColorIndex = wdRed
Selection.Font.Bold = True
SendKeys "^v"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()
cbphan.list = Array("Dai 10", "Dai 11", "Dai 12", "Hinh 10", "Hinh 11", "Hinh 12")
cbmucdo.list = Array("MD1", "MD2", "MD3", "MD4")
End Sub

Private Sub cbphan_Change()
Select Case cbphan.text
    Case "Dai 10"
        phan = "0D"
    Case "Dai 11"
        phan = "1D"
    Case "Dai 12"
        phan = "2D"
    Case "Hinh 10"
        phan = "0H"
    Case "Hinh 11"
        phan = "1H"
    Case "Hinh 12"
        phan = "2H"
End Select
Arrphan_0 = Array("Dai 10", "Dai 11", "Dai 12", "Hinh 10", "Hinh 11", "Hinh 12")
cbchuong.Clear
        Select Case phan
            Case "0D"
                cbchuong.list = Array("1.Menh de. Tap hop", "2.Ham so bac nhat va bac hai", "3.Phuong trinh. He phuong trinh", "4.Bat dang thuc. Bat phuong trinh", "5.Thong ke", "6.Cung va goc luong giac. Cong thuc luong giac")
            Case "1D"
                cbchuong.list = Array("1.Ham so luong giac va phuong trinh luong giac", "2.To hop- Xac suat", "3.Day so- Cap so cong va cap so nhan", "4.Gioi han", "5.Dao ham")
            Case "2D"
                cbchuong.list = Array("1.Ung dung dao ham de khao sat va ve do thi cua ham so", "2.Ham so luy thua. Ham so mu va ham so logarit", "3.Nguyen ham-Tich phan va ung dung", "4.So phuc")
            Case "0H"
                cbchuong.list = Array("1.Cung va goc luong giac", "2.Gia tri luong giac cua mot cung", "3.Cong thuc luong giac")
            Case "1H"
                cbchuong.list = Array("1.Phep doi hinh va phep dong dang trong mat phang", "2.Duong thang va mat phang trong khong gian. Quan he song song", "3.Vecto trong khong gian. Quan he vuong goc trong khong gian")
            Case "2H"
                cbchuong.list = Array("1.Khoi da dien va the tich cua chung", "2.Mat cau, mat tru, mat non", "3.Hinh hoc Oxyz")
    End Select

End Sub


