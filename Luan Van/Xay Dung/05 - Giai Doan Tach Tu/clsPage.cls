VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Checking
'Description: clsPage Class - Word Page Demonstration
'------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private page() As clsCPage ' Mang cac clsCPage
Attribute page.VB_VarDescription = "Mang cac trang"
Private iItem As Integer ' So luong trang
Attribute iItem.VB_VarDescription = "So luong trang"
'Khoi tao cac gia tri
Private Sub Class_Initialize()
    iItem = 0
    ReDim page(iItem)
End Sub
'Them mot tu vao tu dien
Public Function AddWord(ByVal i As Long, ByVal vWord As clsWord) As Integer
    AddWord = page(i).AddWord(vWord)
End Function
'Tim kiem mot tu trong tu dien
Public Function FindWord(ByVal i As Long, ByVal Content As String) As clsLocation
Attribute FindWord.VB_Description = "Tim kiem nhi phan mot tu nao do trong trang"
Dim Loca As New clsLocation
    Set Loca = page(i).FindWord(i, Content)
    Set FindWord = Loca
End Function
'Xoa mot tu trong tu dien
Public Function DelWord(ByVal i As Long, ByVal Content As String) As Integer
    DelWord = page(i).DelWord(Content)
End Function
'Tong so tu trong tu dien
Public Property Get WCount() As Long
Dim i As Long
    i = 1
    WCount = 0
    Do While i <= iItem ' duyet qua tung trang de tinh tong
        WCount = WCount + page(i).WCount
        i = i + 1
    Loop
End Property
'So tu trong trang i
Public Property Get WiPCount(ByVal i As Integer) As Long
    WiPCount = page(i).WCount
End Property
'Lay tu va tu loai tai vi tri i, j
'Public Function GetWord(ByVal i As Long, ByVal j As Long, vWord As clsWord) As Integer
'    GetWord = page(i).GetWord(j, vWord)
'End Function
'Dat tu va tu loai vao vi tri i, j
'Public Function SetWord(ByVal i As Long, ByVal j As Long, ByVal vWord As clsWord) As Integer
'    SetWord = page(i).SetWord(j, vWord)
'End Function
'Dat tu loai vao tu o vi tri i, j
Public Property Let WordClass(ByVal i As Long, ByVal j As Long, ByVal Class As String)
Attribute WordClass.VB_Description = "Lay ra hay dua vao tu loai tai vi tri i, j"
    page(i).WordClass(j) = Class
End Property
'Lay tu loai cua tu o vi tri i, j
Public Property Get WordClass(ByVal i As Long, ByVal j As Long) As String
    WordClass = page(i).WordClass(j)
End Property
'Dat tu vao vi tri i, j
Public Property Let WordCont(ByVal i As Long, ByVal j As Long, ByVal Content As String)
Attribute WordCont.VB_Description = "Lay ra hay dua vao mot tu tai vi tri i, j"
    page(i).WordCont(j) = Content
End Property
'Lay tu o vi tri i, j
Public Property Get WordCont(ByVal i As Long, ByVal j As Long) As String
    WordCont = page(i).WordCont(j)
End Property
'Them mot trang vao tu dien
Public Function AddPage(ByVal i As Long, ByVal vWord As clsWord) As Integer
On Error GoTo Result
Dim j As Long
    AddPage = EError.NoError
    iItem = iItem + 1 ' tang so luong trang
    ReDim Preserve page(iItem) ' dat lai kich thuoc mang trang
    Set page(iItem) = New clsCPage
    j = iItem
    Do While j > i
        If page(j).Copy(page(j - 1)) <> 0 Then GoTo Result
        j = j - 1
    Loop
    If page(i).Clear <> 0 Then GoTo Result
    If page(i).AddWord(vWord) <> 0 Then GoTo Result
    Exit Function
Result:
    AddPage = EError.AddPageError
End Function
