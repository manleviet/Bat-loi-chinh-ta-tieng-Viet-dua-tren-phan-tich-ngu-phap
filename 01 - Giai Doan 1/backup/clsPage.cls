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
Private ipage() As clsIndex
Private iRitem As Long
Private iItem As Long  ' So luong trang
Attribute iItem.VB_VarDescription = "So luong trang"
'Khoi tao cac gia tri
Private Sub Class_Initialize()
    iItem = 0
    iRitem = 0
    ReDim page(iItem)
    ReDim ipage(iRitem)
End Sub
'Them mot tu vao tu dien
Public Function AddWord(ByVal i As Long, ByVal vWord As clsWord) As Integer
    AddWord = page(ipage(i).i).AddWord(vWord)
End Function
'Tim kiem mot tu trong tu dien
Public Function FindWord(ByVal i As Long, ByVal Content As String) As clsLocation
Attribute FindWord.VB_Description = "Tim kiem nhi phan mot tu nao do trong trang"
Dim Loca As New clsLocation
    Set Loca = page(ipage(i).i).FindWord(i, Content)
    Set FindWord = Loca
End Function
'Xoa mot tu trong tu dien
Public Function DelWord(ByVal i As Long, ByVal Content As String) As Integer
On Error GoTo Result
    DelWord = 0
    If page(ipage(i).i).WCount = 1 Then
        Do While i < iRitem
            ipage(i).i = ipage(i + 1).i
            i = i + 1
        Loop
        iRitem = iRitem - 1
        ReDim Preserve ipage(iRitem)
    Else
        DelWord = page(ipage(i).i).DelWord(Content)
    End If
    Exit Function
Result:
    DelWord = 1000
End Function
'Tong so tu trong tu dien
Public Property Get WCount() As Long
Dim i As Long
    i = 1
    WCount = 0
    Do While i <= iRitem ' duyet qua tung trang de tinh tong
        WCount = WCount + page(ipage(i).i).WCount
        i = i + 1
    Loop
End Property
'So tu trong trang i
Public Property Get WiPCount(ByVal i As Integer) As Long
    WiPCount = page(ipage(i).i).WCount
End Property
'Dat tu loai vao tu o vi tri i, j
Public Property Let WordClass(ByVal i As Long, ByVal j As Long, ByVal Class As String)
Attribute WordClass.VB_Description = "Lay ra hay dua vao tu loai tai vi tri i, j"
    page(ipage(i).i).WordClass(j) = Class
End Property
'Lay tu loai cua tu o vi tri i, j
Public Property Get WordClass(ByVal i As Long, ByVal j As Long) As String
    WordClass = page(ipage(i).i).WordClass(j)
End Property
'Dat tu vao vi tri i, j
Public Property Let WordCont(ByVal i As Long, ByVal j As Long, ByVal Content As String)
Attribute WordCont.VB_Description = "Lay ra hay dua vao mot tu tai vi tri i, j"
    page(ipage(i).i).WordCont(j) = Content
End Property
'Lay tu o vi tri i, j
Public Property Get WordCont(ByVal i As Long, ByVal j As Long) As String
    WordCont = page(ipage(i).i).WordCont(j)
End Property
'Them mot trang vao tu dien
Public Function AddPage(ByVal i As Long, ByVal vWord As clsWord) As Integer
On Error GoTo Result
Dim j As Long
    AddPage = EError.NoError
    iItem = iItem + 1 ' tang so luong trang
    iRitem = iRitem + 1
    ReDim Preserve page(iItem) ' dat lai kich thuoc mang trang
    ReDim Preserve ipage(iRitem)
    Set page(iItem) = New clsCPage
    Set ipage(iRitem) = New clsIndex
    j = iItem
    Do While j > i
        ipage(j).i = ipage(j - 1).i
        j = j - 1
    Loop
    ipage(i).i = iItem
    If page(ipage(i).i).AddWord(vWord) <> 0 Then
        GoTo Result
    End If
    Exit Function
Result:
    AddPage = EError.AddPageError
End Function
