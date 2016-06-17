VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Checking
'Description: clsCPage Class - Content Page Demonstration
'--------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private cPage() As clsWord ' Mang cac clsWord
Attribute cPage.VB_VarDescription = "Mot mang noi dung trang, luu tru cac tu trong mot trang"
Private cIPage() As clsIndex
Private iRitem As Long
Private iItem As Long ' So luong phan tu trong mang
Attribute iItem.VB_VarDescription = "So luong phan tu trong mang"
'Khoi tao cac gia tri
Private Sub Class_Initialize()
    iItem = 0
    iRitem = 0
    ReDim cPage(iItem)
    ReDim cIPage(iRitem)
End Sub
'Them mot tu moi
Public Function AddWord(ByVal vWord As clsWord) As Integer
On Error GoTo Result
Dim Loca As clsLocation
    AddWord = EError.NoError
    Set Loca = FindWord(0, vWord.WordCont)
    If Not (Loca Is Nothing) And Loca.ok <> 0 Then
        iItem = iItem + 1 ' tang so luong phan tu len mot
        iRitem = iRitem + 1
        ReDim Preserve cPage(iItem) ' thiet dat lai kich thuoc mang
        ReDim Preserve cIPage(iRitem)
        Set cPage(iItem) = New clsWord  ' gan gia tri tu moi vao vi tri cuoi mang
        Set cIPage(iRitem) = New clsIndex
        If cPage(iItem).Copy(vWord) <> 0 Then GoTo Result
        cIPage(iRitem).i = iItem
        If Sort <> 0 Then GoTo Result ' sap xep lai mang theo thu tu tang dan
    End If
    Exit Function
Result:
    AddWord = EError.AddWordError
End Function
' Tim mot tu trong mang CPage
Public Function FindWord(ByVal x As Long, ByVal Content As String) As clsLocation
Attribute FindWord.VB_Description = "Tim kiem theo nhi phan mot tu trong trang"
On Error GoTo Result
Dim Low As Long
Dim High As Long
Dim Mid As Long
Dim Loca As New clsLocation
    Low = 1
    High = iRitem
    Do While High >= Low ' tim kiem nhi phan
        Mid = (High + Low) \ 2
        Select Case SoSanh(Trim(cPage(cIPage(Mid).i).WordCont), Content)
            Case 1: High = Mid - 1 ' lon hon
            Case -1: Low = Mid + 1 ' nho hon
            Case 0: Exit Do ' bang
        End Select
    Loop
    Loca.x = x
    Loca.y = Mid
    If High >= Low Then
        Loca.ok = 0
    ElseIf Mid = High Then
        Loca.ok = 1
    ElseIf Mid = Low Then
        Loca.ok = -1
    End If
    Set FindWord = Loca
    Exit Function
Result:
    FindWord = Nothing
End Function
'Xoa mot tu
Public Function DelWord(ByVal Content As String) As Integer
On Error GoTo Result
Dim i As Long
Dim Loca As New clsLocation
    DelWord = EError.NoError
    Set Loca = FindWord(0, Content) 'tim tu trong trang
    If Not (Loca Is Nothing) And Loca.ok = 0 Then 'co tu
        i = Loca.y 'lay vi tri
        Do While i < iRitem 'don cac tu khac len
            cIPage(i).i = cIPage(i + 1).i
            i = i + 1
        Loop
        iRitem = iRitem - 1 'giam so luong phan tu xuong mot
        ReDim Preserve cIPage(iRitem)
    Else
        DelWord = EError.NoHaveWord 'khong co tu
    End If
    Exit Function
Result:
    DelWord = EError.DelError
End Function
'Lay so luong tu trong trang
Public Property Get WCount() As Long
    WCount = iRitem
End Property
' Dat tu loai vao tu o vi tri i
Public Property Let WordClass(ByVal i As Long, ByVal Class As String)
Attribute WordClass.VB_Description = "Tra ra hay lay ve tu loai cua tu thu i nao do trong mang"
    cPage(cIPage(i).i).WordClass = Class
End Property
' Lay tu loai cua tu o vi tri i
Public Property Get WordClass(ByVal i As Long) As String
    WordClass = cPage(cIPage(i).i).WordClass
End Property
'Dat tu vao vi tri i
Public Property Let WordCont(ByVal i As Long, ByVal Content As String)
Attribute WordCont.VB_Description = "Tra ra hay lay ve mot tu thu i trong mang"
    cPage(cIPage(i).i).WordCont = Content
End Property
'Lay tu o vi tri i
Public Property Get WordCont(ByVal i As Long) As String
    WordCont = cPage(cIPage(i).i).WordCont
End Property
' Ham sap xep mang CPage()
Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
Dim t As Long
    Sort = EError.NoError
    i = 1
    Do While (i < iRitem) And (cPage(cIPage(i).i).WordCont < cPage(cIPage(iItem).i).WordCont)
        i = i + 1
    Loop
    If i < iRitem Then
        t = cIPage(iRitem).i
        For j = iRitem To i + 1 Step -1
            cIPage(j).i = cIPage(j - 1).i
        Next j
        cIPage(i).i = t
    End If
    Exit Function
Result:
    Sort = EError.SortError
End Function
'So sanh hai string
Private Function SoSanh(ByVal st1 As String, ByVal st2 As String) As Integer
    If st1 > st2 Then
        SoSanh = 1
    ElseIf st1 < st2 Then
        SoSanh = -1
    Else
        SoSanh = 0
    End If
End Function
