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
Private iItem As Long ' So luong phan tu trong mang
Attribute iItem.VB_VarDescription = "So luong phan tu trong mang"
'Khoi tao cac gia tri
Private Sub Class_Initialize()
    iItem = 0
    ReDim cPage(iItem)
End Sub
'Them mot tu moi
Public Function AddWord(ByVal vWord As clsWord) As Integer
On Error GoTo Result
Dim loca As clsLocation
    AddWord = eError.NoError
    Set loca = FindWord(0, vWord.WordCont)
    If Not (loca Is Nothing) And loca.ok <> 0 Then
        iItem = iItem + 1 ' tang so luong phan tu len mot
        ReDim Preserve cPage(iItem) ' thiet dat lai kich thuoc mang
        Set cPage(iItem) = New clsWord  ' gan gia tri tu moi vao vi tri cuoi mang
        If cPage(iItem).Copy(vWord) <> 0 Then GoTo Result
        If Sort <> 0 Then GoTo Result ' sap xep lai mang theo thu tu tang dan
    End If
    Exit Function
Result:
    AddWord = eError.AddWordError
End Function
' Tim mot tu trong mang CPage
Public Function FindWord(ByVal x As Long, ByVal Content As String) As clsLocation
Attribute FindWord.VB_Description = "Tim kiem theo nhi phan mot tu trong trang"
On Error GoTo Result
Dim Low As Long
Dim High As Long
Dim Mid As Long
Dim loca As New clsLocation
    Low = 1
    High = iItem
    Do While High >= Low ' tim kiem nhi phan
        Mid = (High + Low) \ 2
        Select Case SoSanh(Trim(cPage(Mid).WordCont), Content)
            Case 1: High = Mid - 1 ' lon hon
            Case -1: Low = Mid + 1 ' nho hon
            Case 0: Exit Do ' bang
        End Select
    Loop
    loca.x = x
    loca.y = Mid
    If High >= Low Then
        loca.ok = 0
    ElseIf Mid = High Then
        loca.ok = 1
    ElseIf Mid = Low Then
        loca.ok = -1
    End If
    Set FindWord = loca
    Exit Function
Result:
    FindWord = Nothing
End Function
'Xoa mot tu
Public Function DelWord(ByVal Content As String) As Integer
On Error GoTo Result
Dim i As Long
Dim loca As New clsLocation
    DelWord = eError.NoError
    Set loca = FindWord(0, Content) 'tim tu trong trang
    If Not (loca Is Nothing) And loca.ok = 0 Then 'co tu
        i = loca.y 'lay vi tri
        Do While i < iItem 'don cac tu khac len
            If cPage(i).Copy(cPage(i + 1)) <> 0 Then GoTo Result
            i = i + 1
        Loop
        iItem = iItem - 1 'giam so luong phan tu xuong mot
        ReDim Preserve cPage(iItem) 'khai bao lai kich thuoc mang
    Else
        DelWord = eError.NoHaveWord 'khong co tu
    End If
    Exit Function
Result:
    DelWord = eError.DelError
End Function
'Lay so luong tu trong trang
Public Property Get WCount() As Long
    WCount = iItem
End Property
' Dat tu loai vao tu o vi tri i
Public Property Let WordClass(ByVal i As Long, ByVal Class As String)
Attribute WordClass.VB_Description = "Tra ra hay lay ve tu loai cua tu thu i nao do trong mang"
    cPage(i).WordClass = Class
End Property
' Lay tu loai cua tu o vi tri i
Public Property Get WordClass(ByVal i As Long) As String
    WordClass = cPage(i).WordClass
End Property
'Dat tu vao vi tri i
Public Property Let WordCont(ByVal i As Long, ByVal Content As String)
Attribute WordCont.VB_Description = "Tra ra hay lay ve mot tu thu i trong mang"
    cPage(i).WordCont = Content
End Property
'Lay tu o vi tri i
Public Property Get WordCont(ByVal i As Long) As String
    WordCont = cPage(i).WordCont
End Property
' Ham sap xep mang CPage()
Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
    Sort = eError.NoError
    For i = 1 To iItem - 1
        For j = i + 1 To iItem
            If cPage(i).WordCont > cPage(j).WordCont Then ' kiem tra lon hon
                If Swap(i, j) <> 0 Then GoTo Result ' Hoan chuyen hai phan tu
            End If
        Next j
    Next i
    Exit Function
Result:
    Sort = eError.SortError
End Function
' Ham hoan chuyen hai gia tri trong mang CPage
Private Function Swap(i As Integer, j As Integer) As Integer
On Error GoTo Result
Dim temp As New clsWord
    If temp.Copy(cPage(i)) <> 0 Then GoTo Result
    If cPage(i).Copy(cPage(j)) <> 0 Then GoTo Result
    If cPage(j).Copy(temp) <> 0 Then GoTo Result
    Swap = eError.NoError
    Exit Function
Result:
    Swap = eError.SwapError
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
'Sao chep noi dung
Public Function Copy(ByVal p As clsCPage) As Integer
On Error GoTo Result
Dim i As Long
    Copy = eError.NoError
    iItem = p.WCount
    ReDim cPage(iItem)
    i = 1
    Do While i <= iItem
        Set cPage(i) = New clsWord
        cPage(i).WordCont = p.WordCont(i)
        cPage(i).WordClass = p.WordClass(i)
        i = i + 1
    Loop
    Exit Function
Result:
    Copy = eError.CopyCPageError
End Function
'Xoa noi dung cu
Public Function Clear() As Integer
On Error GoTo Result
    Clear = eError.NoError
    iItem = 0
    ReDim cPage(iItem)
    Exit Function
Result:
    Clear = eError.ClearCPageError
End Function
