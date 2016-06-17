VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Chechking
'Description: clsWordDic Class - Word Dictionary Demonstration
'------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences, IT Department

Option Explicit
Private page As clsPage ' Mang tu dien hai chieu
Private IPage() As clsIPage ' Chi muc trang
Attribute IPage.VB_VarDescription = "Mang cac tu chi muc"
Private iItem As Integer ' So chi muc trang
Attribute iItem.VB_VarDescription = "So luong tu chi muc"
'Khoi tao cac gia tri
Private Sub Class_Initialize()
    iItem = 0
    ReDim IPage(iItem)
    Set page = New clsPage
End Sub
'Them mot tu vao tu dien tu
Public Function AddWord(ByVal vWord As clsWord) As Integer
On Error GoTo Result
Dim fWord As String
Dim loca As New clsLocation
    AddWord = eError.NoError
    fWord = FirstWord(vWord.WordCont)  ' cat tu dau trong content
    Set loca = FindSWord(fWord)
    If loca Is Nothing Then GoTo Result
    If loca.ok = 0 Then
        If page.AddWord(loca.x, vWord) <> 0 Then GoTo Result
    Else
        iItem = iItem + 1 ' tang so luong trang len mot
        ReDim Preserve IPage(iItem) ' thiet dat lai kich thuoc mang
        Set IPage(iItem) = New clsIPage
        IPage(iItem).WordCont = fWord ' dat tu vao cuoi mang
        If Sort <> 0 Then GoTo Result ' sap xep lai mang
        Set loca = FindSWord(fWord) ' lay vi tri hien tai cua tu can tim
        'them mot trang trong page
        If loca.ok = 0 Then
            If page.AddPage(loca.x, vWord) <> 0 Then GoTo Result
        Else
            GoTo Result
        End If
    End If
    Exit Function
Result:
    AddWord = eError.AddWordError
End Function
' Tim tu don trong IPage
Public Function FindSWord(ByVal Content As String) As clsLocation
Attribute FindSWord.VB_Description = "Tim kiem tu don"
On Error GoTo Result
Dim Low As Long
Dim High As Long
Dim Mid As Long
Dim loca As New clsLocation
    Low = 1
    High = iItem
    Do While High >= Low ' tim kiem nhi phan
        Mid = (High + Low) \ 2
        Select Case SoSanh(Trim(IPage(Mid).WordCont), Content)
            Case 1: High = Mid - 1 ' lon hon
            Case -1: Low = Mid + 1 ' nho hon
            Case 0: Exit Do ' bang
        End Select
    Loop
    loca.x = Mid
    If High >= Low Then
        loca.ok = 0
    ElseIf Mid = High Then
        loca.ok = 1
    ElseIf Mid = Low Then
        loca.ok = -1
    End If
    Set FindSWord = loca
    Exit Function
Result:
    FindSWord = Nothing
End Function
' Tim tu ghep va tu don trong CPage
Public Function FindWord(ByVal Content As String) As clsLocation
Attribute FindWord.VB_Description = "Tim kiem tu ghep va don"
On Error GoTo Result
Dim fWord As String
Dim loca As New clsLocation
    fWord = FirstWord(Content)
    Set loca = FindSWord(fWord)
    If Not (loca Is Nothing) And loca.ok = 0 Then
        Set loca = page.FindWord(loca.x, Content)
    End If
    Set FindWord = loca
    Exit Function
Result:
    FindWord = Nothing
End Function
' Xoa mot tu trong tu dien thong qua Content
Public Function DelWord(ByVal Content As String) As Integer
Attribute DelWord.VB_Description = "Xoa mot tu khoi tu dien"
On Error GoTo Result
Dim fWord As String
Dim loca As New clsLocation
    DelWord = eError.NoError
    fWord = FirstWord(Content)
    Set loca = FindSWord(fWord)
    If Not (loca Is Nothing) And loca.ok = 0 Then
        If page.WiPCount(loca.x) = 1 Then Exit Function
        If page.DelWord(loca.x, Content) <> 0 Then GoTo Result
    Else
        DelWord = eError.NoHaveWord
    End If
    Exit Function
Result:
    DelWord = eError.DelError
End Function
'So trang trong tu dien
Public Property Get PCount() As Long
    PCount = iItem
End Property
'Tong so tu trong tu dien
Public Property Get WCount() As Long
Attribute WCount.VB_Description = "So tu trong tu dien"
    WCount = page.WCount
End Property
'So tu trong 1 trang nao do
Public Property Get WiPCount(ByVal i As Integer) As Long
    WiPCount = page.WiPCount(i)
End Property
'Dat gia tri tu loai vao cho tu o vi tri i, j
Public Property Let WordClass(ByVal i As Long, ByVal j As Long, ByVal Class As String)
Attribute WordClass.VB_Description = "Lay ra hay dua vao tu loai tai vi tri i, j"
    page.WordClass(i, j) = Class
End Property
'Lay ra tu loai cua tu tai vi tri i, j
Public Property Get WordClass(ByVal i As Long, ByVal j As Long) As String
    WordClass = page.WordClass(i, j)
End Property
'Dat gia tri tu vao vi tri i, j
Public Property Let WordCont(ByVal i As Long, ByVal j As Long, ByVal Content As String)
Attribute WordCont.VB_Description = "Lay ra hay dua vao mot tu tai vi tri i, j"
    page.WordCont(i, j) = Content
End Property
'Lay ra mot tu tai vi tri i, j
Public Property Get WordCont(ByVal i As Long, ByVal j As Long) As String
    WordCont = page.WordCont(i, j)
End Property
'Lay du lieu tu dien tu file
Public Function LoadDic() As Integer
On Error GoTo Result
Dim FH As tFILEHEADER, ub As Long
Dim FileNum As Integer, i As Integer, j As Long, p As Long
Dim cWord As New clsWord
Dim bPage() As Integer
Dim bWord() As Integer
Dim bType() As Integer
Dim iIPage As Long
Dim bdau As Long
    LoadDic = eError.NoError
    'load file iPage.vch
    FileNum = FreeFile
    Open fIPage For Binary As #FileNum
    If FileLen(fIPage) <> 0 Then
        'Doc FILEHEADER
        bdau = 1
        Get #FileNum, bdau, FH
        bdau = bdau + 8
        iItem = FH.iSize
        ReDim IPage(iItem)
        'Doc cac am tiet
        For i = 1 To iItem
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bPage(ub)
            Get #FileNum, bdau, bPage
            bdau = bdau + ((ub + 1) * 2)
            Set IPage(i) = New clsIPage
            IPage(i).WordCont = UTF16ToUniStr(bPage)
        Next i
    End If
    Close #FileNum
    'load file worddic.vch
    FileNum = FreeFile
    Open fWordDic For Binary As #FileNum
    bdau = 1
    If FileLen(fWordDic) <> 0 Then
        Get #FileNum, bdau, FH
        bdau = bdau + 8
        j = FH.iSize
        p = 0
        For i = 1 To j
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bWord(ub)
            Get #FileNum, bdau, bWord
            bdau = bdau + ((ub + 1) * 2)
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bType(ub)
            Get #FileNum, bdau, bType
            bdau = bdau + ((ub + 1) * 2)
            Get #FileNum, bdau, iIPage
            bdau = bdau + 4
            cWord.WordCont = UTF16ToUniStr(bWord)
            cWord.WordClass = UTF16ToUniStr(bType)
            If iIPage <> p Then
                If page.AddPage(iIPage, cWord) <> 0 Then GoTo Result
                p = iIPage
            Else
                If page.AddWord(iIPage, cWord) <> 0 Then GoTo Result
            End If
        Next i
    End If
    Close #FileNum
    Exit Function
Result:
    LoadDic = eError.LoadDicError
End Function
'Ghi du lieu tu dien ra file
Public Function SaveDic() As Integer
Dim FH As tFILEHEADER
Dim FileNum As Integer, i As Long, j As Long
Dim bPage() As Integer
Dim bWord() As Integer
Dim bType() As Integer
Dim iIPage As Long
Dim bdau As Long
    SaveDic = eError.NoError
    ' Save ipage.vch
    FileNum = FreeFile
    Open fIPage For Binary As #FileNum
    bdau = 1
    If iItem <> 0 Then
        FH.iSize = iItem
        FH.iType = 0
        Put #FileNum, bdau, FH
        bdau = bdau + 8
        For i = 1 To iItem
            bPage = UniStrToUTF16(Trim(IPage(i).WordCont))
            Put #FileNum, bdau, UBound(bPage)
            bdau = bdau + 4
            Put #FileNum, bdau, bPage
            bdau = bdau + ((UBound(bPage) + 1) * 2)
        Next i
    End If
    Close #FileNum
    'Save worddic.vch
    FileNum = FreeFile
    Open fWordDic For Binary As #FileNum
    bdau = 1
    If Me.WCount <> 0 Then
        FH.iSize = Me.WCount
        FH.iType = 1
        Put #FileNum, bdau, FH
        bdau = bdau + 8
        For i = 1 To iItem
            For j = 1 To page.WiPCount(i)
                bWord = UniStrToUTF16(Trim(Me.WordCont(i, j)))
                bType = UniStrToUTF16(Trim(Me.WordClass(i, j)))
                iIPage = i
                Put #FileNum, bdau, UBound(bWord)
                bdau = bdau + 4
                Put #FileNum, bdau, bWord
                bdau = bdau + ((UBound(bWord) + 1) * 2)
                Put #FileNum, bdau, UBound(bType)
                bdau = bdau + 4
                Put #FileNum, bdau, bType
                bdau = bdau + (UBound(bType) + 1) * 2
                Put #FileNum, bdau, iIPage
                bdau = bdau + 4
            Next j
        Next i
    End If
    Close #FileNum
    Exit Function
Result:
    SaveDic = eError.SaveDicError
End Function
' Ham sap xep mang IPage()
Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
    Sort = eError.NoError
    For i = 1 To iItem - 1
        For j = i + 1 To iItem
            If IPage(i).WordCont > IPage(j).WordCont Then   ' kiem tra lon hon
                If Swap(i, j) <> 0 Then GoTo Result ' Hoan chuyen hai phan tu
            End If
        Next j
    Next i
    Exit Function
Result:
    Sort = eError.SortError
End Function
' Ham hoan chuyen hai gia tri trong mang IPage
Private Function Swap(i As Integer, j As Integer) As Integer
On Error GoTo Result
Dim temp As New clsIPage
    temp.WordCont = IPage(i).WordCont
    IPage(i).WordCont = IPage(j).WordCont
    IPage(j).WordCont = temp.WordCont
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
