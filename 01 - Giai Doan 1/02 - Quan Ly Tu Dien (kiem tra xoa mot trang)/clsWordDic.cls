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
Private ipage() As clsIPage ' Chi muc trang
Attribute ipage.VB_VarDescription = "Mang cac tu chi muc"
Private index() As clsIndex
Private iRitem As Long
Private iItem As Long  ' So chi muc trang
Attribute iItem.VB_VarDescription = "So luong tu chi muc"
'Khoi tao cac gia tri
Private Sub Class_Initialize()
    iItem = 0
    iRitem = 0
    ReDim ipage(iItem)
    ReDim index(iRitem)
    Set page = New clsPage
End Sub
'Them mot tu vao tu dien tu
Public Function AddWord(ByVal vWord As clsWord) As Integer
On Error GoTo Result
Dim fWord As String
Dim Loca As New clsLocation
    AddWord = EError.NoError
    fWord = FirstWord(vWord.WordCont)  ' cat tu dau trong content
    Set Loca = FindSWord(fWord)
    If Loca Is Nothing Then GoTo Result
    If Loca.ok = 0 Then
        AddWord = page.AddWord(index(Loca.x).i, vWord)
    Else
        iItem = iItem + 1 ' tang so luong trang len mot
        iRitem = iRitem + 1
        ReDim Preserve ipage(iItem) ' thiet dat lai kich thuoc mang
        ReDim Preserve index(iRitem)
        Set ipage(iItem) = New clsIPage
        Set index(iRitem) = New clsIndex
        ipage(iItem).WordCont = fWord ' dat tu vao cuoi mang
        index(iRitem).i = iItem
        AddWord = Sort 'sap xep lai mang
        If AddWord <> 0 Then Exit Function
        Set Loca = FindSWord(fWord) ' lay vi tri hien tai cua tu can tim
        'them mot trang trong page
        If Loca.ok = 0 Then
            AddWord = page.AddPage(index(Loca.x).i, vWord)
        Else
            GoTo Result
        End If
    End If
    Exit Function
Result:
    AddWord = EError.AddWordError
End Function
' Tim tu don trong IPage
Public Function FindSWord(ByVal Content As String) As clsLocation
Attribute FindSWord.VB_Description = "Tim kiem tu don"
On Error GoTo Result
Dim Low As Long
Dim High As Long
Dim Mid As Long
Dim Loca As New clsLocation
    Low = 1
    High = iRitem
    Do While High >= Low ' tim kiem nhi phan
        Mid = (High + Low) \ 2
        Select Case SoSanh(Trim(ipage(index(Mid).i).WordCont), Content)
            Case 1: High = Mid - 1 ' lon hon
            Case -1: Low = Mid + 1 ' nho hon
            Case 0: Exit Do ' bang
        End Select
    Loop
    Loca.x = Mid
    If High >= Low Then
        Loca.ok = 0
    ElseIf Mid = High Then
        Loca.ok = 1
    ElseIf Mid = Low Then
        Loca.ok = -1
    End If
    Set FindSWord = Loca
    Exit Function
Result:
    FindSWord = Nothing
End Function
' Tim tu ghep va tu don trong CPage
Public Function FindWord(ByVal Content As String) As clsLocation
Attribute FindWord.VB_Description = "Tim kiem tu ghep va don"
On Error GoTo Result
Dim fWord As String
Dim Loca As New clsLocation
    fWord = FirstWord(Content)
    Set Loca = FindSWord(fWord)
    If Not (Loca Is Nothing) And Loca.ok = 0 Then
        Set Loca = page.FindWord(index(Loca.x).i, Content)
    End If
    Set FindWord = Loca
    Exit Function
Result:
    FindWord = Nothing
End Function
' Xoa mot tu trong tu dien thong qua Content
Public Function DelWord(ByVal Content As String) As Integer
Attribute DelWord.VB_Description = "Xoa mot tu khoi tu dien"
On Error GoTo Result
Dim fWord As String, i As Long
Dim Loca As New clsLocation
    DelWord = EError.NoError
    fWord = FirstWord(Content)
    Set Loca = FindSWord(fWord)
    If Not (Loca Is Nothing) And Loca.ok = 0 Then
        If page.WiPCount(index(Loca.x).i) = 1 Then
            'DelWord = page.DelWord(Loca.x, Content)
            'If DelWord <> 0 Then Exit Function
            i = Loca.x
            Do While i < iRitem
                index(i).i = index(i + 1).i
                i = i + 1
            Loop
            iRitem = iRitem - 1
            ReDim Preserve index(iRitem)
        Else
            DelWord = page.DelWord(index(Loca.x).i, Content)
            If DelWord <> 0 Then Exit Function
        End If
    Else
        DelWord = EError.NoHaveWord
    End If
    Exit Function
Result:
    DelWord = EError.DelError
End Function
'So trang trong tu dien
Public Property Get PCount() As Long
    PCount = iRitem
End Property
'Tong so tu trong tu dien
Public Property Get WCount() As Long
Attribute WCount.VB_Description = "So tu trong tu dien"
    WCount = page.WCount
End Property
'So tu trong 1 trang nao do
Public Property Get WiPCount(ByVal i As Integer) As Long
    WiPCount = page.WiPCount(index(i).i)
End Property
'Dat gia tri tu loai vao cho tu o vi tri i, j
Public Property Let WordClass(ByVal i As Long, ByVal j As Long, ByVal Class As String)
Attribute WordClass.VB_Description = "Lay ra hay dua vao tu loai tai vi tri i, j"
    page.WordClass(index(i).i, j) = Class
End Property
'Lay ra tu loai cua tu tai vi tri i, j
Public Property Get WordClass(ByVal i As Long, ByVal j As Long) As String
    WordClass = page.WordClass(index(i).i, j)
End Property
'Dat gia tri tu vao vi tri i, j
Public Property Let WordCont(ByVal i As Long, ByVal j As Long, ByVal Content As String)
Attribute WordCont.VB_Description = "Lay ra hay dua vao mot tu tai vi tri i, j"
    page.WordCont(index(i).i, j) = Content
End Property
'Lay ra mot tu tai vi tri i, j
Public Property Get WordCont(ByVal i As Long, ByVal j As Long) As String
    WordCont = page.WordCont(index(i).i, j)
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
Dim iipage As Long
Dim bdau As Long
    LoadDic = EError.NoError
    'load file iPage.vch
    FileNum = FreeFile
    Open fIPage For Binary As #FileNum
    If FileLen(fIPage) <> 0 Then
        'Doc FILEHEADER
        bdau = 1
        Get #FileNum, bdau, FH
        bdau = bdau + 8
        iItem = FH.iSize
        iRitem = FH.iSize
        ReDim ipage(iItem)
        ReDim index(iRitem)
        'Doc cac am tiet
        For i = 1 To iItem
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bPage(ub)
            Get #FileNum, bdau, bPage
            bdau = bdau + ((ub + 1) * 2)
            Set ipage(i) = New clsIPage
            Set index(i) = New clsIndex
            ipage(i).WordCont = UTF16ToUniStr(bPage)
            index(i).i = i
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
            Get #FileNum, bdau, iipage
            bdau = bdau + 4
            cWord.WordCont = UTF16ToUniStr(bWord)
            cWord.WordClass = UTF16ToUniStr(bType)
            If iipage <> p Then
                If page.AddPage(iipage, cWord) <> 0 Then GoTo Result
                p = iipage
            Else
                If page.AddWord(iipage, cWord) <> 0 Then GoTo Result
            End If
        Next i
    End If
    Close #FileNum
    Exit Function
Result:
    LoadDic = EError.LoadDicError
End Function
'Ghi du lieu tu dien ra file
Public Function SaveDic() As Integer
Dim FH As tFILEHEADER
Dim FileNum As Integer, i As Long, j As Long
Dim bPage() As Integer
Dim bWord() As Integer
Dim bType() As Integer
Dim iipage As Long
Dim bdau As Long
    SaveDic = EError.NoError
    ' Save ipage.vch
    FileNum = FreeFile
    Open fIPage For Binary As #FileNum
    bdau = 1
    If iRitem <> 0 Then
        FH.iSize = iRitem
        FH.iType = eFILETYPE.ipage
        Put #FileNum, bdau, FH
        bdau = bdau + 8
        For i = 1 To iRitem
            DoEvents
            bPage = UniStrToUTF16(Trim(ipage(index(i).i).WordCont))
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
        For i = 1 To iRitem
            For j = 1 To Me.WiPCount(i)
                DoEvents
                bWord = UniStrToUTF16(Trim(Me.WordCont(i, j)))
                bType = UniStrToUTF16(Trim(Me.WordClass(i, j)))
                iipage = i
                Put #FileNum, bdau, UBound(bWord)
                bdau = bdau + 4
                Put #FileNum, bdau, bWord
                bdau = bdau + ((UBound(bWord) + 1) * 2)
                Put #FileNum, bdau, UBound(bType)
                bdau = bdau + 4
                Put #FileNum, bdau, bType
                bdau = bdau + (UBound(bType) + 1) * 2
                Put #FileNum, bdau, iipage
                bdau = bdau + 4
            Next j
        Next i
    End If
    Close #FileNum
    Exit Function
Result:
    SaveDic = EError.SaveDicError
End Function
' Ham sap xep mang IPage()
Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
Dim t As Integer
    Sort = EError.NoError
    i = 1
    Do While (i < iRitem) And (ipage(index(i).i).WordCont < ipage(index(iRitem).i).WordCont)
        i = i + 1
    Loop
    If i < iRitem Then
        t = index(iRitem).i
        For j = iRitem To i + 1 Step -1
            index(j).i = index(j - 1).i
        Next j
        index(i).i = t
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
