VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Checking
'Desciption: clsWordClass Class - A Word Class Array Demonstration
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private WCItem() As clsWCItem
Private iItem As Long

Public Function FindWC(ByVal Sign As String) As clsLocation
On Error GoTo Result
Dim Low As Long
Dim High As Long
Dim Mid As Long
Dim Loca As New clsLocation
    Low = 1
    High = iItem
    Do While High >= Low ' tim kiem nhi phan
        Mid = (High + Low) \ 2
        Select Case SoSanh(Trim(WCItem(Mid).Sign), Sign)
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
    Set FindWC = Loca
    Exit Function
Result:
    Set FindWC = Nothing
End Function

Public Function DelWC(ByVal Sign As String) As Integer
On Error GoTo Result
Dim i As Long
Dim Loca As New clsLocation
    DelWC = eError.NoError
    Set Loca = FindWC(Sign) ' lay vi tri tu can xoa
    If Loca Is Nothing Then GoTo Result
    If Loca.ok <> 0 Then ' Khong co tu can xoa
        DelWC = eError.NoHaveWord
    Else ' Co
        i = Loca.x
        Do While i < iItem 'Don cac tu khac len
            If Swap(i, i + 1) <> 0 Then GoTo Result
            i = i + 1
        Loop
        iItem = iItem - 1 ' giam so luong phan tu di mot
        ReDim Preserve WCItem(iItem) ' thiet dat lai kich thuoc mang
    End If
    Exit Function
Result:
    DelWC = eError.DelError
End Function

Public Function AddWC(ByVal WC As clsWCItem) As Integer
On Error GoTo Result
Dim fWord As String
Dim l As Long
Dim Loca As New clsLocation
    AddWC = eError.NoError
    Set Loca = FindWC(WC.Sign)
    If Loca Is Nothing Then GoTo Result
    If Loca.ok <> 0 Then
        iItem = iItem + 1 ' tang so luong phan tu len mot
        ReDim Preserve WCItem(iItem) ' thiet dat lai kich thuoc mang
        Set WCItem(iItem) = New clsWCItem
        If WCItem(iItem).Copy(WC) <> 0 Then GoTo Result ' dat tu vao cuoi mang
        If Sort <> 0 Then GoTo Result ' sap xep lai mang
    End If
    Exit Function
Result:
    AddWC = eError.AddWordError
End Function

Public Property Let Sense(ByVal i As Long, ByVal vData As String)
    WCItem(i).Sense = vData
End Property

Public Property Get Sense(ByVal i As Long) As String
    Sense = WCItem(i).Sense
End Property

Public Property Let Sign(ByVal i As Long, ByVal vData As String)
    WCItem(i).Sign = vData
End Property

Public Property Get Sign(ByVal i As Long) As String
    Sign = WCItem(i).Sign
End Property

Public Property Get Count() As Long
     Count = iItem
End Property

'Public Sub Item(ByVal i As Long, ByVal vData As clsWCItem)
'    WCItem(i).Sense = vData.Sense
'    WCItem(i).Sign = vData.Sign
'End Sub

'Public Sub Item(ByVal i As Long, cItem As clsWCItem)
'    cItem.Sense = WCItem(i).Sense
'    cItem.Sign = WCItem(i).Sign
'End Sub

Private Sub Class_Initialize()
    iItem = 0
    ReDim WCItem(iItem)
End Sub

Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
    Sort = eError.NoError
    For i = 1 To iItem - 1
        For j = i + 1 To iItem
            If WCItem(i).Sign > WCItem(j).Sign Then   ' kiem tra lon hon
                If Swap(i, j) <> 0 Then GoTo Result ' Hoan chuyen hai phan tu
            End If
        Next j
    Next i
    Exit Function
Result:
    Sort = eError.SortError
End Function
' Ham hoan chuyen hai gia tri trong mang IPage
Private Function Swap(ByVal i As Integer, ByVal j As Integer) As Integer
On Error GoTo Result
Dim temp As New clsWCItem
    Swap = eError.NoError
    If temp.Copy(WCItem(i)) <> 0 Then GoTo Result
    If WCItem(i).Copy(WCItem(j)) <> 0 Then GoTo Result
    If WCItem(j).Copy(temp) <> 0 Then GoTo Result
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
'Lay du lieu tu file tu loai
Public Function LoadDic() As Integer
On Error GoTo Result
Dim FH As tFILEHEADER, ub As Long
Dim FileNum As Integer, i As Long
Dim WC As New clsWCItem
Dim bSign() As Integer
Dim bSense() As Integer
Dim bdau As Long
    LoadDic = eError.NoError
    'load file wclass.vch
    FileNum = FreeFile
    Open fWClass For Binary As #FileNum
    If FileLen(fWClass) <> 0 Then
        'Doc FILEHEADER
        bdau = 1
        Get #FileNum, bdau, FH
        bdau = bdau + 8
        'Doc cac am tiet
        For i = 1 To FH.iSize
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bSign(ub)
            Get #FileNum, bdau, bSign
            bdau = bdau + ((ub + 1) * 2)
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bSense(ub)
            Get #FileNum, bdau, bSense
            bdau = bdau + ((ub + 1) * 2)
            WC.Sign = UTF16ToUniStr(bSign)
            WC.Sense = UTF16ToUniStr(bSense)
            If Me.AddWC(WC) <> 0 Then GoTo Result
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
Dim FileNum As Integer, i As Long
Dim bSign() As Integer
Dim bSense() As Integer
Dim bdau As Long
    SaveDic = eError.NoError
    'Save wclass.vch
    FileNum = FreeFile
    Open fWClass For Binary As #FileNum
    bdau = 1
    If Me.Count <> 0 Then
        FH.iSize = Me.Count
        FH.iType = 2
        Put #FileNum, bdau, FH
        bdau = bdau + 8
        For i = 1 To iItem
            bSign = UniStrToUTF16(Trim(Me.Sign(i)))
            bSense = UniStrToUTF16(Trim(Me.Sense(i)))
            Put #FileNum, bdau, UBound(bSign)
            bdau = bdau + 4
            Put #FileNum, bdau, bSign
            bdau = bdau + ((UBound(bSign) + 1) * 2)
            Put #FileNum, bdau, UBound(bSense)
            bdau = bdau + 4
            Put #FileNum, bdau, bSense
            bdau = bdau + (UBound(bSense) + 1) * 2
        Next i
    End If
    Close #FileNum
    Exit Function
Result:
    SaveDic = eError.SaveDicError
End Function
