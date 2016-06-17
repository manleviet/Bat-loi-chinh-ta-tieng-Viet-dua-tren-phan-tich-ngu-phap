VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "crlRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: crlRule Class - a class demonstrating rule dictionary
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private RRight As crlRuleRight
Private RLeft() As crlRuleItem
Private index() As clsIndex
Private iItem As Long
Private iRItem As Long
'so luong luat sinh
Public Property Get RCount() As Long
    RCount = RRight.RCount
End Property
'so luong luat sinh trong trang thu i
Public Property Get RiPCount(ByVal i As Long) As Long
    RiPCount = RRight.RiPCount(index(i).i)
End Property
'so trang luat
Public Property Get PCount() As Long
    PCount = iRItem
End Property
'lay luat thu j trong trang i
Public Property Get Rule(ByVal i As Long, ByVal j As Long) As String
    'Rule = RLeft(index(i).i).Content & " " & RRight.RuleRight(index(i).i, j)
    Rule = RLeft(index(i).i).Content & " " & RRight.RuleRight(i, j)
    'can xem lai
    '
    '
    '
    
End Property

Public Property Let Rule(ByVal i As Long, ByVal j As Long, ByVal st As String)
Dim VT As String, VP As String
Dim p As Integer
    st = Trim(st)
    p = InStr(1, st, " ")
    VT = Left(st, p - 1)
    VP = Mid(st, p + 1)
    RLeft(index(i).i).Content = VT
    RRight.RuleRight(i, j) = VP
End Property

Private Sub Class_Initialize()
    iItem = 0
    iRItem = 0
    ReDim RLeft(iItem)
    ReDim index(iRItem)
    Set RRight = New crlRuleRight
End Sub

Public Function AddRule(ByVal Content As String) As Integer
On Error GoTo Result
Dim fWord As String
Dim loca As New clsLocation
Dim vRule As New crlRuleItem
Dim p As Integer
    AddRule = eError.NoError
    fWord = FirstWord(Content)  ' cat tu dau trong content
    Set loca = FindRuleLeft(fWord)
    If loca Is Nothing Then GoTo Result
    p = InStr(1, Content, " ")
    vRule.Content = Mid(Content, p + 1)
    If loca.ok = 0 Then
        If RRight.AddRuleRight(index(loca.x).i, vRule) <> 0 Then GoTo Result
    Else
        iItem = iItem + 1 ' tang so luong trang len mot
        iRItem = iRItem + 1
        ReDim Preserve RLeft(iItem) ' thiet dat lai kich thuoc mang
        ReDim Preserve index(iRItem)
        Set RLeft(iItem) = New crlRuleItem
        Set index(iRItem) = New clsIndex
        RLeft(iItem).Content = fWord  ' dat tu vao cuoi mang
        index(iRItem).i = iItem
        If Sort <> 0 Then GoTo Result ' sap xep lai mang
        Set loca = FindRuleLeft(fWord) ' lay vi tri hien tai cua tu can tim
        'them mot trang trong page
        If loca.ok = 0 Then
            If RRight.AddPage(index(loca.x).i, vRule) <> 0 Then GoTo Result
        Else
            GoTo Result
        End If
    End If
    Exit Function
Result:
    AddRule = eError.AddWordError
End Function

Public Function FindRuleLeft(ByVal Content As String) As clsLocation
On Error GoTo Result
Dim Low As Long
Dim High As Long
Dim Mid As Long
Dim loca As New clsLocation
    Low = 1
    High = iRItem
    Do While High >= Low ' tim kiem nhi phan
        Mid = (High + Low) \ 2
        Select Case SoSanh(RLeft(index(Mid).i).Content, Content)
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
    Set FindRuleLeft = loca
    Exit Function
Result:
    FindRuleLeft = Nothing
End Function

Public Function FindRule(ByVal Content As String) As clsLocation
On Error GoTo Result
Dim fWord As String
Dim loca As New clsLocation
Dim p As Integer
    fWord = FirstWord(Content)
    Set loca = FindRuleLeft(fWord)
    If Not (loca Is Nothing) And loca.ok = 0 Then
        p = InStr(1, Content, " ")
        fWord = Mid(Content, p + 1)
        Set loca = RRight.FindRuleRight(index(loca.x).i, fWord)
    End If
    Set FindRule = loca
    Exit Function
Result:
    FindRule = Nothing
End Function

Public Function DelRule(ByVal Content As String) As Integer
On Error GoTo Result
Dim fWord As String, i As Long
Dim loca As New clsLocation
Dim p As Integer, lWord As String
    DelRule = eError.NoError
    fWord = FirstWord(Content)
    p = InStr(1, Content, " ")
    lWord = Mid(Content, p + 1)
    Set loca = FindRuleLeft(fWord)
    If Not (loca Is Nothing) And loca.ok = 0 Then
        If RRight.RiPCount(index(loca.x).i) = 1 Then
            i = loca.x
            Do While i < iItem
                index(i).i = index(i + 1).i
                i = i + 1
            Loop
            iRItem = iRItem - 1
            ReDim Preserve index(iRItem)
            'If RRight.DelRuleRightPage(index(loca.x).i) <> 0 Then GoTo Result
        Else
            If RRight.DelRuleRight(index(loca.x).i, lWord) <> 0 Then GoTo Result
        End If
    Else
        DelRule = eError.NoHaveWord
    End If
    Exit Function
Result:
    DelRule = eError.DelError
End Function
'Lay du lieu tu dien tu file
Public Function LoadDic() As Integer
'On Error GoTo Result
Dim FH As tFILEHEADER, ub As Long
Dim FileNum As Integer, i As Integer, j As Long, p As Long
Dim vRule As New crlRuleItem
Dim bIRule() As Integer
Dim bRule() As Integer
Dim iIRule As Long
Dim bdau As Long
    LoadDic = eError.NoError
    'load file irule.vch
    FileNum = FreeFile
    Open fIRule For Binary As #FileNum
    If FileLen(fIRule) <> 0 Then
        'Doc FILEHEADER
        bdau = 1
        Get #FileNum, bdau, FH
        bdau = bdau + 8
        iItem = FH.iSize
        iRItem = iItem
        ReDim RLeft(iItem)
        ReDim index(iRItem)
        For i = 1 To iItem
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bIRule(ub)
            Get #FileNum, bdau, bIRule
            bdau = bdau + ((ub + 1) * 2)
            Set RLeft(i) = New crlRuleItem
            Set index(i) = New clsIndex
            RLeft(i).Content = UTF16ToUniStr(bIRule)
            index(i).i = i
        Next i
    End If
    Close #FileNum
    'load file rule.vch
    FileNum = FreeFile
    Open fRule For Binary As #FileNum
    bdau = 1
    If FileLen(fRule) <> 0 Then
        Get #FileNum, bdau, FH
        bdau = bdau + 8
        j = FH.iSize
        p = 0
        For i = 1 To j
            Get #FileNum, bdau, ub
            bdau = bdau + 4
            ReDim bRule(ub)
            Get #FileNum, bdau, bRule
            bdau = bdau + ((ub + 1) * 2)
            Get #FileNum, bdau, iIRule
            bdau = bdau + 4
            vRule.Content = UTF16ToUniStr(bRule)
            If iIRule <> p Then
                If RRight.AddPage(iIRule, vRule) <> 0 Then GoTo Result
                p = iIRule
            Else
                If RRight.AddRuleRight(iIRule, vRule) <> 0 Then GoTo Result
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
Dim bIRule() As Integer
Dim bRule() As Integer
Dim iIRule As Long
Dim bdau As Long
    SaveDic = eError.NoError
    ' Save irule.vch
    FileNum = FreeFile
    Open fIRule For Binary As #FileNum
    bdau = 1
    If iRItem <> 0 Then
        FH.iSize = iRItem
        FH.iType = 3
        Put #FileNum, bdau, FH
        bdau = bdau + 8
        For i = 1 To iRItem
            bIRule = UniStrToUTF16(Trim(RLeft(index(i).i).Content))
            Put #FileNum, bdau, UBound(bIRule)
            bdau = bdau + 4
            Put #FileNum, bdau, bIRule
            bdau = bdau + ((UBound(bIRule) + 1) * 2)
        Next i
    End If
    Close #FileNum
    'Save rule.vch
    FileNum = FreeFile
    Open fRule For Binary As #FileNum
    bdau = 1
    If Me.RCount <> 0 Then
        FH.iSize = Me.RCount
        FH.iType = 3
        Put #FileNum, bdau, FH
        bdau = bdau + 8
        For i = 1 To iRItem
            For j = 1 To RRight.RiPCount(index(i).i)
                bRule = UniStrToUTF16(Trim(RRight.RuleRight(index(i).i, j)))
                iIRule = i
                Put #FileNum, bdau, UBound(bRule)
                bdau = bdau + 4
                Put #FileNum, bdau, bRule
                bdau = bdau + ((UBound(bRule) + 1) * 2)
                Put #FileNum, bdau, iIRule
                bdau = bdau + 4
            Next j
        Next i
    End If
    Close #FileNum
    Exit Function
Result:
    SaveDic = eError.SaveDicError
End Function

Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
Dim t As Long
    Sort = eError.NoError
    i = 1
    Do While (i < iRItem) And (RLeft(index(i).i).Content < RLeft(index(iRItem).i).Content)
        i = i + 1
    Loop
    If i < iRItem Then
        t = index(iRItem).i
        For j = iRItem To i + 1 Step -1
            index(j).i = index(j - 1).i
        Next j
        index(i).i = t
    End If
    Exit Function
Result:
    Sort = eError.SortError
End Function

'Private Function Swap(i As Integer, j As Integer) As Integer
'On Error GoTo Result
'Dim temp As New crlRuleItem
'    If temp.Copy(RLeft(i)) <> 0 Then GoTo Result
'    If RLeft(i).Copy(RLeft(j)) <> 0 Then GoTo Result
'    If RLeft(j).Copy(temp) <> 0 Then GoTo Result
'    Swap = eError.NoError
'    Exit Function
'Result:
'    Swap = eError.SwapError
'End Function
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

