VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "crlRuleRightPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: crlRuleRightPage Class - a class demonstrating rule right page
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private RRight() As crlRuleItem
Private iItem As Long

Private Sub Class_Initialize()
    iItem = 0
    ReDim RRight(iItem)
End Sub

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Get RuleRight(ByVal i As String) As String
    RuleRight = RRight(i).Content
End Property

Public Property Let RuleRight(ByVal i As String, ByVal st As String)
    RRight(i).Content = st
End Property

Public Function AddRuleRight(ByVal vRule As crlRuleItem) As Integer
On Error GoTo Result
Dim loca As New clsLocation
    AddRuleRight = 0
    Set loca = FindRuleRight(0, vRule.Content)
    If Not (loca Is Nothing) And loca.ok <> 0 Then
        iItem = iItem + 1
        ReDim Preserve RRight(iItem)
        Set RRight(iItem) = New crlRuleItem
        If RRight(iItem).Copy(vRule) <> 0 Then GoTo Result
        If Sort <> 0 Then GoTo Result
    End If
    Exit Function
Result:
    AddRuleRight = 1000
End Function

Public Function FindRuleRight(ByVal i As Long, ByVal st As String) As clsLocation
On Error GoTo Result
Dim Low As Long, High As Long, Mid As Long
Dim loca As New clsLocation
    Low = 1
    High = iItem
    Do While High >= Low
        Mid = (High + Low) \ 2
        Select Case SoSanh(RRight(Mid).Content, st)
            Case 1: High = Mid - 1
            Case -1: Low = Mid + 1
            Case 0: Exit Do
        End Select
    Loop
    loca.x = i
    loca.y = Mid
    If High >= Low Then
        loca.ok = 0
    ElseIf Mid = High Then
        loca.ok = 1
    ElseIf Mid = Low Then
        loca.ok = -1
    End If
    Set FindRuleRight = loca
    Exit Function
Result:
    Set FindRuleRight = Nothing
End Function

Public Function DelRuleRight(ByVal st As String) As Integer
On Error GoTo Result
Dim i As Long
Dim loca As New clsLocation
    DelRuleRight = 0
    Set loca = FindRuleRight(0, st)
    If Not (loca Is Nothing) And loca.ok = 0 Then
        i = loca.y
        Do While i < iItem
            If RRight(i).Copy(RRight(i + 1)) <> 0 Then GoTo Result
            i = i + 1
        Loop
        iItem = iItem - 1
        ReDim Preserve RRight(iItem)
    Else
        DelRuleRight = 1000
    End If
    Exit Function
Result:
    DelRuleRight = 1000
End Function
' Ham sap xep mang
Private Function Sort() As Integer
On Error GoTo Result
Dim i As Integer
Dim j As Integer
    Sort = eError.NoError
    For i = 1 To iItem - 1
        For j = i + 1 To iItem
            If RRight(i).Content > RRight(j).Content Then   ' kiem tra lon hon
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
Dim temp As New crlRuleItem
    If temp.Copy(RRight(i)) <> 0 Then GoTo Result
    If RRight(i).Copy(RRight(j)) <> 0 Then GoTo Result
    If RRight(j).Copy(temp) <> 0 Then GoTo Result
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

Public Function Copy(ByVal vPage As crlRuleRightPage) As Integer
On Error GoTo Result
Dim i As Long
    Copy = 0
    iItem = vPage.Count
    ReDim RRight(iItem)
    i = 1
    Do While i <= iItem
        Set RRight(i) = New crlRuleItem
        RRight(i).Content = vPage.RuleRight(i)
        i = i + 1
    Loop
    Exit Function
Result:
    Copy = 1000
End Function

Public Function Clear() As Integer
On Error GoTo Result
    Clear = 0
    iItem = 0
    ReDim RRight(iItem)
    Exit Function
Result:
    Clear = 1000
End Function
