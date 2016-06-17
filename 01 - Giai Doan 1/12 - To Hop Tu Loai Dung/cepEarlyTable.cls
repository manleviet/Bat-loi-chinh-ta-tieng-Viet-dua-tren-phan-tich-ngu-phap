VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cepEarlyTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cepEarlyTable Class - a class demonstrating a early table
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sTable() As cepEarlyCol
Private iItem As Long
Private iScan As Long

Public Sub Init()
    iItem = 1
    ReDim sTable(iItem)
    Set sTable(iItem) = New cepEarlyCol
    sTable(iItem).Init
    iScan = 1
End Sub

Public Property Get Scan() As Long
    Scan = iScan
End Property

Public Property Get WordAfterDot() As String
    WordAfterDot = sTable(iScan).WordAfterDot
End Property

Public Property Get Origin() As String
    Origin = sTable(iScan).Origin
End Property

Public Property Get RuleLeft() As String
    RuleLeft = sTable(iScan).RuleLeft
End Property

Public Function IsEmpty() As Boolean
    IsEmpty = False
    If iScan > iItem Then IsEmpty = True
End Function

Public Function IsEndCol() As Boolean
    IsEndCol = sTable(iScan).IsEndCol
End Function

Public Function IncScan() As Integer
    IncScan = sTable(iScan).IncScan
End Function

Private Function DelDot(ByVal st As String) As String
Dim p As Long
    p = InStr(1, st, ".")
    If p <> 0 Then
        DelDot = Left(st, p - 1) & Mid(st, p + 1)
    End If
    p = InStr(1, st, " ")
    DelDot = Mid(st, p + 1)
End Function

Public Function AddRule(ByVal k As Byte, luat As String) As Integer ' 1:predict;0:completer
On Error GoTo Result
Dim p As Long, i As Long
Dim ok As Boolean
    AddRule = 0
    If k = 1 Then
        p = InStr(1, luat, " ")
        'Do While p <> 0
        '    luat = Left(luat, p - 1) & Mid(luat, p + 1)
        '    p = InStr(p, luat, " ")
        'Loop
        luat = Left(luat, p - 1) & "." & Mid(luat, p + 1)
        luat = iScan & " " & luat
    End If
    ok = False
    For i = 1 To sTable(iScan).Count
        If DelDot(sTable(iScan).Cell(i)) = DelDot(luat) Then
            ok = True
            Exit For
        End If
    Next i
    If Not ok Then
        If sTable(iScan).AddRule(luat) <> 0 Then GoTo Result
    End If
    Exit Function
Result:
    AddRule = 1000
End Function

Public Function SoKhop(ByVal luat As String) As Integer
On Error GoTo Result
Dim p As Long
    SoKhop = 0
    If AddRule(1, luat) <> 0 Then GoTo Result
    p = InStr(1, luat, ".")
    If p <> 0 Then luat = Left(luat, p - 1) & " " & Mid(luat, p + 1)
    luat = luat & "."
    iItem = iItem + 1
    ReDim Preserve sTable(iItem)
    Set sTable(iItem) = New cepEarlyCol
    iScan = iItem
    If sTable(iScan).SoKhop(luat) <> 0 Then GoTo Result
    Exit Function
Result:
    SoKhop = 1000
End Function

Private Function WAD(ByVal st As String) As String
Dim p As Long, p1 As Long
    p = InStr(1, st, ".")
    WAD = ""
    If p < Len(st) Then
        p1 = InStr(p + 1, st, " ")
        If p1 = Len(st) Then WAD = Mid(st, p + 1)
        If p1 < Len(st) Then WAD = Mid(st, p + 1, p1 - p - 1)
    End If
End Function

Public Function GetRuleWithRuleLeft(ByVal Origin As String, ByVal VT As String, aRule() As String) As Integer
On Error GoTo Result
Dim i As Long
Dim tu As String
Dim r() As String, iR As Long
Dim luat As String, p As Long, p1 As Long
    GetRuleWithRuleLeft = 0
    iR = 0
    ReDim r(iR)
    For i = 1 To sTable(Origin).Count
        luat = sTable(Origin).Cell(i)
        tu = WAD(luat)
        If tu = VT Then
            iR = iR + 1
            ReDim Preserve r(iR)
            p = InStr(1, luat, ".")
            p1 = InStr(p + 1, luat, " ")
            If p1 = Len(luat) Then
                luat = Left(luat, p - 1) & " " & Mid(luat, p + 1) & "."
            ElseIf p1 < Len(luat) Then
                luat = Left(luat, p - 1) & " " & Mid(luat, p + 1, p1 - p - 1) & "." & Mid(luat, p1 + 1)
            End If
            r(iR) = luat
        End If
    Next i
    aRule = r
    Exit Function
Result:
    GetRuleWithRuleLeft = 1000
End Function

Public Function Reset(ByVal n As Long) As Integer
On Error GoTo Result
    Reset = 0
    iItem = n
    ReDim Preserve sTable(iItem)
    iScan = iItem
    If sTable(iItem).Reset <> 0 Then GoTo Result
    Exit Function
Result:
    Reset = 1000
End Function

