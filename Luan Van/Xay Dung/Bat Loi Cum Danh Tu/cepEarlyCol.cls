VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cepEarlyCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cepEarlyCol Class - a class demonstrating a colummn of early table
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sValue() As cepEarlyCell
Private iItem As Long
Private iScan As Long

Public Sub Init()
    iItem = 1
    ReDim sValue(iItem)
    Set sValue(iItem) = New cepEarlyCell
    sValue(iItem).Init
    iScan = 1
End Sub

Public Property Get Cell(ByVal i As Long) As String
    Cell = sValue(i).Cell
End Property

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Get WordAfterDot() As String
    WordAfterDot = sValue(iScan).WordAfterDot
End Property

Public Property Get RuleLeft() As String
    RuleLeft = sValue(iScan).RuleLeft
End Property

Public Property Get Origin() As String
    Origin = sValue(iScan).Origin
End Property

Public Function IsEndCol() As Boolean
    IsEndCol = False
    If iScan > iItem Then IsEndCol = True
End Function

Public Function IncScan() As Integer
On Error GoTo Result
    IncScan = 0
    iScan = iScan + 1
    Exit Function
Result:
    IncScan = 1000
End Function

Public Function AddRule(ByVal luat As String) As Integer
On Error GoTo Result
    AddRule = 0
    iItem = iItem + 1
    ReDim Preserve sValue(iItem)
    Set sValue(iItem) = New cepEarlyCell
    sValue(iItem).Cell = luat
    Exit Function
Result:
    AddRule = 1000
End Function

Public Function SoKhop(ByVal luat As String) As Integer
    If iItem = 0 Then
        iItem = 1
        ReDim sValue(iItem)
        Set sValue(iItem) = New cepEarlyCell
        iScan = 1
        sValue(iScan).Cell = luat
    Else
        iItem = iItem + 1
        ReDim Preserve sValue(iItem)
        Set sValue(iItem) = New cepEarlyCell
        sValue(iItem).Cell = luat
    End If
End Function

Public Function Reset(ByVal col As Long) As Integer
On Error GoTo Result
Dim dem As Long
Dim isBegin As Long
Dim i As LoadPictureColorConstants
    Reset = 0
    dem = 0
    For i = 1 To iItem
        If sValue(i).Origin = col - 1 And Right(sValue(i).Cell, 1) = "." And _
            IsTerminal(Left(Right(sValue(i).Cell, 4), 3)) Then
            dem = dem + 1
        Else
            Exit For
        End If
    Next i
    iItem = dem
    ReDim Preserve sValue(iItem)
    iScan = 1
    Exit Function
Result:
    Reset = 1000
End Function

Private Function IsTerminal(ByVal tu As String) As Boolean
Dim p As Long
    IsTerminal = False
    p = InStr(1, Ter, Left(tu, 1))
    If p <> 0 Then IsTerminal = True
End Function
