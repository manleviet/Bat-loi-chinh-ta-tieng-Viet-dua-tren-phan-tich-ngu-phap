VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkIgnore Class - a class demonstrating a array of Ignore Value
'--------------------------------
'Author: Le Viet Man
'   Unviersity of Hue
'   College of Sciences - IT Department

Option Explicit
Private sValue() As String
Private lIndex() As Long
Private iItem As Long

Private Sub Class_Initialize()
    iItem = 0
    ReDim sValue(iItem)
    ReDim lIndex(iItem)
End Sub

Public Function Clear() As Integer
On Error GoTo Result
    Clear = 0
    iItem = 0
    ReDim sValue(iItem)
    ReDim lIndex(iItem)
    Exit Function
Result:
    Clear = 1000
End Function

Public Function AddIgnore(ByVal st As String, ByVal i As Long) As Integer
On Error GoTo Result
    AddIgnore = 0
    iItem = iItem + 1
    ReDim Preserve sValue(iItem)
    ReDim Preserve lIndex(iItem)
    sValue(iItem) = st
    lIndex(iItem) = i
    Exit Function
Result:
    AddIgnore = 1000
End Function

Public Function CheckIgnore(ByVal st As String, ByVal i As Long) As Boolean
Dim j As Long
    CheckIgnore = False
    For j = 1 To iItem
        If sValue(j) = st And lIndex(j) = i Then
            CheckIgnore = True
            Exit Function
        End If
    Next j
End Function

