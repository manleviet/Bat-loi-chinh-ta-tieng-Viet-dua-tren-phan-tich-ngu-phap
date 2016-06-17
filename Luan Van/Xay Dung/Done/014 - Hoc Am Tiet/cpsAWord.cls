VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsAWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsAWord class - a class demonstrating the array words
'-------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private AW() As cpsWord
Private iItem As Long

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Let Count(ByVal Length As Long)
Dim i As Long
    iItem = Length
    ReDim AW(1 To 3, iItem)
    For i = 1 To iItem
        Set AW(1, i) = New cpsWord
        Set AW(2, i) = New cpsWord
        Set AW(3, i) = New cpsWord
    Next i
End Property

Public Property Get Word(ByVal i As Long, ByVal j As Long) As cpsWord
    Set Word = AW(i, j)
End Property

Public Function AddWord(ByVal i As Long, ByVal j As Long, ByVal vWord As cpsWord) As Integer
On Error GoTo Result
    AddWord = 0
    If i = 0 Or i > 3 Then GoTo Result
    If j = 0 Or j > iItem Then GoTo Result
    If AW(i, j).Copy(vWord) <> 0 Then GoTo Result
    Exit Function
Result:
    AddWord = 1000
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim AW(1 To 3, iItem)
End Sub
