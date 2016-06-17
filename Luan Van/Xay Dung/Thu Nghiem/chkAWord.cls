VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkAWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkAWord class - a class demonstrating the array words
'-------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private AW() As chkAWRow
Private iItem As Long

Public Property Get Ignore(ByVal i As Long) As Boolean
    Ignore = AW(i).Ignore
End Property

Public Property Let Ignore(ByVal i As Long, ByVal b As Boolean)
    AW(i).Ignore = b
End Property

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Let Count(ByVal Length As Long)
Dim i As Long
    iItem = Length
    ReDim AW(iItem)
    For i = 1 To iItem
        Set AW(i) = New chkAWRow
        Set AW(i) = New chkAWRow
        Set AW(i) = New chkAWRow
    Next i
End Property

Public Property Get Word(ByVal i As Long, ByVal j As Long) As chkWord
    Set Word = AW(j).Cell(i)
End Property

Public Function AddWord(ByVal i As Long, ByVal j As Long, ByVal vWord As chkWord) As Integer
On Error GoTo Result
    AddWord = 0
    If i = 0 Or i > 3 Then GoTo Result
    If j = 0 Or j > iItem Then GoTo Result
    If AW(j).Cell(i).Copy(vWord) <> 0 Then GoTo Result
    Exit Function
Result:
    AddWord = 1000
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim AW(iItem)
End Sub
