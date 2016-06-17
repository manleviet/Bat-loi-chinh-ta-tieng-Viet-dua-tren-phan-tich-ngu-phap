VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkAWRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkAWRow Class - a class demonstrating a row of AWord
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private aRow() As chkWord
Private bIgnore As Boolean

Public Property Get Ignore() As Boolean
    Ignore = bIgnore
End Property

Public Property Let Ignore(ByVal b As Boolean)
    bIgnore = b
End Property

Public Property Get Cell(ByVal i As Long) As chkWord
    If i >= 1 And i <= 3 Then
        Set Cell = aRow(i)
    Else
        Set Cell = Nothing
    End If
End Property

Public Property Let Cell(ByVal i As Long, ByVal vWord As chkWord)
    If i >= 1 And i <= 3 Then
        Call aRow(i).Copy(vWord)
    End If
End Property

Public Function Copy(ByVal vWord As chkAWRow) As Integer
On Error GoTo Result
    Copy = 0
    If aRow(1).Copy(vWord.Cell(1)) <> 0 Then GoTo Result
    If aRow(2).Copy(vWord.Cell(2)) <> 0 Then GoTo Result
    If aRow(3).Copy(vWord.Cell(3)) <> 0 Then GoTo Result
    Exit Function
Result:
    Copy = 1000
End Function

Private Sub Class_Initialize()
    ReDim aRow(1 To 3)
    Set aRow(1) = New chkWord
    Set aRow(2) = New chkWord
    Set aRow(3) = New chkWord
    bIgnore = False
End Sub
