VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkRSItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkRSItems Class - a class demonstrating Items of Right Sentence
'---------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private RSItem() As chkRSItem
Private iItem As Long

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Get Word(ByVal i As Long) As chkRSItem
    Set Word = RSItem(i)
End Property

Public Function AddWord(ByVal vWord As chkRSItem) As Integer
On Error GoTo Result
    AddWord = 0
    iItem = iItem + 1
    ReDim Preserve RSItem(iItem)
    Set RSItem(iItem) = New chkRSItem
    If RSItem(iItem).Copy(vWord) <> 0 Then GoTo Result
    Exit Function
Result:
    AddWord = 1000
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim RSItem(0)
End Sub

