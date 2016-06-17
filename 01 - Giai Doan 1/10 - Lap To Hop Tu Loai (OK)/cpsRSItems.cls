VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsRSItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private RSItem() As cpsRSItem
Private iItem As Long

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Get Word(ByVal i As Long) As cpsRSItem
    Set Word = RSItem(i)
End Property

Public Function AddWord(ByVal vWord As cpsRSItem) As Integer
On Error GoTo Result
    AddWord = 0
    iItem = iItem + 1
    ReDim Preserve RSItem(iItem)
    Set RSItem(iItem) = New cpsRSItem
    If RSItem(iItem).Copy(vWord) <> 0 Then GoTo Result
    Exit Function
Result:
    AddWord = 1000
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim RSItem(0)
End Sub
