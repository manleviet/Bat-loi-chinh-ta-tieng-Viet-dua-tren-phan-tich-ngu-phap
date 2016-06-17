VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsRS Class - a class demonstrating a Right Sentence
'--------------------------------
'Author: Le Viet Man
'   Unversity of Hue
'   College of Sciences - IT Department

Private cRS() As cpsRSItems
Private iItem As Long

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Get SCount(ByVal i As Long) As Long
    SCount = cRS(i).Count
End Property

Public Property Get Word(ByVal i As Long, ByVal j As Long) As cpsRSItem
    If i >= 0 And i <= iItem Then
        If j >= 0 And j <= cRS(i).Count Then
            Set Word = cRS(i).Word(j)
        End If
    End If
End Property

Public Function AddSentence(ByVal st As String) As Integer
On Error GoTo Result
Dim p As Long, p1 As Long
Dim c As New cpsRSItem
    AddSentence = 0
    iItem = iItem + 1
    ReDim Preserve cRS(iItem)
    Set cRS(iItem) = New cpsRSItems
    p = InStr(1, st, "|")
    p1 = InStr(p + 1, st, "|")
    Do While p <> 0
        c.x = CLng(Left(st, p - 1))
        c.y = CLng(Mid(st, p + 1, p1 - p - 1))
        If cRS(iItem).AddWord(c) <> 0 Then GoTo Result
        st = Mid(st, p1 + 1)
        p = InStr(1, st, "|")
        p1 = InStr(p + 1, st, "|")
    Loop
    Exit Function
Result:
    AddSentence = 1000
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim cRS(0)
End Sub
