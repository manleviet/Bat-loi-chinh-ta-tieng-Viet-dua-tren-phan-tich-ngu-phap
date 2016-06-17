VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsSentence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsSetWC Class - a class demonstrating a set of word classes
'----------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private aSetWC() As cpsSetWCItem
Private iItem As Long

Private Sub Class_Initialize()
    iItem = 0
    ReDim aSetWC(iItem)
End Sub

Public Property Get Count() As Long
    Count = iItem
End Property

Public Property Get WCS(ByVal i As Long) As String
    If i <= 0 Or i > iItem Then Exit Property
    WCS = aSetWC(i).WCS
End Property

Public Property Let WCS(ByVal i As Long, ByVal st As String)
    If i <= 0 Or i > iItem Then Exit Property
    aSetWC(i).WCS = st
End Property

Public Property Get RSItem(ByVal i As Long) As Long
    If i <= 0 Or i > iItem Then Exit Property
    RSItem = aSetWC(i).RSItem
End Property

Public Property Let RSItem(ByVal i As Long, ByVal j As Long)
    If i <= 0 Or i > iItem Then Exit Property
    aSetWC(i).RSItem = j
End Property

Public Property Get Right(ByVal i As Long) As Boolean
    If i <= 0 Or i > iItem Then Exit Property
    Right = aSetWC(i).Right
End Property

Public Property Let Right(ByVal i As Long, ByVal b As Boolean)
    If i <= 0 Or i > iItem Then Exit Property
    aSetWC(i).Right = b
End Property

Public Property Get RNum(ByVal i As Long) As Long
    If i <= 0 Or i > iItem Then Exit Property
    RNum = aSetWC(i).RNum
End Property

Public Property Let RNum(ByVal i As Long, ByVal j As Long)
    If i <= 0 Or i > iItem Then Exit Property
    aSetWC(i).RNum = j
End Property

Public Function AddWCS(ByVal st As String, ByVal R As Long) As Integer
On Error GoTo Result
    AddWCS = 0
    iItem = iItem + 1
    ReDim Preserve aSetWC(iItem)
    Set aSetWC(iItem) = New cpsSetWCItem
    aSetWC(iItem).WCS = st
    aSetWC(iItem).RSItem = R
    If Sort <> 0 Then GoTo Result
    Exit Function
Result:
    AddWCS = 1000
End Function

Public Function Sort() As Integer
On Error GoTo Result
Dim i As Long, j As Long
    Sort = 0
    For i = 1 To iItem - 1
        For j = i + 1 To iItem
            If aSetWC(i).WCS > aSetWC(j).WCS Then
                If Swap(i, j) <> 0 Then GoTo Result
            End If
        Next j
    Next i
    Exit Function
Result:
    Sort = 1000
End Function
' Ham hoan chuyen hai gia tri trong mang IPage
Private Function Swap(ByVal i As Integer, ByVal j As Integer) As Integer
On Error GoTo Result
Dim temp As New cpsSetWCItem
    If temp.Copy(aSetWC(i)) <> 0 Then GoTo Result
    If aSetWC(i).Copy(aSetWC(j)) <> 0 Then GoTo Result
    If aSetWC(j).Copy(temp) <> 0 Then GoTo Result
    Swap = eError.NoError
    Exit Function
Result:
    Swap = eError.SwapError
End Function

