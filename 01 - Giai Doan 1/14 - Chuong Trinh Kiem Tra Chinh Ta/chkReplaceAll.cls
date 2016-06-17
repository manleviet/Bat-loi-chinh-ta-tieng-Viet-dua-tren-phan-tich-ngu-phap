VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkReplaceAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkReplaceAll Class - a class demonstrating a array of Replace All
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sValue() As String
Private sReplace() As String
Private iItem As Long

Public Function Clear() As Integer
On Error GoTo Result
    Clear = 0
    iItem = 0
    ReDim sValue(iItem)
    ReDim sReplace(iItem)
    Exit Function
Result:
    Clear = 1000
End Function

Public Function AddReplace(ByVal stValue As String, ByVal stRepl As String) As Integer
On Error GoTo Result
    AddReplace = 0
    iItem = iItem + 1
    ReDim Preserve sValue(iItem)
    ReDim Preserve sReplace(iItem)
    sValue(iItem) = stValue
    sReplace(iItem) = stRepl
    Exit Function
Result:
    AddReplace = 1000
End Function

Public Function CheckReplace(st As String) As Boolean
Dim i As Long
    CheckReplace = False
    For i = 1 To iItem
        If st = sValue(i) Then
            st = sReplace(i)
            CheckReplace = True
            Exit Function
        End If
    Next i
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim sValue(iItem)
    ReDim sReplace(iItem)
End Sub
