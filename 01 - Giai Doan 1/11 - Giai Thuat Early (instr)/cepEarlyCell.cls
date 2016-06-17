VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cepEarlyCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cepEarlyCell Class - a class demonstrating a cell of early table
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sValue As String

Public Sub Init()
    sValue = "1 ROT.S01"
End Sub

Public Property Get WordAfterDot() As String
Dim p As Long, p1 As Long
    p = InStr(1, sValue, ".")
    If p = Len(sValue) Then
        WordAfterDot = ""
    Else
        p1 = InStr(p + 1, sValue, " ")
        If p1 <> Len(sValue) Then WordAfterDot = Mid(sValue, p + 1, p1 - p - 1)
    End If
End Property

Public Property Let Cell(ByVal st As String)
    sValue = st
End Property

Public Property Get Cell() As String
    Cell = sValue
End Property

Public Property Get RuleLeft() As String
Dim p As Long, p1 As Long, p2 As Long
    p = InStr(1, sValue, " ")
    p1 = InStr(p + 1, sValue, " ")
    p2 = InStr(p + 1, sValue, ".")
    If p1 < p2 Then
        RuleLeft = Mid(sValue, p + 1, p1 - p - 1)
    Else
        RuleLeft = Mid(sValue, p + 1, p2 - p - 1)
    End If
End Property

Public Property Get Origin() As String
Dim p As Long
    p = InStr(1, sValue, " ")
    Origin = Left(sValue, p - 1)
End Property
