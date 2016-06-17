VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "crlRuleRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: crlRuleRight Class - a class demonstrating rule right
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private vRRight() As crlRuleRightPage
Private iItem As Long

Public Property Get RCount() As Long
Dim i As Long
    i = 1
    RCount = 0
    Do While i <= iItem
        RCount = RCount + vRRight(i).Count
        i = i + 1
    Loop
End Property

Public Property Get RiPCount(ByVal i As Long) As Long
    RiPCount = vRRight(i).Count
End Property

Public Property Get RuleRight(ByVal i As String, ByVal j As String) As String
    RuleRight = vRRight(i).RuleRight(j)
End Property

Public Property Let RuleRight(ByVal i As String, ByVal j As String, ByVal st As String)
    vRRight(i).RuleRight(j) = st
End Property

Public Function AddRuleRight(ByVal i As Long, ByVal vRule As crlRuleItem) As Integer
    AddRuleRight = vRRight(i).AddRuleRight(vRule)
End Function

Public Function FindRuleRight(ByVal i As Long, ByVal st As String) As clsLocation
    Set FindRuleRight = vRRight(i).FindRuleRight(i, st)
End Function

Public Function DelRuleRightPage(ByVal i As Long) As Integer
On Error GoTo Result
    DelRuleRightPage = 0
    Do While i < iItem
        If vRRight(i).Copy(vRRight(i + 1)) <> 0 Then GoTo Result
        i = i + 1
    Loop
    iItem = iItem - 1
    ReDim Preserve vRRight(iItem)
    Exit Function
Result:
    DelRuleRightPage = 1000
End Function

Public Function DelRuleRight(ByVal i As Long, ByVal st As String) As Integer
    DelRuleRight = vRRight(i).DelRuleRight(st)
End Function

Public Function AddPage(ByVal i As Long, ByVal vRule As crlRuleItem) As Integer
On Error GoTo Result
Dim j As Long
    AddPage = 0
    iItem = iItem + 1
    ReDim Preserve vRRight(iItem)
    Set vRRight(iItem) = New crlRuleRightPage
    j = iItem
    Do While j > i
        If vRRight(j).Copy(vRRight(j - 1)) <> 0 Then GoTo Result
        j = j - 1
    Loop
    If vRRight(i).Clear <> 0 Then GoTo Result
    If vRRight(i).AddRuleRight(vRule) <> 0 Then GoTo Result
    Exit Function
Result:
    AddPage = 1000
End Function

Private Sub Class_Initialize()
    iItem = 0
    ReDim vRRight(iItem)
End Sub
