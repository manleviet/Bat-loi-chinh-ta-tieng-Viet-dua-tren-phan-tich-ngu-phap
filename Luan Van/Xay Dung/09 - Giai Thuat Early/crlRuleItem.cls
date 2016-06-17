VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "crlRuleItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: crlRuleItem Class - a class demonstrating the rule item
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sContent As String

Public Property Get Content() As String
    Content = Trim(sContent)
End Property

Public Property Let Content(ByVal st As String)
    sContent = Trim(st)
End Property

Public Function Copy(ByVal vRule As crlRuleItem) As Integer
On Error GoTo Result
    Copy = 0
    sContent = vRule.Content
    Exit Function
Result:
    Copy = 1000
End Function
