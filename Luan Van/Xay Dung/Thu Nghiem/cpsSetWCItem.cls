VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsSetWCItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsSetWCItem Class - a class demonstrating a item's set word classes
'-----------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sWCS As String
Private iRSItem As Long
Private bRight As Boolean
Private iRNum As Long

Public Property Get WCS() As String
    WCS = Trim(sWCS)
End Property

Public Property Let WCS(ByVal st As String)
    sWCS = st
End Property

Public Property Get RSItem() As Long
    RSItem = iRSItem
End Property

Public Property Let RSItem(ByVal i As Long)
    iRSItem = i
End Property

Public Property Get Right() As Boolean
    Right = bRight
End Property

Public Property Let Right(ByVal a As Boolean)
    bRight = a
End Property

Public Property Get RNum() As Long
    RNum = iRNum
End Property

Public Property Let RNum(ByVal i As Long)
    iRNum = i
End Property

Private Sub Class_Initialize()
    sWCS = ""
    bRight = False
    iRSItem = 0
    iRNum = 0
End Sub

Public Function Copy(ByVal s As cpsSetWCItem) As Integer
On Error GoTo Result
    Copy = 0
    sWCS = s.WCS
    iRSItem = s.RSItem
    bRight = s.Right
    iRNum = s.RNum
    Exit Function
Result:
    Copy = 1000
End Function
