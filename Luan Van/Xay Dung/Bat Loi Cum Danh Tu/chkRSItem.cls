VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkRSItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkRSItem Class - a class demonstrating a Right Sentence Item
'---------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private ix As Long
Private iy As Long

Public Property Get x() As Long
    x = ix
End Property

Public Property Let x(ByVal i As Long)
    ix = i
End Property

Public Property Get y() As Long
    y = iy
End Property

Public Property Let y(ByVal i As Long)
    iy = i
End Property

Public Function Copy(ByVal cx As chkRSItem) As Integer
On Error GoTo Result
    Copy = 0
    ix = cx.x
    iy = cx.y
    Exit Function
Result:
    Copy = 1000
End Function

Private Sub Class_Initialize()
    ix = 0
    iy = 0
End Sub

