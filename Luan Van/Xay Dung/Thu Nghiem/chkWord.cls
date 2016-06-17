VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Checking
'Description: chkWord Class - a class storing selstart and sellength of word
'----------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department


Option Explicit

Private lStart As Long
Private lLength As Long
Private lx As Long
Private ly As Long

Public Property Get Start() As Long
    Start = lStart
End Property

Public Property Let Start(ByVal i As Long)
    lStart = i
End Property

Public Property Get Length() As Long
    Length = lLength
End Property

Public Property Let Length(ByVal i As Long)
    lLength = i
End Property

Public Property Get x() As Long
    x = lx
End Property

Public Property Let x(ByVal i As Long)
    lx = i
End Property

Public Property Get y() As Long
    y = ly
End Property

Public Property Let y(ByVal i As Long)
    ly = i
End Property

Public Function Copy(ByVal vWord As chkWord) As Integer
On Error GoTo Result
    Copy = 0
    lStart = vWord.Start
    lLength = vWord.Length
    lx = vWord.x
    ly = vWord.y
    Exit Function
Result:
    Copy = 1000
End Function

Private Sub Class_Initialize()
    lStart = 0
    lLength = 0
    lx = 0
    ly = 0
End Sub
