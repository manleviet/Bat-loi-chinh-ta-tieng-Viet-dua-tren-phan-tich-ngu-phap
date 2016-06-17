VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: clsLocation Class - Word's Location In Dictionary Demonstration
'---------------------------------
'Author: Le Viet Man
'   Universitys of Hue
'   College of Sciences - IT Department

Option Explicit
Public x As Long ' trang trong dictionary
Public y As Long ' tu trong dictionary
Public ok As Integer ' nhan 3 gia tri 1, 0, -1 (tuong ung voi >, =, <)
