VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWCItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Checking
'Description: clsWCItem Class - A Word Class Item Demonstration
'---------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sSign As String * 3
Private sSense As String * 100

Public Property Let Sign(ByVal st As String)
    sSign = st
End Property

Public Property Get Sign() As String
    Sign = Trim(sSign)
End Property

Public Property Let Sense(ByVal st As String)
    sSense = st
End Property

Public Property Get Sense() As String
    Sense = Trim(sSense)
End Property

Public Function Copy(ByVal WC As clsWCItem) As Integer
On Error GoTo Result
    sSign = WC.Sign
    sSense = WC.Sense
    Copy = EError.NoError
    Exit Function
Result:
    Copy = EError.CopyWordError
End Function
