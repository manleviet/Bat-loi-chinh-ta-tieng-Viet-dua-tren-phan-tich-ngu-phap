VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Project: Vietnamese Checking
'Description: clsWord Class - Word Demonstration
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private sText As String
Private sType As String

Public Property Let WordClass(ByVal Class As String)
    sType = Trim(Class)
End Property

Public Property Get WordClass() As String
    WordClass = Trim(sType)
End Property

Public Property Let WordCont(ByVal Content As String)
    sText = Trim(Content)
End Property

Public Property Get WordCont() As String
    WordCont = Trim(sText)
End Property

Public Function Copy(ByVal vWord As clsWord) As Integer
On Error GoTo Result
    sText = vWord.WordCont
    sType = vWord.WordClass
    Copy = EError.NoError
    Exit Function
Result:
    Copy = EError.CopyWordError
End Function
