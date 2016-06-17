VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsSyllable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsSyllable Class - a class demonstrating syllables of sentence
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private Syll() As cpsWord
Private iItem As Long

Public Property Get SyllCount() As Long
    SyllCount = iItem
End Property

Public Property Get Syllable(ByVal i As Long) As cpsWord
    Set Syllable = Syll(i)
End Property

Public Function AddSyll(ByVal vWord As cpsWord) As Integer
On Error GoTo Result
    AddSyll = 0
    iItem = iItem + 1
    ReDim Preserve Syll(iItem)
    Set Syll(iItem) = New cpsWord
    If Syll(iItem).Copy(vWord) <> 0 Then GoTo Result
    Exit Function
Result:
    AddSyll = 1000
End Function

Public Function Clear() As Integer
On Error GoTo Result
    Clear = 0
    iItem = 0
    ReDim Syll(iItem)
    Exit Function
Result:
    Clear = 1000
End Function
