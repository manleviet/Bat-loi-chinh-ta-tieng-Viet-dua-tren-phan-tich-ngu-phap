Attribute VB_Name = "mdlString"
'Project: Vietnamese Checking
'Description: mdlString Modul - String type's Functions Declaration
'----------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

Public Function FirstWord(ByVal vWord As String) As String
Dim i As Integer
    i = InStr(1, vWord, " ")
    If i = 0 Then
        FirstWord = vWord
    Else
        FirstWord = Mid(vWord, 1, i - 1)
    End If
End Function
