Attribute VB_Name = "mdlMain"
'Project: Vietnamese Checking
'Description: mdlMain Modul - a Main Modul
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

'Public Dic As clsWordDic
'Public WC As clsWordClass

'Public TypingStyle As String
'Public HaveMDIChild As Boolean
'Public mForm As Form

Sub Main()
Dim e As Integer
'    Call InitUnicode
'    Call GetInformation
    
'    HaveMDIChild = False
'    TypingStyle = "Telex"
      
'    Set Dic = New clsWordDic
'    e = Dic.LoadDic
'    If e <> 0 Then
'        Call ErrorHandle(e)
'        End
'    End If

'    Set WC = New clsWordClass
'    e = WC.LoadDic
'    If e <> 0 Then
'        Call ErrorHandle(e)
'        End
'    End If
    
    frmTachCau.Show
End Sub
