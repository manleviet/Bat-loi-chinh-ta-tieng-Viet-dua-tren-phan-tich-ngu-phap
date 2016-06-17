Attribute VB_Name = "mdlMain"
'Project: Vietnamese Checking
'Description: mdlMain Modul - a Main Modul
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

Public Dic As clsWordDic
Public WC As clsWordClass
Public RDic As crlRule

Sub Main()
Dim e As Integer
    Call InitUnicode
    Call GetInformation
    
    Set Dic = New clsWordDic
    e = Dic.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        End
    End If

    Set WC = New clsWordClass
    e = WC.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        End
    End If
    
    Set RDic = New crlRule
    e = RDic.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        Exit Sub
    End If
    
    frmTHTLDung.Show
End Sub
