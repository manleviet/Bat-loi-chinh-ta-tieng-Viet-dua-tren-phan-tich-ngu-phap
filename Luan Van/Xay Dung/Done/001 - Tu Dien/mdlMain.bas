Attribute VB_Name = "mdlMain"
'Project: Vietnamese Checking
'Description: mdlMain Modul - Main Modul
'--------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Public WC As New clsWordClass
Public Dic As New clsWordDic

Sub main()
Dim e As Integer
    Call GetInformation
    e = Dic.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        Exit Sub
    End If
    e = WC.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        Exit Sub
    End If
    frmWManage.Show
End Sub
