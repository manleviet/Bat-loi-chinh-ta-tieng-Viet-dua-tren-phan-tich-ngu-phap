Attribute VB_Name = "mdlMain"
'Project: Vietnamese Checking
'Description: mdlMain Modul - Main Modul
'--------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Public WC As clsWordClass
'Public Dic As New clsWordDic

Sub main()
    Call GetInformation
    Call InitUnicode
    Call InitTCVN
    frmWCManage.Show
End Sub

