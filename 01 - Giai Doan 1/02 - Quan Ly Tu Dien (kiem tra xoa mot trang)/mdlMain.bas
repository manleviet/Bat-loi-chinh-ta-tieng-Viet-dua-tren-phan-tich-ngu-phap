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
    Call InitUnicode
    
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(76) & ChrW(111) & ChrW(97) & ChrW(100) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    
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
    
    Unload frmFlash
    frmWManage.Show
End Sub
