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
    frmWCManage.Show
End Sub

Public Sub ErrorHandle(ByVal e As eError)
    Select Case e
        Case eError.AddPageError: MsgBox "Error when adding a dictionary page."
        Case eError.AddWordError: MsgBox "Error when adding word to dictionary."
        Case eError.ClearCPageError: MsgBox "Error when clear a dictionary page."
        Case eError.CopyCPageError: MsgBox "Error when copy a dictionary page."
        Case eError.CopyWordError: MsgBox "Error when copy a word."
        Case eError.DelError: MsgBox "Error when del."
        Case eError.GetWordError: MsgBox "Error when get a word."
        Case eError.LoadDicError: MsgBox "Error when load dictionary."
        Case eError.NoHaveWord: MsgBox "No have that word."
        Case eError.SaveDicError: MsgBox "Error when save dictionary."
        Case eError.SetWordError: MsgBox "Error when set a word."
        Case eError.SortError: MsgBox "Error when sort dictionary."
        Case eError.SwapError: MsgBox "Error when swap two words."
        Case Else: MsgBox "Error."
    End Select
End Sub

