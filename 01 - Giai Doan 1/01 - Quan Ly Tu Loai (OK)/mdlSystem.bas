Attribute VB_Name = "mdlSystem"
'Project: Vietnamese Checking
'Description: mdlSystem Modul - a modul for system functoins
'-------------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

Public Declare Function GetTickCount& Lib "kernel32" ()

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Const SW_SHOWNORMAL = 1
Const WM_CLOSE = &H10

Public Sub TeminateWindows(WindowsTitle As String)
Dim WinWnd As Long, Ret As String, RetVal As Long, lpClassName As String
    Ret = WindowsTitle
    'Search the window
    WinWnd = FindWindow(vbNullString, Ret)
    If WinWnd = 0 Then Exit Sub
    'Show the window
    ShowWindow WinWnd, SW_SHOWNORMAL
    'Create a buffer
    lpClassName = Space(256)
    'retrieve the class name
    RetVal = GetClassName(WinWnd, lpClassName, 256)
    'Post a message to the window to close itself
    PostMessage WinWnd, WM_CLOSE, 0&, 0&
End Sub

'Public Sub ShowSplash(Msg As String, Optional Title As String, Optional Delay As Long)
'    With frmSplash
'        .lblMessage.Caption = Msg
'        .Caption = Title
'        .Show
'        .Refresh
'    End With
'End Sub

Public Sub ErrorHandle(ByVal e As EError)
    Select Case e
        Case EError.AddPageError: MsgBox "Error when adding a dictionary page."
        Case EError.AddWordError: MsgBox "Error when adding word to dictionary."
        Case EError.ClearCPageError: MsgBox "Error when clear a dictionary page."
        Case EError.CopyCPageError: MsgBox "Error when copy a dictionary page."
        Case EError.CopyWordError: MsgBox "Error when copy a word."
        Case EError.DelError: MsgBox "Error when del."
        Case EError.GetWordError: MsgBox "Error when get a word."
        Case EError.LoadDicError: MsgBox "Error when load dictionary."
        Case EError.NoHaveWord: MsgBox "No have that word."
        Case EError.SaveDicError: MsgBox "Error when save dictionary."
        Case EError.SetWordError: MsgBox "Error when set a word."
        Case EError.SortError: MsgBox "Error when sort dictionary."
        Case EError.SwapError: MsgBox "Error when swap two words."
        Case EError.TheSameWord: MsgBox "The Same Word."
        Case Else: MsgBox "Error."
    End Select
End Sub
