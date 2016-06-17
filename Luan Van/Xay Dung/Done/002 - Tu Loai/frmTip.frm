VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   9150
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Desription: frmTip Form - a form for the tip
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
    Private Const WS_CAPTION = &HC00000            ' WS_BORDER Or WS_DLGFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Function ShowTitleBar(ByVal bState As Boolean)
Dim lStyle As Long
Dim tR As RECT

    ' Get the window's position:
    GetWindowRect Me.hwnd, tR

    ' Modify whether title bar will be visible:
    lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    If (bState) Then
    Me.Caption = Me.Tag
    If Me.ControlBox Then
        lStyle = lStyle Or WS_SYSMENU
    End If
    If Me.MaxButton Then
        lStyle = lStyle Or WS_MAXIMIZEBOX
    End If
    If Me.MinButton Then
        lStyle = lStyle Or WS_MINIMIZEBOX
    End If
    If Me.Caption <> "" Then
        lStyle = lStyle Or WS_CAPTION
    End If
    Else
    Me.Tag = Me.Caption
    Me.Caption = ""
    lStyle = lStyle And Not WS_SYSMENU
    lStyle = lStyle And Not WS_MAXIMIZEBOX
    lStyle = lStyle And Not WS_MINIMIZEBOX
    lStyle = lStyle And Not WS_CAPTION
End If
SetWindowLong Me.hwnd, GWL_STYLE, lStyle

' Ensure the style takes and make the window the
' same size, regardless that the title bar etc
' is now a different size:
SetWindowPos Me.hwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
Me.Refresh

' Ensure that your resize code is fired, as the client area
' has changed:
'Form_Resize

End Function

Private Sub Form_Load()
    ShowTitleBar False
End Sub
