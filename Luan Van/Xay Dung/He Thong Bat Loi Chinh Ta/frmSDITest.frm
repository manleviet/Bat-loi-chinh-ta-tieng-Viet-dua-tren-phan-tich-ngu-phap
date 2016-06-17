VERSION 5.00
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmSDITest 
   Caption         =   "He Thong Bat Loi Chinh Ta"
   ClientHeight    =   5490
   ClientLeft      =   1545
   ClientTop       =   1275
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSDITest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8115
   Begin vbalTBar6.cToolbar cToolbar1 
      Height          =   1125
      Left            =   120
      Top             =   960
      Width           =   3000
      _ExtentX        =   5424
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbar tbrMenu 
      Height          =   1125
      Left            =   60
      Top             =   720
      Width           =   3000
      _ExtentX        =   5424
      _ExtentY        =   661
   End
   Begin vbalTBar6.cReBar cReBar1 
      Left            =   60
      Top             =   60
      _ExtentX        =   13785
      _ExtentY        =   979
   End
   Begin VB.PictureBox picTest 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   7380
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   1560
      Width           =   315
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   0
         Picture         =   "frmSDITest.frx":23D2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
   End
   Begin vbalIml6.vbalImageList vbalImageList2 
      Left            =   5070
      Top             =   840
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   1148
      Images          =   "frmSDITest.frx":251C
      Version         =   131072
      KeyCount        =   1
      Keys            =   ""
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   4395
      Top             =   855
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   24
      IconSizeY       =   24
      ColourDepth     =   32
      Size            =   39360
      Images          =   "frmSDITest.frx":29B8
      Version         =   131072
      KeyCount        =   16
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbalIml6.vbalImageList vbaImageMenuToolbar 
      Left            =   3720
      Top             =   840
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   18368
      Images          =   "frmSDITest.frx":C398
      Version         =   131072
      KeyCount        =   16
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Image picAnh 
      Height          =   4920
      Left            =   120
      Picture         =   "frmSDITest.frx":10B78
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   6420
   End
End
Attribute VB_Name = "frmSDITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&

Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1

Private Sub TestSaveLayout()
Dim iFile As Integer
Dim sFile As String
Dim sXml As String
   
    sXml = cReBar1.SaveLayout()
    sFile = App.Path & "\restore.xml"
   
    On Error Resume Next
    Kill sFile
   
    On Error GoTo errorHandler
    iFile = FreeFile
    Open sFile For Binary Access Write As #iFile
    Put #iFile, , sXml
    Close #iFile
    Exit Sub
errorHandler:
    MsgBox "An error occurred trying to save the layout:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub TestLoadLayout()
Dim iFile As Integer
Dim sXml As String
    iFile = FreeFile
    Open App.Path & "\restore.xml" For Binary Access Read As #iFile
    If FileLen(App.Path & "\restore.xml") <> 0 Then
        sXml = Space$(LOF(iFile))
        Get #iFile, , sXml
        Close #iFile
   
        ReDim sData(1 To 3) As String
        ReDim lhWnd(1 To 3) As Long
        sData(1) = "MenuBar": lhWnd(1) = tbrMenu.hwnd
        sData(2) = "Logo": lhWnd(2) = picTest.hwnd
        sData(3) = "Toolbar1": lhWnd(3) = cToolbar1.hwnd

        cReBar1.DestroyRebarDontDestroyChildren
        cReBar1.CreateRebar Me.hwnd
        cReBar1.RestoreLayout sXml, sData(), lhWnd()
    End If
End Sub

Private Sub pCreateMenu()
Dim iP As Long
Dim iP2 As Long
    Set m_cMenu = New cPopupMenu
    With m_cMenu
        .ImageList = vbaImageMenuToolbar.hIml
        .hWndOwner = Me.hwnd
        .OfficeXpStyle = True
                  
        iP = .AddItem("He Thong", , , , , , , "mnuHeThongTOP")
        .AddItem "Thoat", , , iP, 0, , , "mnuHeThong(0)"
        
        iP = .AddItem("Bat Loi Chinh Ta", , , , , , , "mnuKiemTraTOP")
        .AddItem "Bat Loi Chinh Ta Muc Am Tiet", , , iP, 1, , , "mnuKiemTra(0)"
        .AddItem "Bat Loi Chinh Ta Muc Cu Phap", , , iP, 2, , , "mnuKiemTra(1)"
        
        iP = .AddItem("Tu Dien", , , , , , , "mnuTuDienTOP")
        .AddItem "Quan Ly Tu", , , iP, 3, , , "mnuTuDien(0)"
        .AddItem "Quan Ly Tu Loai", , , iP, 4, , , "mnuTuDien(1)"
        .AddItem "Hoc Am Tiet", , , iP, 5, , , "mnuTuDien(4)"
        .AddItem "Hoc Tu", , , iP, 6, , , "mnuTuDien(2)"
       
        iP = .AddItem("Luat Sinh", , , , , , , "mnuLuatSinhTOP")
        .AddItem "Quan Ly Luat Sinh", , , iP, 7, , , "mnuLuatSinh(0)"
               
        iP = .AddItem("Mo Phong", , , , , , , "mnuMoPhongTOP")
        .AddItem "Giai Doan Tach Doan", , , iP, 8, , , "mnuMoPhong(6)"
        .AddItem "Giai Doan Tach Cau", , , iP, 9, , , "mnuMoPhong(0)"
        .AddItem "Giai Doan Tach Am Tiet", , , iP, 10, , , "mnuMoPhong(1)"
        .AddItem "Giai Doan Tach Tu", , , iP, 11, , , "mnuMoPhong(2)"
        .AddItem "Giai Doan Gop Tu", , , iP, 12, , , "mnuMoPhong(3)"
        .AddItem "Giai Doan To Hop Tu Loai", , , iP, 13, , , "mnuMoPhong(4)"
        .AddItem "Chuoi Tu Loai Dung", , , iP, 14, , , "mnuMoPhong(5)"
        
        iP = .AddItem("Tuy Chon", , , , , , , "mnuTuyChonTOP")
        .AddItem "Giao Dien Kieu XP", , , iP, , True, , "mnuTuyChon(0)"
        .AddItem "-", , , iP, , , , "mnuTuyChon(1)"
        iP2 = .AddItem("Ben Tren", , , iP, , , , "mnuTuyChon(2)")
        .RadioCheck(iP2) = True
        .AddItem "Ben Trai", , , iP, , , , "mnuTuyChon(3)"
        .AddItem "Ben Phai", , , iP, , , , "mnuTuyChon(4)"
        .AddItem "Ben Duoi", , , iP, , , , "mnuTuyChon(5)"
        
        iP = .AddItem("Tro Giup", , , , , , , "mnuTroGiupTOP")
        .AddItem "Ve Tac Gia", , , iP, 15, , , "mnuTroGiup(0)"
   End With
End Sub

Private Sub cReBar1_ChevronPushed(ByVal wID As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Dim lIndex As Long
Dim v As Variant
    v = cReBar1.BandData(wID)
    If Not IsMissing(v) Then
        Select Case v
            Case "MenuBar"
                tbrMenu.ChevronPress lRight \ Screen.TwipsPerPixelX + 1, lTop \ Screen.TwipsPerPixelY
            Case "Toolbar1"
                cToolbar1.ChevronPress lRight \ Screen.TwipsPerPixelX + 1, lTop \ Screen.TwipsPerPixelY
        End Select
    End If
End Sub

Private Sub cReBar1_HeightChanged(lHeight As Long)
   pResize
End Sub

Private Sub cToolbar1_ButtonClick(ByVal lButton As Long)
Dim sKey As String
    sKey = cToolbar1.ButtonKey(lButton)
    Select Case sKey
        Case "Thoat"
            PostMessage Me.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
        Case "BatLoi0"
            Shell App.Path & "\VietnameseChecking"
        Case "BatLoi1"
            Shell App.Path & "\GrammarChecking"
        Case "TuDien0"
            Shell App.Path & "\WordManage"
        Case "TuDien1"
            Shell App.Path & "\WordClassManage"
        Case "TuDien2"
            Shell App.Path & "\SyllableStudy"
        Case "TuDien3"
            Shell App.Path & "\WordStudy"
        Case "Luat"
            Shell App.Path & "\RuleManage"
        Case "MoPhong0"
            Shell App.Path & "\ParagraphSplit"
        Case "MoPhong1"
            Shell App.Path & "\SentenceSplit"
        Case "MoPhong2"
            Shell App.Path & "\SyllableSplit"
        Case "MoPhong3"
            Shell App.Path & "\WordSplit"
        Case "MoPhong4"
            Shell App.Path & "\WordAddUp"
        Case "MoPhong5"
            Shell App.Path & "\WordClassSet"
        Case "MoPhong6"
            Shell App.Path & "\EarlyParse"
        Case "About"
            frmAbout.Show
    End Select
End Sub

Private Sub Form_Load()
Dim lMajor As Long, lMinor As Long, lBuild As Long
Dim i As Long
    
    Me.Caption = "He Thong Bat Loi Chinh Ta - Phien Ban: " & App.Major & "." & App.Minor
    cToolbar1.GetComCtrlVersionInfo lMajor, lMinor, lBuild
    Call pCreateMenu
   
    With cToolbar1
        .ImageSource = CTBExternalImageList
        .SetImageList vbalImageList1, CTBImageListNormal
        .CreateToolbar 24, , , True, True
        .DrawStyle = CTBDrawOfficeXPStyle
        .AddButton "Thoat", 0, , , , CTBNormal, "Thoat"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Bat Loi Chinh Ta Muc Am Tiet", 1, , , , CTBNormal, "BatLoi0"
        .AddButton "Bat Loi Chinh Ta Muc Cu Phap", 2, , , , CTBNormal, "BatLoi1"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Quan Ly Tu Dien Tu", 3, , , , CTBNormal, "TuDien0"
        .AddButton "Quan Ly Tu Dien Tu Loai", 4, , , , CTBNormal, "TuDien1"
        .AddButton "Hoc Am Tiet", 5, , , , CTBNormal, "TuDien2"
        .AddButton "Hoc Tu", 6, , , , CTBNormal, "TuDien3"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Quan Ly Luat Sinh", 7, , , , CTBNormal, "Luat"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Mo Phong Giai Doan Tach Doan", 8, , , , CTBNormal, "MoPhong0"
        .AddButton "Mo Phong Giai Doan Tach Cau", 9, , , , CTBNormal, "MoPhong1"
        .AddButton "Mo Phong Giai Doan Tach Am Tiet", 10, , , , CTBNormal, "MoPhong2"
        .AddButton "Mo Phong Giai Doan Tach Tu", 11, , , , CTBNormal, "MoPhong3"
        .AddButton "Mo Phong Giai Doan Gop Tu", 12, , , , CTBNormal, "MoPhong4"
        .AddButton "Mo Phong Giai Doan To Hop Tu Loai", 13, , , , CTBNormal, "MoPhong5"
        .AddButton "Chuoi Tu Loai Dung", 14, , , , CTBNormal, "MoPhong6"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Ve Tac Gia", 15, , , , CTBNormal, "About"
    End With
    cToolbar1.DropDownAlign = CTBDropDownAlignBottom
   
    tbrMenu.DrawStyle = CTBDrawOfficeXPStyle
    tbrMenu.CreateFromMenu m_cMenu
    tbrMenu.DropDownAlign = CTBDropDownAlignBottom
    With cReBar1
        .ImageSource = CRBLoadFromFile
        .CreateRebar Me.hwnd
        .AddBandByHwnd tbrMenu.hwnd, , , , "MenuBar"
        .BandChildMinWidth(.BandCount - 1) = 24
        .AddBandByHwnd cToolbar1.hwnd, , , , "Toolbar1"
        .BandChildMinWidth(.BandCount - 1) = 24
        If (lMajor > 4) Or (lMinor > 70) Then          ' Fixed toolbars are not allowed with COMMCTRL before 4.71
            .AddBandByHwnd picTest.hwnd, , False, True, "Logo"
        End If
    End With
    Call TestLoadLayout
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call TestSaveLayout
    cReBar1.RemoveAllRebarBands
End Sub

Private Sub pResize()
Dim lH As Long
On Error Resume Next
    lH = (cReBar1.RebarHeight * Screen.TwipsPerPixelY)
    Select Case cReBar1.Position
    Case erbPositionTop
        picAnh.Move Screen.TwipsPerPixelX * 2, lH + 3 * Screen.TwipsPerPixelY, _
            Me.ScaleWidth - 4 * Screen.TwipsPerPixelX, Me.ScaleHeight - lH - 5 * Screen.TwipsPerPixelY
   Case erbPositionRight
      picAnh.Move Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY * 2, _
         Me.ScaleWidth - lH - 4 * Screen.TwipsPerPixelX, Me.ScaleHeight - 4 * Screen.TwipsPerPixelY
   Case erbPositionLeft
      picAnh.Move lH + 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, _
         Me.ScaleWidth - lH - 4 * Screen.TwipsPerPixelX, Me.ScaleHeight - 4 * Screen.TwipsPerPixelY
   Case erbPositionBottom
      picAnh.Move Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY * 2, _
         Me.ScaleWidth - 2 * Screen.TwipsPerPixelX, Me.ScaleHeight - lH - 5 * Screen.TwipsPerPixelY
   End Select
End Sub
Private Sub Form_Resize()
   cReBar1.RebarSize
   pResize
End Sub

Private Sub Form_Terminate()
   If (Forms.Count = 0) Then
      UnloadApp
   End If
End Sub

Private Sub m_cMenu_Click(ItemNumber As Long)
Dim sKey As String
Dim WinWnd As Long, Ret As String
Dim bs As Boolean
    sKey = m_cMenu.ItemKey(ItemNumber)
    Select Case True
        Case sKey = "mnuHeThong(0)"
            PostMessage Me.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
        Case sKey = "mnuKiemTra(0)"
            Shell App.Path & "\VietnameseChecking"
        Case sKey = "mnuKiemTra(1)"
            Shell App.Path & "\GrammarChecking"
        Case sKey = "mnuTuDien(0)"
            Shell App.Path & "\WordManage"
        Case sKey = "mnuTuDien(1)"
            Shell App.Path & "\WordClassManage"
        Case sKey = "mnuTuDien(2)"
            Shell App.Path & "\SyllableStudy"
        Case sKey = "mnuTuDien(3)"
            Shell App.Path & "\WordStudy"
        Case sKey = "mnuLuatSinh(0)"
            Shell App.Path & "\RuleManage"
        Case sKey = "mnuMoPhong(6)"
            Shell App.Path & "\ParagraphSplit"
        Case sKey = "mnuMoPhong(0)"
            Shell App.Path & "\SentenceSplit"
        Case sKey = "mnuMoPhong(1)"
            Shell App.Path & "\SyllableSplit"
        Case sKey = "mnuMoPhong(2)"
            Shell App.Path & "\WordSplit"
        Case sKey = "mnuMoPhong(3)"
            Shell App.Path & "\WordAddUp"
        Case sKey = "mnuMoPhong(4)"
            Shell App.Path & "\WordClassSet"
        Case sKey = "mnuMoPhong(5)"
            Shell App.Path & "\EarlyParse"
        Case sKey = "mnuTuyChon(0)"
            bs = Not (m_cMenu.Checked(ItemNumber))
            m_cMenu.Checked(ItemNumber) = bs
            If bs Then
                cToolbar1.DrawStyle = CTBDrawOfficeXPStyle
            Else
                cToolbar1.DrawStyle = CTBDrawStandard
            End If
        Case sKey = "mnuTuyChon(2)"
            pPositionMenu ItemNumber, sKey
        Case sKey = "mnuTuyChon(3)"
            pPositionMenu ItemNumber, sKey
        Case sKey = "mnuTuyChon(4)"
            pPositionMenu ItemNumber, sKey
        Case sKey = "mnuTuyChon(5)"
            pPositionMenu ItemNumber, sKey
        Case sKey = "mnuTroGiup(0)"
            frmAbout.Show Normal
   End Select
End Sub

Private Function calcToolbarHeight(ByRef cT As cToolbar) As Long
Dim i As Long
   i = cT.ButtonCount - 1
   calcToolbarHeight = cT.ButtonTop(i) + cT.ButtonHeight(i)
End Function

Private Sub pPositionMenu(ByVal lIndex As Long, ByVal sKey As String)
Dim i As Long
Dim lItemIndex As Long

   LockWindowUpdate Me.hwnd
   lItemIndex = CLng(Mid(sKey, 12, 1))
   Select Case lItemIndex
   Case 2 ' top
      cReBar1.BandVisible(cReBar1.BandIndexForData("picBar")) = True
      cReBar1.BandChildMinHeight(cReBar1.BandIndexForData("MenuBar")) = tbrMenu.ToolbarHeight
      cReBar1.Position = erbPositionTop
   Case 3 ' left
      cReBar1.BandVisible(cReBar1.BandIndexForData("picBar")) = False
      cReBar1.BandChildMinHeight(cReBar1.BandIndexForData("MenuBar")) = tbrMenu.MaxButtonWidth
      cReBar1.Position = erbPositionLeft
   Case 4 ' right
      cReBar1.BandVisible(cReBar1.BandIndexForData("picBar")) = False
      cReBar1.BandChildMinHeight(cReBar1.BandIndexForData("MenuBar")) = tbrMenu.MaxButtonWidth
      cReBar1.Position = erbPositionRight
   Case 5 ' bottom
      cReBar1.BandVisible(cReBar1.BandIndexForData("picBar")) = True
    'cReBar1.BandChildMinHeight(cReBar1.BandIndexForData("MenuBar")) = tbrMenu.ToolbarHeight
      cReBar1.Position = erbPositionBottom
   End Select
   If lItemIndex = 3 Or lItemIndex = 4 Then
      'cReBar1.BandChildMinHeight(cReBar1.BandIndexForData("Toolbar1")) = cToolbar1.ButtonWidth(4)
      cReBar1.BandChildIdealWidth(cReBar1.BandIndexForData("MenuBar")) = calcToolbarHeight(tbrMenu)
      cReBar1.BandChildIdealWidth(cReBar1.BandIndexForData("Toolbar1")) = calcToolbarHeight(cToolbar1)
      tbrMenu.DropDownAlign = CTBDropDownAlignLeft
      cToolbar1.DropDownAlign = CTBDropDownAlignLeft
   Else
      'cReBar1.BandChildMinHeight(cReBar1.BandIndexForData("Toolbar1")) = cToolbar1.MaxButtonHeight
      cReBar1.BandChildIdealWidth(cReBar1.BandIndexForData("MenuBar")) = tbrMenu.ToolbarWidth
      cReBar1.BandChildIdealWidth(cReBar1.BandIndexForData("Toolbar1")) = cToolbar1.ToolbarWidth
      tbrMenu.DropDownAlign = CTBDropDownAlignBottom
      cToolbar1.DropDownAlign = CTBDropDownAlignBottom
   End If
   m_cMenu.GroupToggle m_cMenu.IndexForKey("mnuTuyChon(" & lItemIndex & ")")
   LockWindowUpdate 0
End Sub


