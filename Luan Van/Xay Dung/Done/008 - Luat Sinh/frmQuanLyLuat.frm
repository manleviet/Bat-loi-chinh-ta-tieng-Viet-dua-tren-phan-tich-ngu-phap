VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQuanLyLuat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuanLyLuat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8655
      Begin MSForms.TextBox txtAn 
         Height          =   135
         Left            =   1200
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         VariousPropertyBits=   746604571
         Size            =   "2143;238"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame1 
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1695
         Caption         =   "Thong tin tu loai"
         Size            =   "2990;423"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtVeTrai 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         VariousPropertyBits=   746604571
         Size            =   "2566;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblYnghia 
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   735
         Caption         =   "Y Nghia:"
         Size            =   "1296;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtVePhai 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   360
         Width           =   5055
         VariousPropertyBits=   746604571
         Size            =   "8916;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdLuu 
         Height          =   375
         Left            =   5520
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         Caption         =   "L"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdBoqua 
         Height          =   375
         Left            =   7080
         TabIndex        =   4
         Top             =   840
         Width           =   1335
         Caption         =   "Bo qua"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblMaloai 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   975
         Caption         =   "Ma loai:"
         Size            =   "1720;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   8655
      Begin MSForms.Label lblFrame2 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   1695
         Caption         =   "Danh sach tu loai"
         Size            =   "2990;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdDong 
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
         Caption         =   "Dong"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSuadoi 
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
         Caption         =   "Sua Doi"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ListBox lstLuat 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   6735
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "11880;5318"
         MatchEntry      =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdXoa 
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
         Caption         =   "Xoa"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdTaoMoi 
         Height          =   375
         Left            =   7080
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
         Caption         =   "Tao Moi"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblTongso 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   5415
         Caption         =   "Tong so:"
         Size            =   "9551;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblMaloai1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2655
         Caption         =   "Ma loai -"
         Size            =   "4683;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.Label lblTitle 
      Height          =   240
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   7560
      VariousPropertyBits=   276824091
      Caption         =   "Nghien cuu va phat trien phuong phap bat loi chinh ta Tieng Viet"
      Size            =   "13335;423"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblHeader 
      Height          =   345
      Left            =   3323
      TabIndex        =   17
      Top             =   360
      Width           =   2235
      ForeColor       =   8421631
      Caption         =   "QUAN LY LUAT"
      Size            =   "3942;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmQuanLyLuat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Desription: frmQuanLyLuat Form - a form managing rule
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Dim NoE As Integer ' 1 - New, 2 - Edit
Dim OnWCInfo As Boolean

Private Sub ChangeStatus(ByVal bState As Boolean)
    cmdLuu.Enabled = bState
    cmdBoqua.Enabled = bState
    cmdTaoMoi.Enabled = Not bState
    cmdXoa.Enabled = Not bState
    cmdSuadoi.Enabled = Not bState
End Sub

Private Sub cmdBoqua_Click()
    NoE = 0
    OnWCInfo = False
    Call ChangeStatus(False)
    lstLuat.SetFocus
End Sub

Private Sub cmdDong_Click()
    Unload frmQuanLyLuat
End Sub

Private Sub cmdLuu_Click()
Dim st As String
Dim e As Integer, i As Long
Dim loca As New clsLocation
    DoEvents
    If txtVeTrai.Text = "" Or txtVePhai.Text = "" Then
        MsgBox "Khong the rong"
        Exit Sub
    End If
    st = Trim(txtVeTrai.Text) & " " & Trim(txtVePhai.Text)
    If NoE = 1 Then
        e = RDic.AddRule(st)
    ElseIf NoE = 2 Then
        Set loca = RDic.FindRule(Trim(txtAn.Text))
        If loca.ok = 0 Then
            RDic.Rule(loca.x, loca.y) = Trim(txtVeTrai.Text) & " " & Trim(txtVePhai.Text)
            e = 0
        Else
            e = eError.NoHaveWord
        End If
    End If
    If e <> 0 Then
        Call ErrorHandle(e)
        'Unload frmWCManage
    Else
        Call Load4lstLuat
        For i = 0 To lstLuat.ListCount - 1
            If Left(lstLuat.List(i), 3) = Trim(txtVeTrai) Then Exit For
        Next i
        lstLuat.ListIndex = i
        
        e = RDic.SaveDic
        If e <> 0 Then Call ErrorHandle(e)
    End If
    NoE = 0
    OnWCInfo = False
    Call ChangeStatus(False)
End Sub

Private Sub cmdSuadoi_Click()
    NoE = 2
    OnWCInfo = True
    Call ChangeStatus(True)
    txtVeTrai.SelStart = 0
    txtVeTrai.SelLength = Len(txtVeTrai.Text)
    txtVeTrai.SetFocus
End Sub

Private Sub cmdTaomoi_Click()
    NoE = 1
    OnWCInfo = True
    Call ChangeStatus(True)
    txtVeTrai.Text = ""
    txtVePhai.Text = ""
    txtVeTrai.SetFocus
End Sub

Private Sub cmdXoa_Click()
Dim e As Integer
    If txtVeTrai.Text <> "" Then
        e = RDic.DelRule(txtAn.Text)
        If e <> 0 Then Call ErrorHandle(e)
    End If
    Call Load4lstLuat
End Sub

Private Sub Form_Load()

    DoEvents
    Me.Show
    Me.Refresh
    frmFlash.lblCaption.Caption = ChrW(84) & ChrW(104) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(7863) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    
    Call AddCaption
    
    NoE = 0
    OnWCInfo = False
    
    Call ChangeStatus(False)
    Call Load4lstLuat
    Unload frmFlash
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim e As Integer
'    e = RDic.SaveDic
'    If e <> 0 Then
'        Call ErrorHandle(e)
'    Else
'        Set RDic = Nothing
'        Set RDic = New crlRule
'        e = RDic.LoadDic
'        If e <> 0 Then Call ErrorHandle(e)
'    End If
End Sub

Private Sub Load4lstLuat()
Dim i As Long, j As Long
Dim st As String
    DoEvents
    lstLuat.Clear
    For i = 1 To RDic.PCount
        For j = 1 To RDic.RiPCount(i)
            st = RDic.Rule(i, j)
            lstLuat.AddItem Left(st, 3) & " --> " & Mid(st, 5)
        Next j
    Next i
    lblTongso.Caption = ChrW(84) & ChrW(7893) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(32) & ChrW(108) & ChrW(117) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(115) & ChrW(105) & ChrW(110) & ChrW(104) & ChrW(58) & " " & RDic.RCount
End Sub

Private Sub Load4WCInfo()
Dim i As Long
Dim st As String
    DoEvents
    i = lstLuat.ListIndex
    If i <> -1 Then
        st = lstLuat.List(i)
        txtAn.Text = Left(st, 3) & " " & Mid(st, 9)
        txtVeTrai.Text = Left(st, 3)
        txtVePhai.Text = Mid(st, 9)
    End If
End Sub

Private Sub lstLuat_Click()
    Call Load4WCInfo
End Sub

Private Sub AddCaption()
    Me.Caption = "Quan Ly Luat Sinh - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(81) & ChrW(85) & ChrW(7842) & ChrW(78) & ChrW(32) & ChrW(76) & ChrW(221) & ChrW(32) & ChrW(76) & ChrW(85) & ChrW(7852) & ChrW(84) & ChrW(32) & ChrW(83) & ChrW(73) & ChrW(78) & ChrW(72)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(84) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(108) & ChrW(117) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(115) & ChrW(105) & ChrW(110) & ChrW(104)
    lblMaloai.Caption = ChrW(86) & ChrW(7871) & ChrW(32) & ChrW(84) & ChrW(114) & ChrW(225) & ChrW(105) & ChrW(58)
    lblYnghia.Caption = ChrW(86) & ChrW(7871) & ChrW(32) & ChrW(80) & ChrW(104) & ChrW(7843) & ChrW(105) & ChrW(58)
    cmdLuu.Caption = ChrW(76) & ChrW(432) & ChrW(117)
    cmdBoqua.Caption = ChrW(66) & ChrW(7887) & ChrW(32) & ChrW(81) & ChrW(117) & ChrW(97)
    lblFrame2.Caption = ChrW(68) & ChrW(97) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(115) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(108) & ChrW(117) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(115) & ChrW(105) & ChrW(110) & ChrW(104)
    lblMaloai1.Caption = ChrW(86) & ChrW(7871) & ChrW(32) & ChrW(84) & ChrW(114) & ChrW(225) & ChrW(105) & ChrW(32) & ChrW(45) & ChrW(45) & ChrW(62) & ChrW(32) & ChrW(86) & ChrW(7871) & ChrW(32) & ChrW(80) & ChrW(104) & ChrW(7843) & ChrW(105)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    cmdTaoMoi.Caption = ChrW(84) & ChrW(7841) & ChrW(111) & ChrW(32) & ChrW(77) & ChrW(7899) & ChrW(105)
    cmdSuadoi.Caption = ChrW(83) & ChrW(7917) & ChrW(97) & ChrW(32) & ChrW(272) & ChrW(7893) & ChrW(105)
    cmdXoa.Caption = ChrW(88) & ChrW(111) & ChrW(225)
End Sub

Private Sub txtVeTrai_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtVePhai.SetFocus
    End If
End Sub

Private Sub txtVePhai_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdLuu_Click
    End If
End Sub
