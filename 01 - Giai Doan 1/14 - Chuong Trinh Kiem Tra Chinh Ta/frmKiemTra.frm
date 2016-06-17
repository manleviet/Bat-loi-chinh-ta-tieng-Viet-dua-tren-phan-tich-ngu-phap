VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmKiemTra 
   Caption         =   "Kiem Tra Chinh Ta"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   Icon            =   "frmKiemTra.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      Begin MSForms.CommandButton cmdKiemTra 
         Height          =   375
         Left            =   6840
         TabIndex        =   6
         Top             =   5400
         Width           =   1335
         Caption         =   "Kiem Tra"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDong 
         Height          =   375
         Left            =   8400
         TabIndex        =   5
         Top             =   5400
         Width           =   1335
         Caption         =   "Dong"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblFrame1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   5175
         Caption         =   "Nhap van ban can kiem tra roi nhan nut Kiem Tra ben duoi"
         Size            =   "9128;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtThu 
         Height          =   4935
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9735
         VariousPropertyBits=   -1399830501
         ScrollBars      =   3
         Size            =   "17171;8705"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.Label lblHeader 
      Height          =   345
      Left            =   2655
      TabIndex        =   1
      Top             =   360
      Width           =   4890
      ForeColor       =   8421631
      Caption         =   "KIEM TRA CHINH TA TIENG VIET"
      Size            =   "8625;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblTitle 
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Caption         =   "Nghien cuu va phat trien phuong phap bat loi chinh ta Tieng Viet"
      Size            =   "11245;423"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmKiemTra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sen As chkChecking

Private Sub cmdDong_Click()
    Unload frmKiemTra
End Sub

Private Sub cmdKiemTra_Click()
    Set sen = New chkChecking
    sen.Init txtThu
    If sen.SyllableCheck <> 0 Then Call ErrorHandle(e)
End Sub

Private Sub AddCaption()
    Me.Caption = "Kiem Tra Chinh Ta Tieng Viet - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(75) & ChrW(73) & ChrW(7874) & ChrW(77) & ChrW(32) & ChrW(84) & ChrW(82) & ChrW(65) & ChrW(32) & ChrW(67) & ChrW(72) & ChrW(205) & ChrW(78) & ChrW(72) & ChrW(32) & ChrW(84) & ChrW(7842) & ChrW(32) & ChrW(84) & ChrW(73) & ChrW(7870) & ChrW(78) & ChrW(71) & ChrW(32) & ChrW(86) & ChrW(73) & ChrW(7878) & ChrW(84)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(105) & ChrW(7875) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(97) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(250) & ChrW(116) & ChrW(32) & ChrW(75) & ChrW(105) & ChrW(7875) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(114) & ChrW(97) & ChrW(32) & ChrW(98) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
    cmdKiemTra.Caption = ChrW(75) & ChrW(105) & ChrW(7875) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(114) & ChrW(97)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
End Sub

Private Sub Form_Load()
    Call AddCaption
End Sub
