VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4665
   ClientLeft      =   4110
   ClientTop       =   3150
   ClientWidth     =   6390
   ClipControls    =   0   'False
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3219.866
   ScaleMode       =   0  'User
   ScaleWidth      =   6000.541
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      Top             =   4200
      Width           =   1260
   End
   Begin MSForms.Label Label4 
      Height          =   255
      Left            =   1568
      TabIndex        =   14
      Top             =   3240
      Width           =   3015
      BackColor       =   16777215
      Caption         =   "tren co so do lien ket chat che"
      Size            =   "5318;450"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblDiaChi 
      Height          =   255
      Left            =   1088
      TabIndex        =   13
      Top             =   3000
      Width           =   4215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Kiem tra tinh chinh xac cua van ban tieng Viet"
      Size            =   "7435;450"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblTruong 
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   4320
      Width           =   2415
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Truong Dai Hoc Khoa Hoc - Hue"
      Size            =   "4260;450"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblLop 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Width           =   1815
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Khoa Toan"
      Size            =   "3201;450"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblTacGia 
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3840
      Width           =   3615
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Chu nhiem de tai: PGS.TS. Nguyen Gia Dinh"
      Size            =   "6376;450"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblChinhTa 
      Height          =   735
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   2895
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Tieng Viet"
      Size            =   "5106;1296"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   555
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblBatLoi 
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   4575
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Bat Loi Chinh Ta"
      Size            =   "8070;1085"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   555
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblHeThong 
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   2775
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "He Thong"
      Size            =   "4895;1296"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   555
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblMinhHoa 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   3735
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Chuong Trinh Minh Hoa"
      Size            =   "6588;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "..................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   255
      Width           =   2475
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   ".........................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   12
      Height          =   975
      Index           =   1
      Left            =   360
      Top             =   360
      Width           =   975
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "fAbout.frx":08CA
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5859.683
      Y1              =   1987.827
      Y2              =   1987.827
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   5640
      Top             =   120
      Width           =   555
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H000080FF&
      BorderWidth     =   12
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   5565
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0080C0FF&
      Height          =   345
      Index           =   4
      Left            =   1440
      TabIndex        =   2
      Top             =   660
      Width           =   4845
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1440
      TabIndex        =   9
      Top             =   960
      Width           =   4845
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
   Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "Ve Tac Gia"
    lblMinhHoa.Caption = ChrW(67) & ChrW(104) & ChrW(432) & ChrW(417) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(84) & ChrW(114) & ChrW(236) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(77) & ChrW(105) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(72) & ChrW(7885) & ChrW(97)
    lblHeThong.Caption = ChrW(72) & ChrW(7879) & ChrW(32) & ChrW(84) & ChrW(104) & ChrW(7889) & ChrW(110) & ChrW(103)
    lblBatLoi.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(76) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(67) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(84) & ChrW(7843)
    lblChinhTa.Caption = ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116)
    lblTacGia.Caption = ChrW(67) & ChrW(104) & ChrW(7911) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(109) & ChrW(32) & ChrW(273) & ChrW(7873) & ChrW(32) & ChrW(116) & ChrW(224) & ChrW(105) & ChrW(58) & ChrW(32) & ChrW(80) & ChrW(71) & ChrW(83) & ChrW(46) & ChrW(84) & ChrW(83) & ChrW(46) & ChrW(32) & ChrW(78) & ChrW(103) & ChrW(117) & ChrW(121) & ChrW(7877) & ChrW(110) & ChrW(32) & ChrW(71) & ChrW(105) & ChrW(97) & ChrW(32) & ChrW(272) & ChrW(7883) & ChrW(110) & ChrW(104)
    lblLop.Caption = ChrW(75) & ChrW(104) & ChrW(111) & ChrW(97) & ChrW(32) & ChrW(84) & ChrW(111) & ChrW(225) & ChrW(110)
    lblTruong.Caption = ChrW(84) & ChrW(114) & ChrW(432) & ChrW(7901) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(272) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(72) & ChrW(7885) & ChrW(99) & ChrW(32) & ChrW(75) & ChrW(104) & ChrW(111) & ChrW(97) & ChrW(32) & ChrW(72) & ChrW(7885) & ChrW(99) & ChrW(32) & ChrW(45) & ChrW(32) & ChrW(72) & ChrW(117) & ChrW(7871)
    lblDiaChi.Caption = ChrW(75) & ChrW(105) & ChrW(7875) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(97) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(120) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(99) & ChrW(7911) & ChrW(97) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116)
    Label4.Caption = ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(417) & ChrW(32) & ChrW(115) & ChrW(7903) & ChrW(32) & ChrW(273) & ChrW(7897) & ChrW(32) & ChrW(108) & ChrW(105) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7863) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7869)
End Sub
