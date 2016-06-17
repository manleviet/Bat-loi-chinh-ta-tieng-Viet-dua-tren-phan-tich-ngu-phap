VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTachDoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mo Phong Giai Doan Tach Cau"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTachDoan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   11415
      Begin MSForms.Label lblTip1 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3480
         Width           =   5055
         Size            =   "8916;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Line Line1 
         X1              =   8280
         X2              =   8280
         Y1              =   360
         Y2              =   3240
      End
      Begin MSForms.Label lblHelp 
         Height          =   1095
         Left            =   8520
         TabIndex        =   11
         Top             =   1080
         Width           =   2655
         Caption         =   "Hien tai, chuong trinh xem mot cau duoc ket thuc voi cac ky hieu:"
         Size            =   "4683;1931"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame1 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   6375
         Caption         =   "Nhap van ban can thu tach doan roi nhan nut Tach Doan ben duoi"
         Size            =   "11245;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdTachCau 
         Height          =   375
         Left            =   8520
         TabIndex        =   6
         Top             =   3360
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
         Left            =   9960
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
         Caption         =   "Dong"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox txtKiemTra 
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7935
         VariousPropertyBits=   -1395636197
         BorderStyle     =   1
         ScrollBars      =   3
         Size            =   "13996;5106"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblTip 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4920
         Width           =   5055
         Size            =   "8916;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   11415
      Begin MSForms.ListBox lstCau 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11175
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "19711;5530"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame2 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   2775
         Caption         =   "Nhung doan phan tach duoc"
         Size            =   "4895;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin MSForms.Label lblHeader 
      Height          =   345
      Left            =   3188
      TabIndex        =   1
      Top             =   360
      Width           =   5235
      ForeColor       =   8421631
      Caption         =   "MO PHONG GIAI DOAN TACH DOAN"
      Size            =   "9234;609"
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
Attribute VB_Name = "frmTachDoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: frmTachCau Form - a form demonstrating sentence spliting
'----------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Dim sen As cpsChecking

Private Sub cmdDong_Click()
    Unload frmTachDoan
End Sub

Private Sub cmdTachCau_Click()
Dim i As Long
Dim e As Integer
    lstCau.Clear
    Set sen = New cpsChecking
    sen.Init txtKiemTra
    e = sen.ParagraphSplit
    If e <> 0 Then
        Call ErrorHandle(e)
    Else
        lblTip1.Caption = ChrW(272) & ChrW(227) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & " " & sen.ParagraphCount & ChrW(32) & ChrW(273) & ChrW(111) & ChrW(7841) & ChrW(110)
        For i = 1 To sen.ParagraphCount
            txtKiemTra.SelStart = sen.Paragraph(i).ParagraphStart
            txtKiemTra.SelLength = sen.Paragraph(i).ParagraphLength
            lstCau.AddItem i & " " & txtKiemTra.SelText
        Next i
    End If
End Sub

Private Sub Form_Load()
    Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "Giai Doan Tach Doan - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(77) & ChrW(212) & ChrW(32) & ChrW(80) & ChrW(72) & ChrW(7886) & ChrW(78) & ChrW(71) & ChrW(32) & ChrW(71) & ChrW(73) & ChrW(65) & ChrW(73) & ChrW(32) & ChrW(272) & ChrW(79) & ChrW(7840) & ChrW(78) & ChrW(32) & ChrW(84) & ChrW(193) & ChrW(67) & ChrW(72) & ChrW(32) & ChrW(272) & ChrW(79) & ChrW(7840) & ChrW(78)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7917) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(111) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(250) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(272) & ChrW(111) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
    cmdTachCau.Caption = ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(272) & ChrW(111) & ChrW(7841) & ChrW(110)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    lblFrame2.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(273) & ChrW(111) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblHelp.Caption = ChrW(77) & ChrW(7897) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(111) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(250) & ChrW(99) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(105) & ChrW(32) & ChrW(120) & ChrW(117) & ChrW(7845) & ChrW(116) & ChrW(32) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(253) & ChrW(32) & ChrW(116) & ChrW(7921) & ChrW(32) & ChrW(120) & ChrW(117) & ChrW(7889) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(100) & ChrW(242) & ChrW(110) & ChrW(103)
End Sub
