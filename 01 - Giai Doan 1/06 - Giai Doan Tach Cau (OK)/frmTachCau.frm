VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTachCau 
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
   Icon            =   "frmTachCau.frx":0000
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
         TabIndex        =   16
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
      Begin MSForms.Label lblHelp4 
         Height          =   255
         Left            =   9000
         TabIndex        =   15
         Top             =   2040
         Width           =   1815
         Caption         =   "- Ky tu xuong dong"
         Size            =   "3201;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp3 
         Height          =   255
         Left            =   9000
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
         Caption         =   "- Dau cham than (!)"
         Size            =   "3413;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp2 
         Height          =   255
         Left            =   9000
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
         Caption         =   "- Dau cham hoi (?)"
         Size            =   "2990;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp1 
         Height          =   255
         Left            =   9000
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
         Caption         =   "- Dau cham (.)"
         Size            =   "2355;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
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
         Height          =   735
         Left            =   8520
         TabIndex        =   11
         Top             =   480
         Width           =   2655
         Caption         =   "Hien tai, chuong trinh xem mot cau duoc ket thuc voi cac ky hieu:"
         Size            =   "4683;1296"
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
         Width           =   6135
         Caption         =   "Nhap van ban can thu tach cau roi nhan nut Tach Cau ben duoi"
         Size            =   "10821;450"
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
         Width           =   2655
         Caption         =   "Nhung cau phan tach duoc"
         Size            =   "4683;450"
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
      Left            =   3300
      TabIndex        =   1
      Top             =   360
      Width           =   5010
      ForeColor       =   8421631
      Caption         =   "MO PHONG GIAI DOAN TACH CAU"
      Size            =   "8837;609"
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
Attribute VB_Name = "frmTachCau"
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
    Unload frmTachCau
End Sub

Private Sub cmdTachCau_Click()
Dim i As Long, j As Long
Dim socau As Long
Dim e As Integer
    lstCau.Clear
    Set sen = New cpsChecking
    sen.Init txtKiemTra
    e = sen.ParagraphSplit
    If e <> 0 Then
        Call ErrorHandle(e)
    Else
        For i = 1 To sen.ParagraphCount
            e = sen.SentenceSplit(i)
            If e <> 0 Then
                Call ErrorHandle(e)
                Exit Sub
            End If
        Next i
        socau = 0
        
        For i = 1 To sen.ParagraphCount
            socau = socau + sen.SentenceCount(i)
            For j = 1 To sen.SentenceCount(i)
                txtKiemTra.SelStart = sen.Sentence(i, j).SentenceStart
                txtKiemTra.SelLength = sen.Sentence(i, j).SentenceLength
                lstCau.AddItem i & " " & j & " " & txtKiemTra.SelText
            Next j
        Next i
        lblTip1.Caption = ChrW(272) & ChrW(227) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & " " & socau & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117)
    End If
End Sub

Private Sub Form_Load()
    Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "Giai Doan Tach Cau - Phien Ban " & App.Major & "." & App.Minor
    lblHeader.Caption = ChrW(77) & ChrW(212) & ChrW(32) & ChrW(80) & ChrW(72) & ChrW(7886) & ChrW(78) & ChrW(71) & ChrW(32) & ChrW(71) & ChrW(73) & ChrW(65) & ChrW(73) & ChrW(32) & ChrW(272) & ChrW(79) & ChrW(7840) & ChrW(78) & ChrW(32) & ChrW(84) & ChrW(193) & ChrW(67) & ChrW(72) & ChrW(32) & ChrW(67) & ChrW(194) & ChrW(85)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7917) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(250) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(67) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(98) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
    cmdTachCau.Caption = ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(67) & ChrW(226) & ChrW(117)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    lblFrame2.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblHelp.Caption = ChrW(72) & ChrW(105) & ChrW(7879) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(7841) & ChrW(105) & ChrW(44) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(432) & ChrW(417) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(236) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(120) & ChrW(101) & ChrW(109) & ChrW(32) & ChrW(109) & ChrW(7897) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(107) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(250) & ChrW(99) & ChrW(32) & ChrW(118) & ChrW(7899) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(107) & ChrW(253) & ChrW(32) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(117) & ChrW(32) & ChrW(115) & ChrW(97) & ChrW(117) & ChrW(58)
    lblHelp1.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7845) & ChrW(109) & ChrW(32) & ChrW(40) & ChrW(46) & ChrW(41)
    lblHelp2.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7845) & ChrW(109) & ChrW(32) & ChrW(104) & ChrW(7887) & ChrW(105) & ChrW(32) & ChrW(40) & ChrW(63) & ChrW(41)
    lblHelp3.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7845) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(97) & ChrW(110) & ChrW(32) & ChrW(40) & ChrW(33) & ChrW(41)
    lblHelp4.Caption = ChrW(45) & ChrW(32) & ChrW(75) & ChrW(253) & ChrW(32) & ChrW(116) & ChrW(7921) & ChrW(32) & ChrW(120) & ChrW(117) & ChrW(7889) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(100) & ChrW(242) & ChrW(110) & ChrW(103)
End Sub
