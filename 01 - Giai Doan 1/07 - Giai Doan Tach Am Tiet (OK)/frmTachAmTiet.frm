VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTachAmTiet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mo Phong Giai Doan Tach Tu"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTachAmTiet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   12
      Top             =   4560
      Width           =   11415
      Begin MSForms.ListBox lstCau 
         Height          =   3135
         Left            =   120
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   0
         Width           =   3015
         Caption         =   "Nhung am tiet phan tach duoc"
         Size            =   "5318;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
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
      TabIndex        =   0
      Top             =   720
      Width           =   11415
      Begin MSForms.Label lblTip 
         Height          =   255
         Left            =   120
         TabIndex        =   11
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
      Begin MSForms.TextBox txtKiemTra 
         Height          =   2895
         Left            =   120
         TabIndex        =   10
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
      Begin MSForms.CommandButton cmdDong 
         Height          =   375
         Left            =   9960
         TabIndex        =   9
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
      Begin MSForms.CommandButton cmdTachTu 
         Height          =   375
         Left            =   8520
         TabIndex        =   8
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
      Begin MSForms.Label lblFrame1 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   6495
         Caption         =   "Nhap van ban can tach am tiet roi nhan nut Tach Am Tiet ben duoi"
         Size            =   "11456;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp 
         Height          =   855
         Left            =   8520
         TabIndex        =   6
         Top             =   360
         Width           =   2655
         Caption         =   "Hien tai, chuong trinh phan biet cac tu qua bon ky tu:"
         Size            =   "4683;1508"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   8280
         X2              =   8280
         Y1              =   360
         Y2              =   3240
      End
      Begin MSForms.Label lblHelp1 
         Height          =   255
         Left            =   8880
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         Caption         =   "- Khoang trong (' ')"
         Size            =   "2990;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp2 
         Height          =   255
         Left            =   8880
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
         Caption         =   "- Ky tu Tab"
         Size            =   "2990;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp3 
         Height          =   255
         Left            =   8880
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
         Caption         =   "- Dau phay (,)"
         Size            =   "3413;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp4 
         Height          =   255
         Left            =   8880
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
         Caption         =   "- Dau cham phay (;)"
         Size            =   "3201;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblTip1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
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
   End
   Begin MSForms.Label lblTitle 
      Height          =   240
      Left            =   120
      TabIndex        =   16
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
   Begin MSForms.Label lblHeader 
      Height          =   345
      Left            =   3022
      TabIndex        =   15
      Top             =   360
      Width           =   5610
      ForeColor       =   8421631
      Caption         =   "MO PHONG GIAI DOAN TACH AM TIET"
      Size            =   "9895;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmTachAmTiet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietmanese Checking
'Description: frmTachAmTiet Form - a form demonstrating syllable split sphase
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Dim sen As cpsChecking

Private Sub cmdDong_Click()
    Unload frmTachAmTiet
End Sub

Private Sub cmdTachtu_Click()
Dim i As Long, sotu As Long, j As Long, k As Long
Dim e As Integer, tu As String, socau As Long
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
        For i = 1 To sen.ParagraphCount
            For j = 1 To sen.SentenceCount(i)
                e = sen.SyllSplit(i, j)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    Exit Sub
                End If
            Next j
        Next i
        sotu = 0
        socau = 0
        For i = 1 To sen.ParagraphCount
            socau = socau + sen.SentenceCount(i)
            For j = 1 To sen.SentenceCount(i)
                sotu = sotu + sen.SyllCount(i, j)
                tu = ""
                For k = 1 To sen.SyllCount(i, j)
                    txtKiemTra.SelStart = sen.Syllable(i, j, k).Start
                    txtKiemTra.SelLength = sen.Syllable(i, j, k).Length
                    tu = tu & txtKiemTra.SelText
                    If k < sen.SyllCount(i, j) Then tu = tu & " | "
                Next k
                lstCau.AddItem i & " " & j & " " & tu
            Next j
        Next i
        lblTip1.Caption = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & socau & ChrW(46) & ChrW(32) & ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sotu & "."
    End If
End Sub

Private Sub Form_Load()
    Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "Giai Doan Tach Am Tiet - Phien Ban " & App.Major & "." & App.Minor
    lblHeader.Caption = ChrW(77) & ChrW(212) & ChrW(32) & ChrW(80) & ChrW(72) & ChrW(7886) & ChrW(78) & ChrW(71) & ChrW(32) & ChrW(71) & ChrW(73) & ChrW(65) & ChrW(73) & ChrW(32) & ChrW(272) & ChrW(79) & ChrW(7840) & ChrW(78) & ChrW(32) & ChrW(84) & ChrW(193) & ChrW(67) & ChrW(72) & ChrW(32) & ChrW(194) & ChrW(77) & ChrW(32) & ChrW(84) & ChrW(73) & ChrW(7870) & ChrW(84)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(250) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(194) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(98) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
    cmdTachTu.Caption = ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(194) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(116)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    lblFrame2.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblHelp.Caption = ChrW(72) & ChrW(105) & ChrW(7879) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(7841) & ChrW(105) & ChrW(44) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(432) & ChrW(417) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(236) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(97) & ChrW(32) & ChrW(98) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(253) & ChrW(32) & ChrW(116) & ChrW(7921) & ChrW(32) & ChrW(115) & ChrW(97) & ChrW(117) & ChrW(58)
    lblHelp1.Caption = ChrW(45) & ChrW(32) & ChrW(75) & ChrW(104) & ChrW(111) & ChrW(7843) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(7889) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(40) & ChrW(39) & ChrW(32) & ChrW(39) & ChrW(41)
    lblHelp2.Caption = ChrW(45) & ChrW(32) & ChrW(75) & ChrW(253) & ChrW(32) & ChrW(116) & ChrW(7921) & ChrW(32) & ChrW(84) & ChrW(97) & ChrW(98)
    lblHelp3.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(7849) & ChrW(121) & ChrW(32) & ChrW(40) & ChrW(44) & ChrW(41)
    lblHelp4.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7845) & ChrW(109) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(7849) & ChrW(121) & ChrW(32) & ChrW(40) & ChrW(59) & ChrW(41)
End Sub
