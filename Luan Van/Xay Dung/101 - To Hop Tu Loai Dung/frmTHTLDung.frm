VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTHTLDung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mo Phong Giai Doan Gop Tu"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTHTLDung.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11415
      Begin MSForms.Label lblTip 
         Height          =   255
         Left            =   120
         TabIndex        =   7
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
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7815
         VariousPropertyBits=   -1395636197
         BorderStyle     =   1
         ScrollBars      =   3
         Size            =   "13785;1931"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdDong 
         Height          =   375
         Left            =   9840
         TabIndex        =   5
         Top             =   1560
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
         Left            =   8400
         TabIndex        =   4
         Top             =   1560
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
         TabIndex        =   3
         Top             =   0
         Width           =   4095
         Caption         =   "Nhap mot cau vao day roi nhan Phan Tich"
         Size            =   "7223;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp 
         Height          =   1095
         Left            =   8400
         TabIndex        =   2
         Top             =   360
         Width           =   2775
         Caption         =   "Chi nen nhap mot cau"
         Size            =   "4895;1931"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   8160
         X2              =   8160
         Y1              =   360
         Y2              =   1440
      End
      Begin MSForms.Label lblTip1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1560
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
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   11415
      Begin MSForms.ListBox lstCau 
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11175
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "19711;2143"
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
         TabIndex        =   10
         Top             =   0
         Width           =   2055
         Caption         =   "Nhung cau gop duoc"
         Size            =   "3625;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   11415
      Begin MSForms.ListBox lstToHop 
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   11175
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "19711;1931"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame3 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   2895
         Caption         =   "Nhung to hop tu loai co duoc"
         Size            =   "5106;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   11415
      Begin MSForms.ListBox lstTLDung 
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11175
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "19711;1931"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame4 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   1815
         Caption         =   "Chuoi tu loai dung"
         Size            =   "3201;450"
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
      TabIndex        =   12
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
      Left            =   4095
      TabIndex        =   11
      Top             =   360
      Width           =   3435
      ForeColor       =   8421631
      Caption         =   "TO HOP TU LOAI DUNG"
      Size            =   "6059;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmTHTLDung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: frmGopTu Form - a form demonstrating word add up sphase
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Dim sen As cpsChecking

Private Sub cmdDong_Click()
    Unload frmTHTLDung
End Sub

Private Sub cmdTachtu_Click()
Dim i As Long, j As Long, k As Long, x As Long, y As Long, t As Long
Dim e As Integer, tu As String, sotohop As Long, sotohop2 As Long
    lstCau.Clear
    lstToHop.Clear
    lstTLDung.Clear
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
                e = sen.AWSplit(i, j)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    Exit Sub
                End If
                e = sen.MakeRS(i, j)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    Exit Sub
                End If
                e = sen.MakeWCS(i, j)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    Exit Sub
                End If
                e = sen.EarlyParse(i, j, RDic)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    Exit Sub
                End If
            Next j
        Next i
        sotohop = 0
        For i = 1 To sen.ParagraphCount
            For j = 1 To sen.SentenceCount(i)
                sotohop = sotohop + sen.RSCount(i, j)
                For k = 1 To sen.RSCount(i, j)
                    tu = ""
                    For t = 1 To sen.RSSCount(i, j, k)
                        x = sen.RS(i, j, k, t).x
                        y = sen.RS(i, j, k, t).y
                        txtKiemTra.SelStart = sen.AW(i, j, x, y).Start
                        txtKiemTra.SelLength = sen.AW(i, j, x, y).Length
                        tu = tu & txtKiemTra.SelText
                        If k <> sen.RSSCount(i, j, k) Then
                            tu = tu & " | "
                        End If
                    Next t
                    lstCau.AddItem i & " " & j & " " & k & " " & tu
                Next k
            Next j
        Next i
        sotohop2 = 0
        For i = 1 To sen.ParagraphCount
            For j = 1 To sen.SentenceCount(i)
                sotohop2 = sotohop2 + sen.WCSCount(i, j)
                For t = 1 To sen.WCSCount(i, j)
                    lstToHop.AddItem i & " " & sen.WCSRSItem(i, j, t) & " " & j & " " & t & " " & sen.WCS(i, j, t)
                Next t
            Next j
        Next i
        lblTip1.Caption = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sen.ParagraphCount & ChrW(46) & ChrW(32) & ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7893) & ChrW(32) & ChrW(104) & ChrW(7907) & ChrW(112) & ChrW(32) & ChrW(103) & ChrW(7897) & ChrW(112) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sotohop & ChrW(46) & ChrW(32) & ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7893) & ChrW(32) & ChrW(104) & ChrW(7907) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sotohop2 & "."
        For i = 1 To sen.ParagraphCount
            For j = 1 To sen.SentenceCount(i)
                lstTLDung.AddItem sen.WCS(i, j, sen.EarlyRight(i, j))
            Next j
        Next i
    End If
End Sub

Private Sub Form_Load()
    Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "To Hop Tu Loai Dung - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(84) & ChrW(7892) & ChrW(32) & ChrW(72) & ChrW(7906) & ChrW(80) & ChrW(32) & ChrW(84) & ChrW(7914) & ChrW(32) & ChrW(76) & ChrW(79) & ChrW(7840) & ChrW(73) & ChrW(32) & ChrW(272) & ChrW(218) & ChrW(78) & ChrW(71)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(109) & ChrW(7897) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(118) & ChrW(224) & ChrW(111) & ChrW(32) & ChrW(273) & ChrW(226) & ChrW(121) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(80) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(84) & ChrW(237) & ChrW(99) & ChrW(104)
    cmdTachTu.Caption = ChrW(80) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(84) & ChrW(237) & ChrW(99) & ChrW(104)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    lblFrame3.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7893) & ChrW(32) & ChrW(104) & ChrW(7907) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblFrame2.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(103) & ChrW(7897) & ChrW(112) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblFrame4.Caption = ChrW(67) & ChrW(104) & ChrW(117) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(273) & ChrW(250) & ChrW(110) & ChrW(103)
    lblHelp.Caption = ChrW(67) & ChrW(104) & ChrW(7881) & ChrW(32) & ChrW(110) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(109) & ChrW(7897) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(46)
End Sub
