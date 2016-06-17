VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHocAmTiet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoc Am Tiet"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHocAmTiet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10200
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
      TabIndex        =   3
      Top             =   720
      Width           =   9975
      Begin MSForms.Label lblTip1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   6615
         Size            =   "11668;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp4 
         Height          =   255
         Left            =   7560
         TabIndex        =   13
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
      Begin MSForms.Label lblHelp3 
         Height          =   255
         Left            =   7560
         TabIndex        =   12
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
      Begin MSForms.Label lblHelp2 
         Height          =   255
         Left            =   7560
         TabIndex        =   11
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
      Begin MSForms.Label lblHelp1 
         Height          =   255
         Left            =   7560
         TabIndex        =   10
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
      Begin VB.Line Line1 
         X1              =   6960
         X2              =   6960
         Y1              =   360
         Y2              =   3240
      End
      Begin MSForms.Label lblHelp 
         Height          =   855
         Left            =   7200
         TabIndex        =   9
         Top             =   360
         Width           =   2655
         Caption         =   "Hien tai, chuong trinh phan biet cac tu qua bon ky tu:"
         Size            =   "4683;1508"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame1 
         Height          =   255
         Left            =   120
         TabIndex        =   8
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
      Begin MSForms.CommandButton cmdTachTu 
         Height          =   375
         Left            =   7080
         TabIndex        =   7
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
         Left            =   8520
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   360
         Width           =   6615
         VariousPropertyBits=   -1395636197
         BorderStyle     =   1
         ScrollBars      =   3
         Size            =   "11668;5106"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblTip 
         Height          =   255
         Left            =   120
         TabIndex        =   4
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   9975
      Begin MSForms.CommandButton cmdHoc 
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
         Caption         =   "Hoc Am Tiet"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdLoc 
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   720
         Width           =   1335
         Caption         =   "Loc"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ListBox lstHoc 
         Height          =   1935
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   3975
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "7011;3413"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ListBox lstCau 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "7011;3413"
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
         TabIndex        =   2
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
      Begin MSForms.Label lblFrame3 
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   0
         Width           =   2655
         Caption         =   "Nhung am tiet da loc duoc"
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
      Left            =   4095
      TabIndex        =   16
      Top             =   360
      Width           =   2010
      ForeColor       =   8421631
      Caption         =   "HOC AM TIET"
      Size            =   "3545;609"
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
      TabIndex        =   15
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
Attribute VB_Name = "frmHocAmTiet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietnamese Checking

Option Explicit
Dim sen As cpsSentences
Dim bsave As Boolean

Private Sub cmdDong_Click()
    Unload frmHocAmTiet
End Sub

Private Sub cmdHoc_Click()
Dim i As Long, Pos As Long, st As String
Dim e As Integer
Dim vWord As New clsWord
    DoEvents
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(103) & ChrW(104) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(118) & ChrW(224) & ChrW(111) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    For i = 0 To lstHoc.ListCount - 1
        st = lstHoc.List(i)
        Pos = InStr(2, st, " ")
        st = Mid(st, Pos + 1)
        st = UniLCase(st)
        vWord.WordCont = st
        vWord.WordClass = "KN0|"
        e = Dic.AddWord(vWord)
        If e <> 0 Then
            Call ErrorHandle(e)
            Exit Sub
        End If
        Me.Refresh
        frmFlash.Refresh
    Next i
    bsave = True
    Unload frmFlash
End Sub

Private Sub cmdLoc_Click()
Dim i As Long, st As String, Pos As String
Dim loca As New clsLocation
    DoEvents
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(108) & ChrW(7885) & ChrW(99) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    
    lstHoc.Clear
    For i = 0 To lstCau.ListCount - 1
        st = lstCau.List(i)
        Pos = InStr(2, st, " ")
        st = Mid(st, Pos + 1)
        st = UniLCase(st)
        Set loca = Dic.FindSWord(st)
        If loca.ok <> 0 Then
            lstHoc.AddItem lstCau.List(i)
        End If
        Me.Refresh
        frmFlash.Refresh
    Next i
    
    Unload frmFlash
End Sub

Private Sub cmdTachtu_Click()
Dim i As Long, sotu As Long, j As Long
Dim e As Integer
    DoEvents
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh

    lstCau.Clear
    lstHoc.Clear
    Set sen = New cpsSentences
    sen.Init txtKiemTra
    e = sen.SentenceSplit
    If e <> 0 Then
        Call ErrorHandle(e)
    Else
        For i = 1 To sen.Length
            e = sen.SyllSplit(i)
            If e <> 0 Then
                Call ErrorHandle(e)
                Exit Sub
            End If
        Next i
        sotu = 0
        For i = 1 To sen.Length
            For j = 1 To sen.Sentence(i).WCount
                sotu = sotu + 1
                txtKiemTra.SelStart = sen.Sentence(i).Syllable(j).Start
                txtKiemTra.SelLength = sen.Sentence(i).Syllable(j).Length
                lstCau.AddItem sotu & " " & txtKiemTra.SelText
            Next j
        Next i
        lblTip1.Caption = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sen.Length & ChrW(46) & ChrW(32) & ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sotu & "."
    End If
    
    Unload frmFlash
End Sub

Private Sub Form_Load()
    bsave = False
    Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "Hoc Am Tiet - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(72) & ChrW(7884) & ChrW(67) & ChrW(32) & ChrW(194) & ChrW(77) & ChrW(32) & ChrW(84) & ChrW(73) & ChrW(7870) & ChrW(84)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(250) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(194) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(98) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
    cmdTachTu.Caption = ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(194) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(116)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    lblFrame2.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblFrame3.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(108) & ChrW(7885) & ChrW(99) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    cmdHoc.Caption = ChrW(72) & ChrW(7885) & ChrW(99) & ChrW(32) & ChrW(194) & ChrW(109) & ChrW(32) & ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(116)
    cmdLoc.Caption = ChrW(76) & ChrW(7885) & ChrW(99)
    lblHelp.Caption = ChrW(72) & ChrW(105) & ChrW(7879) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(7841) & ChrW(105) & ChrW(44) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(432) & ChrW(417) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(236) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(97) & ChrW(32) & ChrW(98) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(253) & ChrW(32) & ChrW(116) & ChrW(7921) & ChrW(32) & ChrW(115) & ChrW(97) & ChrW(117) & ChrW(58)
    lblHelp1.Caption = ChrW(45) & ChrW(32) & ChrW(75) & ChrW(104) & ChrW(111) & ChrW(7843) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(7889) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(40) & ChrW(39) & ChrW(32) & ChrW(39) & ChrW(41)
    lblHelp2.Caption = ChrW(45) & ChrW(32) & ChrW(75) & ChrW(253) & ChrW(32) & ChrW(116) & ChrW(7921) & ChrW(32) & ChrW(84) & ChrW(97) & ChrW(98)
    lblHelp3.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(7849) & ChrW(121) & ChrW(32) & ChrW(40) & ChrW(44) & ChrW(41)
    lblHelp4.Caption = ChrW(45) & ChrW(32) & ChrW(68) & ChrW(7845) & ChrW(117) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7845) & ChrW(109) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(7849) & ChrW(121) & ChrW(32) & ChrW(40) & ChrW(59) & ChrW(41)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim e As Integer
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(83) & ChrW(97) & ChrW(118) & ChrW(101) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    If bsave Then
        e = Dic.SaveDic
        If e <> 0 Then Call ErrorHandle(e)
    End If
    Unload frmFlash
End Sub
