VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHocTu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoc Tu"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHocTu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   9975
      Begin MSForms.CommandButton cmdBoChon 
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
         Caption         =   "Chon"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdHocTu 
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
         Caption         =   "Hoc Tu"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdChon 
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   600
         Width           =   1335
         Caption         =   "Chon"
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
         TabIndex        =   13
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   0
         Width           =   2535
         Caption         =   "Nhung tu phan tach duoc"
         Size            =   "4471;450"
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
         TabIndex        =   16
         Top             =   0
         Width           =   1815
         Caption         =   "Nhung tu da chon"
         Size            =   "3201;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9975
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
         Height          =   2895
         Left            =   120
         TabIndex        =   6
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
      Begin MSForms.CommandButton cmdDong 
         Height          =   375
         Left            =   8520
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
      Begin MSForms.CommandButton cmdTachTu 
         Height          =   375
         Left            =   7080
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   0
         Width           =   5535
         Caption         =   "Nhap van ban can tach tu roi nhan nut Tach Tu ben duoi"
         Size            =   "9763;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblHelp 
         Height          =   2295
         Left            =   7200
         TabIndex        =   2
         Top             =   360
         Width           =   2655
         Caption         =   "Hien tai, chuong trinh phan biet cac tu qua bon ky tu:"
         Size            =   "4683;4048"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         X1              =   6960
         X2              =   6960
         Y1              =   360
         Y2              =   3240
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
      Left            =   4522
      TabIndex        =   11
      Top             =   360
      Width           =   1170
      ForeColor       =   8421631
      Caption         =   "HOC TU"
      Size            =   "2064;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmHocTu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sen As cpsSentences
Dim bsave As Boolean

Private Sub cmdBoChon_Click()
Dim i As Long
    For i = 0 To lstHoc.ListCount - 1
        If lstHoc.Selected(i) Then
            lstCau.AddItem lstHoc.List(i)
            lstHoc.RemoveItem i
            Exit For
        End If
    Next i
End Sub

Private Sub cmdHocTu_Click()
Dim i As Long, e As Integer
Dim vWord As New clsWord
Dim st As String, Pos As Long
    DoEvents
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(103) & ChrW(104) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(118) & ChrW(224) & ChrW(111) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
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

Private Sub cmdChon_Click()
Dim i As Long
    For i = 0 To lstCau.ListCount - 1
        If lstCau.Selected(i) Then
            lstHoc.AddItem lstCau.List(i)
            lstCau.RemoveItem i
            Exit For
        End If
    Next i
End Sub

Private Sub cmdDong_Click()
    Unload frmHocTu
End Sub

Private Sub cmdTachtu_Click()
Dim i As Long, sotu As Long, j As Long, k As Long
Dim e As Integer, tu As String
    DoEvents
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
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
            e = sen.WordSplit(i)
            If e <> 0 Then
                Call ErrorHandle(e)
                Exit Sub
            End If
        Next i
        sotu = 0
        For i = 1 To sen.Length
            For j = 1 To sen.Sentence(i).WCount
                For k = 2 To 3
                    If sen.Sentence(i).AW(k, j).Length <> 0 Then
                        sotu = sotu + 1
                        txtKiemTra.SelStart = sen.Sentence(i).AW(k, j).Start
                        txtKiemTra.SelLength = sen.Sentence(i).AW(k, j).Length
                        tu = txtKiemTra.SelText
                        lstCau.AddItem sotu & " " & tu
                    End If
                Next k
            Next j
        Next i
        lblTip1.Caption = ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(99) & ChrW(226) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sen.Length & ChrW(46) & ChrW(32) & ChrW(83) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(58) & sotu & "."
    End If
    Unload frmFlash
End Sub

Private Sub Form_Load()
    bsave = False
    Call AddCaption
End Sub

Private Sub AddCaption()
    Me.Caption = "Hoc Tu - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(72) & ChrW(7884) & ChrW(67) & ChrW(32) & ChrW(84) & ChrW(7914)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(118) & ChrW(259) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(110) & ChrW(250) & ChrW(116) & ChrW(32) & ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(84) & ChrW(7915) & ChrW(32) & ChrW(98) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(100) & ChrW(432) & ChrW(7899) & ChrW(105)
    cmdTachTu.Caption = ChrW(84) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(84) & ChrW(7915)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    lblFrame2.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(226) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99)
    lblFrame3.Caption = ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7885) & ChrW(110)
    cmdChon.Caption = ChrW(67) & ChrW(104) & ChrW(7885) & ChrW(110)
    cmdHocTu.Caption = ChrW(72) & ChrW(7885) & ChrW(99) & ChrW(32) & ChrW(84) & ChrW(7915)
    cmdBoChon.Caption = ChrW(66) & ChrW(7887) & ChrW(32) & ChrW(67) & ChrW(104) & ChrW(7885) & ChrW(110)
    lblHelp.Caption = ChrW(83) & ChrW(97) & ChrW(117) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(105) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(97) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(97) & ChrW(105) & ChrW(32) & ChrW(273) & ChrW(111) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(225) & ChrW(99) & ChrW(104) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(44) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(103) & ChrW(7897) & ChrW(112) & ChrW(32) & ChrW(108) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(109) & ChrW(7897) & ChrW(116) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & _
        ChrW(116) & ChrW(44) & ChrW(32) & ChrW(104) & ChrW(97) & ChrW(105) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(118) & ChrW(224) & ChrW(32) & ChrW(98) & ChrW(97) & ChrW(32) & ChrW(226) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(7875) & ChrW(32) & ChrW(107) & ChrW(105) & ChrW(7875) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(97) & ChrW(32) & ChrW(120) & ChrW(101) & ChrW(109) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(46) & ChrW(32) & ChrW(78) & ChrW(104) & ChrW(7919) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & _
        ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(108) & ChrW(432) & ChrW(117) & ChrW(32) & ChrW(108) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(273) & ChrW(7875) & ChrW(32) & ChrW(120) & ChrW(7917) & ChrW(32) & ChrW(108) & ChrW(253) & ChrW(32) & ChrW(115) & ChrW(97) & ChrW(117) & ChrW(32) & ChrW(110) & ChrW(224) & ChrW(121) & ChrW(46)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim e As Integer
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(83) & ChrW(97) & ChrW(118) & ChrW(101) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    If bsave Then
        e = Dic.SaveDic
        If e <> 0 Then Call ErrorHandle(e)
    End If
    Unload frmFlash
End Sub
