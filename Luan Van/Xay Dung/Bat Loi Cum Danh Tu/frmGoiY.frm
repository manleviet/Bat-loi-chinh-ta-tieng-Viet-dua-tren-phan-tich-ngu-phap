VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmGoiY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goi Y"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGoiY.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6615
      Begin MSForms.Label lblFrame1 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   975
         Caption         =   "Ngu Canh"
         Size            =   "1720;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdBoqua 
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   600
         Width           =   1455
         Caption         =   "Bo Qua"
         Size            =   "2566;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdBoQuaHet 
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
         Caption         =   "Bo Qua Het"
         Size            =   "2566;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox txtNguCanh 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4575
         VariousPropertyBits=   -1400879077
         Size            =   "8070;1931"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAmTiet 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   4575
         VariousPropertyBits=   746604571
         Size            =   "8070;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   6615
      Begin MSForms.Label lblFrame2 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1455
         Caption         =   "De Nghi Sua Loi"
         Size            =   "2566;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdThayThe 
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   600
         Width           =   1455
         Caption         =   "Thay The"
         Size            =   "2566;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdThayTheHet 
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
         Caption         =   "Thay The Het"
         Size            =   "2566;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDung 
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
         Caption         =   "Dung"
         Size            =   "2566;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ListBox lstTuongDuong 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4575
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "8070;3413"
         MatchEntry      =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.Label lblTitle 
      Height          =   240
      Left            =   120
      TabIndex        =   1
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
      Left            =   2947
      TabIndex        =   0
      Top             =   360
      Width           =   930
      ForeColor       =   8421631
      Caption         =   "GOI Y"
      Size            =   "1640;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmGoiY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sResult As eGoiY

Public Property Get Result() As eGoiY
    Result = sResult
End Property

Private Sub cmdBoqua_Click()
    sResult = BoQua
    Me.Hide
End Sub

Private Sub cmdBoQuaHet_Click()
    sResult = BoQuaHet
    Me.Hide
End Sub

Private Sub cmdDung_Click()
    sResult = Dung
    Unload Me
End Sub

Private Sub cmdThayThe_Click()
    sResult = ThayThe
    Me.Hide
End Sub

Private Sub cmdThayTheHet_Click()
    sResult = ThayTheHet
    Me.Hide
End Sub

Private Sub Form_Load()
    Call AddCaption
    lstTuongDuong.Clear
End Sub

Private Sub lstTuongDuong_Click()
    txtAmTiet.Text = lstTuongDuong.Text
End Sub

Public Function SimilarSyllables() As Integer
On Error GoTo Result
Dim i As Long
Dim st As String
    SimilarSyllables = 0
    For i = 1 To Dic.PCount
        st = Dic.WordPage(i)
        If SimilarSyllable(UniLCase(txtAmTiet.Text), Len(txtAmTiet.Text), st, Len(st)) Then
            lstTuongDuong.AddItem st
        End If
    Next i
    Exit Function
Result:
    SimilarSyllables = 1000
End Function
'ham kiem tra hai tu co tuong duong hay khong
Private Function SimilarSyllable(ByVal Word As String, ByVal ILen As Long, ByVal TWord As String, ByVal TLen As Long) As Boolean
Dim L1 As Long, L2 As Long, W1 As String, W2 As String
Dim i As Long, MatchCount As Long, Pos As Long
    If ILen > TLen Then
        If ILen - TLen > 2 Then
            SimilarSyllable = False
            Exit Function
        End If
        L1 = ILen
        L2 = TLen
        W1 = Word
        W2 = TWord
    Else
        If TLen - ILen > 2 Then
            SimilarSyllable = False
            Exit Function
        End If
        L1 = TLen
        L2 = ILen
        W1 = TWord
        W2 = Word
    End If
    MatchCount = 0
    If ILen >= 4 Then
        For i = 1 To L2
            Pos = InStr(W1, Mid(W2, i, 1))
            If Pos > 0 And (Pos >= i - 2 Or Pos <= i + 2) Then
                MatchCount = MatchCount + 1
                W1 = Left(W1, Pos - 1) & Mid(W1, Pos + 1)
            End If
        Next
        If MatchCount >= (L1 - 1) Then
            SimilarSyllable = True
        Else
            SimilarSyllable = False
        End If
    Else
        For i = 1 To L2
            Pos = InStr(W1, Mid(W2, i, 1))
            If Pos > 0 And (Pos >= i - 1 Or Pos <= i + 1) Then
                MatchCount = MatchCount + 1
                W1 = Left(W1, Pos - 1) & Mid(W1, Pos + 1)
            End If
        Next
        If MatchCount >= (L1 - 1) Then
            SimilarSyllable = True
        Else
            SimilarSyllable = False
        End If
    End If
End Function

Private Sub AddCaption()
    Me.Caption = "Goi Y Sua Loi - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
    lblHeader.Caption = ChrW(71) & ChrW(7906) & ChrW(73) & ChrW(32) & ChrW(221)
    lblTitle.Caption = ChrW(78) & ChrW(103) & ChrW(104) & ChrW(105) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7913) & ChrW(117) & ChrW(32) & ChrW(118) & ChrW(224) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(225) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(432) & ChrW(417) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(225) & ChrW(112) & ChrW(32) & ChrW(98) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(84) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116)
    lblFrame1.Caption = ChrW(78) & ChrW(103) & ChrW(7919) & ChrW(32) & ChrW(67) & ChrW(7843) & ChrW(110) & ChrW(104)
    lblFrame2.Caption = ChrW(272) & ChrW(7873) & ChrW(32) & ChrW(78) & ChrW(103) & ChrW(104) & ChrW(7883) & ChrW(32) & ChrW(83) & ChrW(7917) & ChrW(97) & ChrW(32) & ChrW(76) & ChrW(7895) & ChrW(105)
    cmdBoqua.Caption = ChrW(66) & ChrW(7887) & ChrW(32) & ChrW(81) & ChrW(117) & ChrW(97)
    cmdBoQuaHet.Caption = ChrW(66) & ChrW(7887) & ChrW(32) & ChrW(81) & ChrW(117) & ChrW(97) & ChrW(32) & ChrW(72) & ChrW(7871) & ChrW(116)
    cmdThayThe.Caption = ChrW(84) & ChrW(104) & ChrW(97) & ChrW(121) & ChrW(32) & ChrW(84) & ChrW(104) & ChrW(7871)
    cmdThayTheHet.Caption = ChrW(84) & ChrW(104) & ChrW(97) & ChrW(121) & ChrW(32) & ChrW(84) & ChrW(104) & ChrW(7871) & ChrW(32) & ChrW(72) & ChrW(7871) & ChrW(116)
    cmdDung.Caption = ChrW(68) & ChrW(7915) & ChrW(110) & ChrW(103)
End Sub

