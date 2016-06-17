VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmWManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Manage"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   5280
      TabIndex        =   16
      Top             =   720
      Width           =   4575
      Begin VB.CommandButton cmdALeft 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdARight 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   20
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   2760
         Width           =   375
      End
      Begin MSForms.ListBox lstWC 
         Height          =   2295
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   4335
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "7646;4048"
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
         TabIndex        =   18
         Top             =   0
         Width           =   735
         Caption         =   "Tu Loai"
         Size            =   "1296;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ListBox lstWCoWord 
         Height          =   2295
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   4335
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "7646;4048"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5055
      Begin MSForms.ListBox lstWord1 
         Height          =   3375
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "5953;5953"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblTongso 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   5280
         Width           =   3375
         Caption         =   "Tong so"
         Size            =   "5953;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ListBox lstType 
         Height          =   3015
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "5530;5318"
         MatchEntry      =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   4920
         Y1              =   1320
         Y2              =   1320
      End
      Begin MSForms.CommandButton cmdBoqua 
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         Caption         =   "Bo qua"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdLuu 
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   240
         Width           =   1335
         Caption         =   "L"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDong 
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   4320
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
         Left            =   3600
         TabIndex        =   11
         Top             =   3360
         Width           =   1335
         Caption         =   "Sua Doi"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdXoa 
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   3840
         Width           =   1335
         Caption         =   "Xoa"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdTaomoi 
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
         Caption         =   "Tao Moi"
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ListBox lstWord 
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3375
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "5953;5953"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblDSTu 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
         Caption         =   "Danh muc tu"
         Size            =   "2143;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblHelp 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   3375
         Caption         =   "Nhap tu can tim roi nhan Enter"
         Size            =   "5953;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtWord 
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   2895
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "5106;661"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFrame1 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   855
         Caption         =   "Tu Vung"
         Size            =   "1508;450"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblTu 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   495
         Caption         =   "Tim"
         Size            =   "873;450"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.CommandButton cmdReLoad 
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   360
      Width           =   1695
      Caption         =   "Cap nhat tu dien"
      Size            =   "2990;661"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblTitle 
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6495
      VariousPropertyBits=   276824091
      Caption         =   "Nghien cuu va phat trien phuong phap bat loi chinh ta"
      Size            =   "11456;423"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblHeader 
      Height          =   345
      Left            =   3622
      TabIndex        =   0
      Top             =   360
      Width           =   2730
      ForeColor       =   8421631
      Caption         =   "QUAN LY TU DIEN"
      Size            =   "4815;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmWManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: frmWManage Form - a form managing words
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Dim NoE As Integer ' 1 - New, 2 - Edit
Dim MouseDown As Boolean

Private Sub ChangeStatus(ByVal bState As Boolean)
    cmdLuu.Enabled = bState
    cmdBoqua.Enabled = bState
    cmdTaomoi.Enabled = Not bState
    cmdXoa.Enabled = Not bState
    cmdSuadoi.Enabled = Not bState
    lstWC.Visible = bState
    cmdRight.Visible = bState
    cmdARight.Visible = bState
    cmdLeft.Visible = bState
    cmdALeft.Visible = bState
    If Not bState Then
        lblHelp.Caption = ChrW(72) & ChrW(227) & ChrW(121) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(236) & ChrW(109) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(69) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114)
        lstWCoWord.Height = 5175
    Else
        lblHelp.Caption = ChrW(78) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(109) & ChrW(7899) & ChrW(105) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(69) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114)
        lstWCoWord.Height = 2295
    End If
End Sub

Private Sub cmdBoqua_Click()
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    NoE = 0
    Call ChangeStatus(False)
    txtWord.Text = ""
    txtWord.SetFocus
End Sub

Private Sub cmdDong_Click()
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    Unload frmWManage
End Sub

Private Sub cmdLuu_Click()
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    If NoE = 1 Then
        Call Save
    ElseIf NoE = 2 Then
        Call Edit
    End If
    txtWord.SetFocus
End Sub

Private Sub cmdReLoad_Click()
Dim e As Integer
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    Me.MousePointer = MousePointerConstants.vbHourglass
    
    Set Dic = New clsWordDic
    Set WC = New clsWordClass
    
    frmFlash.lblCaption.Caption = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(76) & ChrW(111) & ChrW(97) & ChrW(100) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    
    e = Dic.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        GoTo Label
    End If
    
    e = WC.LoadDic
    If e <> 0 Then
        Call ErrorHandle(e)
        GoTo Label
    End If
    
    Call ChangeStatus(False)
    lstWord.Visible = False
    lstWord1.Visible = True
    Call LoadlstWord
    lstWord.Visible = True
    lstWord1.Visible = False
    Call LoadlstWC
    
    txtWord.Text = ""
    Unload frmFlash
Label:
    Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub cmdSuadoi_Click()
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    NoE = 2
    Call ChangeStatus(True)
    lstWCoWord.SetFocus
End Sub

Private Sub cmdTaomoi_Click()
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    NoE = 1
    Call ChangeStatus(True)
    txtWord.Text = ""
    lstWCoWord.Clear
    txtWord.SetFocus
End Sub

Private Sub cmdXoa_Click()
Dim e As Integer
Dim t As Long
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    If lstWord.ListIndex = -1 Then Exit Sub
    If MsgBox("Ban co chac chan xoa hay khong?", vbOKCancel) = vbOK Then
        Me.MousePointer = MousePointerConstants.vbHourglass
        e = Dic.DelWord(Trim(lstWord.Text))
        If e <> 0 Then
            Call ErrorHandle(e)
            Me.MousePointer = MousePointerConstants.vbDefault
            Exit Sub
        End If
        DoEvents
        e = Dic.SaveDic
        If e <> 0 Then
            Call ErrorHandle(e)
            Me.MousePointer = MousePointerConstants.vbDefault
            Exit Sub
        End If
        t = lstWord.ListIndex
        lstWord.RemoveItem t
        lstType.RemoveItem t
        txtWord.Text = lstWord.Text
        Call LoadlstWCoWord
        lblTongso.Caption = ChrW(84) & ChrW(7893) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(58) & " " & Dic.WCount
        Me.MousePointer = MousePointerConstants.vbDefault
    End If
End Sub

Private Sub Form_Load()
Dim e As Integer
Dim i As Long, j As Long
    Me.Show
    Me.MousePointer = MousePointerConstants.vbHourglass
    frmFlash.lblCaption.Caption = ChrW(84) & ChrW(104) & ChrW(105) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(7863) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(46) & ChrW(32) & ChrW(88) & ChrW(105) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46)
    frmFlash.Show
    frmFlash.Refresh
    
    Me.Caption = "Quan Ly Tu Dien - Phien Ban " & App.Major & "." & App.Minor
    Call AddCaption
    lblTongso.Caption = ChrW(84) & ChrW(7893) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(58) & " " & Dic.WCount
    
    NoE = 0
    MouseDown = False

    
    Call ChangeStatus(False)
    lstWord.Visible = False
    lstWord1.Visible = True
    Call LoadlstWord
    lstWord.Visible = True
    lstWord1.Visible = False
    Call LoadlstWC
    Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub LoadlstWord()
Dim i As Long, j As Long, i1 As Long
    lstWord.Clear
    lstType.Clear
    For i = 1 To Dic.PCount
        For j = 1 To Dic.WiPCount(i)
            DoEvents
            lstWord.AddItem Dic.WordCont(i, j)
            lstType.AddItem Dic.WordClass(i, j)
        Next j
    Next i
    Unload frmFlash
    lblTongso.Caption = ChrW(84) & ChrW(7893) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(58) & " " & Dic.WCount
End Sub
' tai tu loai vao lstWC
Private Sub LoadlstWC()
Dim i As Long, j As Long
Dim ok As Boolean
Dim Pos As Long
    lstWC.Clear
    If lstWCoWord.ListCount = 0 Then
        For i = 1 To WC.Count
            lstWC.AddItem WC.Sign(i) & " - " & WC.Sense(i)
        Next i
    Else
        For i = 1 To WC.Count
            ok = False
            For j = 0 To lstWCoWord.ListCount - 1
                Pos = InStr(lstWCoWord.List(j), " ")
                If Mid(lstWCoWord.List(j), 1, Pos - 1) = WC.Sign(i) Then
                    ok = True
                    Exit For
                End If
            Next j
            If ok = False Then lstWC.AddItem WC.Sign(i) & " - " & WC.Sense(i)
        Next i
    End If
End Sub
' dua tu loai cua tu duoc chon vao lstWCoWord
Private Sub LoadlstWCoWord()
Dim Loca As New clsLocation
Dim sign1 As String, sense1 As String
Dim i As Long
Dim st As String
Dim Pos As Long
    lstWCoWord.Clear
    For i = 0 To lstWord.ListCount - 1
        If lstWord.Selected(i) Then
            st = lstType.List(i)
            Pos = InStr(st, "|")
            Do While Pos > 0
                sign1 = Mid(st, 1, Pos - 1)
                st = Mid(st, Pos + 1, Len(st) - Pos)
                Set Loca = WC.FindWC(sign1)
                If Loca.ok = 0 Then
                    sense1 = WC.Sense(Loca.x)
                    lstWCoWord.AddItem sign1 & " - " & sense1
                Else
                    MsgBox "Khong co tu can tim !", , App.Title
                End If
                Pos = InStr(st, "|")
            Loop
            Exit For
        End If
    Next i
    Call LoadlstWC
End Sub

Private Sub txtWord_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case NoE
            Case 0: Call Find
            Case 1: Call Save
            Case 2: Call Edit
        End Select
    End If
End Sub

Private Sub txtWord_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim st As String
    If KeyCode = vbKeyBack Then
    Else
        If NoE = 0 Then
            st = UniLCase(Trim(txtWord.Text))
            lstWord.ListIndex = SearchInlstWord(st)
        End If
    End If
End Sub

Private Sub lstWord_Click()
    If MouseDown Then
        txtWord.Text = ""
        txtWord.Text = lstWord.Text
        Call LoadlstWCoWord
    End If
    MouseDown = False
End Sub

Private Sub lstWord_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseDown = True
End Sub

Private Sub lstWord_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    txtWord.Text = ""
    txtWord.Text = lstWord.Text
    Call LoadlstWCoWord
End Sub

Private Sub Find()
Dim i As Long
    i = lstWord.ListIndex
    If i <> -1 Then
        txtWord.Text = lstWord.List(i)
        lstWord.SetFocus
        Call LoadlstWCoWord
    End If
End Sub

Private Sub Save()
Dim vWord As New clsWord
Dim i As Long, st As String
Dim e As Integer
Dim Pos As Long
    DoEvents
    If Trim(txtWord.Text) = "" Then
        MsgBox "Tu khong the rong", , App.Title
        txtWord.SetFocus
    Else
        Me.MousePointer = MousePointerConstants.vbHourglass
        vWord.WordCont = UniLCase(Trim(txtWord.Text))
        'tao chuoi tu loai
        st = ""
        For i = 0 To lstWCoWord.ListCount - 1
            Pos = InStr(lstWCoWord.List(i), " ")
            st = st & Mid(lstWCoWord.List(i), 1, Pos - 1) & "|"
        Next i
        vWord.WordClass = st
        'them vao tu dien
        e = Dic.AddWord(vWord)
        If e <> 0 Then
            Call ErrorHandle(e)
            Me.MousePointer = MousePointerConstants.vbDefault
        Else
            'can xu ly
            i = SearchForInsert(vWord.WordCont)
            lstWord.AddItem vWord.WordCont, i
            lstType.AddItem vWord.WordClass, i
            'Call LoadlstWord
        End If
        e = Dic.SaveDic
        If e <> 0 Then Call ErrorHandle(e)
        Me.MousePointer = MousePointerConstants.vbDefault
    End If
    NoE = 0
    Call ChangeStatus(False)
    lblTongso.Caption = ChrW(84) & ChrW(7893) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(115) & ChrW(7889) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110) & ChrW(58) & " " & Dic.WCount
End Sub

Private Sub Edit()
Dim Loca As New clsLocation
Dim st As String, i As Long, t As Long
Dim e As Integer
Dim Pos As Long
Dim Word As clsWord
    DoEvents
    If Trim(txtWord.Text) = "" Then
        MsgBox "Tu khong the rong", , App.Title
    Else
        Me.MousePointer = MousePointerConstants.vbHourglass
        st = ""
        For i = 0 To lstWCoWord.ListCount - 1
            Pos = InStr(lstWCoWord.List(i), " ")
            st = st & Mid(lstWCoWord.List(i), 1, Pos - 1) & "|"
        Next i
        Set Loca = Dic.FindWord(Trim(lstWord.Text))
        If Loca.ok = 0 Then
            If Trim(txtWord.Text) <> lstWord.Text Then
                Set Word = New clsWord
                Word.WordCont = UniLCase(Trim(txtWord.Text))
                Word.WordClass = st
                'xoa tu
                e = Dic.DelWord(lstWord.Text)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    GoTo Label
                End If
                t = lstWord.ListIndex
                lstWord.RemoveItem t
                lstType.RemoveItem t
                'them tu
                e = Dic.AddWord(Word)
                If e <> 0 Then
                    Call ErrorHandle(e)
                    GoTo Label
                Else
                    i = SearchForInsert(Word.WordCont)
                    lstWord.AddItem Word.WordCont, i
                    lstType.AddItem Word.WordClass, i
                End If
            Else
                Dic.WordClass(Loca.x, Loca.y) = st
                i = lstWord.ListIndex
                If i <> -1 Then
                    lstWord.List(i) = UniLCase(Trim(txtWord.Text))
                    lstType.List(i) = st
                End If
'                For i = 0 To lstWord.ListCount - 1
'                    If lstWord.Selected(i) = True Then
'
'                    End If
'                Next i
            End If
        Else
            Call ErrorHandle(NoHaveWord)
            GoTo Label
        End If
        e = Dic.SaveDic
        If e <> 0 Then Call ErrorHandle(e)
    End If
Label:
    NoE = 0
    Call ChangeStatus(False)
    Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub cmdLeft_Click()
Dim i As Long
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    For i = 0 To lstWC.ListCount - 1
        If lstWC.Selected(i) Then
            lstWCoWord.AddItem lstWC.List(i)
            Call LoadlstWC
            Exit For
        End If
    Next i
End Sub

Private Sub cmdALeft_Click()
Dim i As Long
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    For i = 0 To lstWC.ListCount - 1
        lstWCoWord.AddItem lstWC.List(i)
    Next i
    Call LoadlstWC
End Sub

Private Sub cmdARight_Click()
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    lstWCoWord.Clear
    Call LoadlstWC
End Sub

Private Sub cmdRight_Click()
Dim i As Long
    If Me.MousePointer = MousePointerConstants.vbHourglass Then Exit Sub
    For i = 0 To lstWCoWord.ListCount - 1
        If lstWCoWord.Selected(i) Then
            lstWCoWord.RemoveItem i
            Call LoadlstWC
            Exit Sub
        End If
    Next i
End Sub

Private Function SearchForInsert(ByVal st As String) As Long
Dim hi, lo As Long
Dim Mid As Long
    hi = lstWord.ListCount - 1
    lo = 0
    Do While hi >= lo
        Mid = (hi + lo) \ 2
        Select Case SoSanh(lstWord.List(Mid), st)
            Case 1: hi = Mid - 1
            Case -1: lo = Mid + 1
            Case 0: Exit Do
        End Select
    Loop
    If Mid = hi Then
        SearchForInsert = Mid + 1
    ElseIf Mid = lo Then
        SearchForInsert = Mid
    End If
End Function

Private Function SearchInlstWord(ByVal st As String) As Long
Dim hi, lo As Long
Dim Mid As Long
    hi = lstWord.ListCount - 1
    lo = 0
    Do While hi >= lo
        Mid = (hi + lo) \ 2
        Select Case SoSanh(lstWord.List(Mid), st)
            Case 1: hi = Mid - 1
            Case -1: lo = Mid + 1
            Case 0: Exit Do
        End Select
    Loop
    SearchInlstWord = Mid
End Function

Private Function SoSanh(ByVal st1 As String, ByVal st2 As String) As Integer
Dim l As Integer, nho As Integer
Dim i As Integer
    If Len(st1) > Len(st2) Then
        l = Len(st2)
    Else
        l = Len(st1)
    End If
    i = 1
    nho = 0
    Do While i <= l
        If Mid(st1, i, 1) > Mid(st2, i, 1) Then
            nho = 1
            Exit Do
        ElseIf Mid(st1, i, 1) < Mid(st2, i, 1) Then
            nho = -1
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    If i > l Then
        If Len(st1) = Len(st2) Then
            SoSanh = 0
        ElseIf Len(st1) > Len(st2) Then
            SoSanh = 1
        Else
            SoSanh = -1
        End If
    Else
        SoSanh = nho
    End If
End Function

Private Sub AddCaption()
    lblHeader.Caption = ChrW(81) & ChrW(85) & ChrW(7842) & ChrW(78) & ChrW(32) & ChrW(76) & ChrW(221) & ChrW(32) & ChrW(84) & ChrW(7914) & ChrW(32) & ChrW(272) & ChrW(73) & ChrW(7874) & ChrW(78)
    lblTitle.Caption = ChrW(66) & ChrW(7855) & ChrW(116) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(116) & ChrW(7843) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(86) & ChrW(105) & ChrW(7879) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(225) & ChrW(121) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104)
    lblFrame1.Caption = ChrW(84) & ChrW(7915) & ChrW(32) & ChrW(118) & ChrW(7921) & ChrW(110) & ChrW(103)
    cmdLuu.Caption = ChrW(76) & ChrW(432) & ChrW(117)
    cmdBoqua.Caption = ChrW(66) & ChrW(7887) & ChrW(32) & ChrW(81) & ChrW(117) & ChrW(97)
    lblFrame2.Caption = ChrW(84) & ChrW(7915) & ChrW(32) & ChrW(76) & ChrW(111) & ChrW(7841) & ChrW(105)
    lblTu.Caption = ChrW(84) & ChrW(236) & ChrW(109)
    lblHelp.Caption = ChrW(72) & ChrW(227) & ChrW(121) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(99) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(236) & ChrW(109) & ChrW(32) & ChrW(114) & ChrW(7891) & ChrW(105) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7845) & ChrW(110) & ChrW(32) & ChrW(69) & ChrW(110) & ChrW(116) & ChrW(101) & ChrW(114)
    lblDSTu.Caption = ChrW(68) & ChrW(97) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(109) & ChrW(7909) & ChrW(99) & ChrW(32) & ChrW(116) & ChrW(7915)
    cmdDong.Caption = ChrW(272) & ChrW(243) & ChrW(110) & ChrW(103)
    cmdTaomoi.Caption = ChrW(84) & ChrW(7841) & ChrW(111) & ChrW(32) & ChrW(77) & ChrW(7899) & ChrW(105)
    cmdSuadoi.Caption = ChrW(83) & ChrW(7917) & ChrW(97) & ChrW(32) & ChrW(272) & ChrW(7893) & ChrW(105)
    cmdXoa.Caption = ChrW(88) & ChrW(111) & ChrW(225)
    cmdReLoad.Caption = ChrW(67) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(273) & ChrW(105) & ChrW(7875) & ChrW(110)
End Sub
