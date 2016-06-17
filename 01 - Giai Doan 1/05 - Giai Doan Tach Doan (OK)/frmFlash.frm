VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFlash 
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFlash.frx":23D2
      Top             =   240
      Width           =   480
   End
   Begin MSForms.Label lblCaption 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Caption         =   "khong biet"
      Size            =   "8493;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = "Quan Ly Tu Dien - Phien Ban " & App.Major & "." & App.Minor
    Me.MousePointer = MousePointerConstants.vbHourglass
End Sub
