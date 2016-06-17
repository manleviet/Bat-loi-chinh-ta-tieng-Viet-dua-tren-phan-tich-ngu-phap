VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFlash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
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
    Me.Caption = "To Hop Tu Loai Dung - Phien Ban " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
