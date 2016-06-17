Attribute VB_Name = "mMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000

Private m_bInIDE As Boolean

Public Sub Main()
   InitCommonControls
   Dim f As New frmSDITest
   f.Show
   Set f = Nothing
End Sub

Public Sub UnloadApp()
'   If Not InIDE() Then
'      SetErrorMode SEM_NOGPFAULTERRORBOX
'   End If
End Sub

Public Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function


