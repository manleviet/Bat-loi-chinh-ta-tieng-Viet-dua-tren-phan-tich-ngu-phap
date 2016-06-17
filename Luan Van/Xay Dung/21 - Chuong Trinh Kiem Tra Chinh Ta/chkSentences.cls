VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkSentences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkSentences Class - a class demonstrating sentences in paragraph
'-------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private TB As MSForms.TextBox
Private iStart As Long


Public Sub Init(TTextBox)
    Set TB = TTextBox
End Sub

Public Sub SyllableCheck()
Dim i As Long
    
End Sub

Private Function SentenceSplit() As Integer
On Error GoTo Result
Dim ls As Long
Dim vSen As New cpsSentence
    'Thuc hien tach cau
    SentenceSplit = 0
    TB.SelStart = 0
    TB.SelLength = 1
    ls = TB.SelStart
    Do While TB.SelLength = 1
        If IsEndSentence(TB.SelText) Then
            If ls < TB.SelStart Then
                iItem = iItem + 1
                ReDim Preserve Item(iItem)
                Set Item(iItem) = New cpsSentence
                vSen.Start = ls
                vSen.Length = TB.SelStart - ls
                vSen.Sign = TB.SelText
                If Item(iItem).Copy(vSen) <> 0 Then GoTo Result
            End If
            ls = TB.SelStart + 1
        End If
        TB.SelStart = TB.SelStart + 1
        TB.SelLength = 1
    Loop
    If TB.SelStart > ls Then
        iItem = iItem + 1
        ReDim Preserve Item(iItem)
        Set Item(iItem) = New cpsSentence
        vSen.Start = ls
        vSen.Length = TB.SelStart - ls
        vSen.Sign = ""
        If Item(iItem).Copy(vSen) <> 0 Then GoTo Result
    End If
    Exit Function
Result:
    SentenceSplit = 1000
End Function

Private Function IsEndSentence(ByVal st As String) As Boolean
    IsEndSentence = False
    If st <> "" Then
        Select Case AscW(st)
            Case 46: IsEndSentence = True 'dau cham
            Case 63: IsEndSentence = True 'dau cham hoi
            Case 33: IsEndSentence = True 'dau cham cam
            Case 13: IsEndSentence = True 'ky tu xuong dong
        End Select
    End If
End Function
