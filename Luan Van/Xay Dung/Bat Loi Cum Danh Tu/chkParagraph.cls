VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkParagraph Class - a class demonstrating a paragraph
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private lStart As Long
Private lLength As Long
'---------------------doan-----------------
Public Function Copy(ByVal vPar As chkParagraph) As Integer
On Error GoTo Result
    Copy = 0
    lStart = vPar.ParagraphStart
    lLength = vPar.ParagraphLength
    Exit Function
Result:
    Copy = 1000
End Function

Public Property Get ParagraphStart() As Long
    ParagraphStart = lStart
End Property

Public Property Let ParagraphStart(ByVal i As Long)
    lStart = i
End Property

Public Property Get ParagraphLength() As Long
    ParagraphLength = lLength
End Property

Public Property Let ParagraphLength(ByVal i As Long)
    lLength = i
End Property
'--------------------cau---------------
Private Function IsEndSentence(ByVal st As String) As Boolean
    IsEndSentence = False
    If st <> "" Then
        Select Case AscW(st)
            Case 46: IsEndSentence = True 'dau cham
            Case 63: IsEndSentence = True 'dau cham hoi
            Case 33: IsEndSentence = True 'dau cham cam
        End Select
    End If
End Function

Public Function SentenceSplit(ByVal TB As MSForms.TextBox, sen As chkSentence) As Boolean
Dim ls As Long, ll As Long, e As Integer
Dim vSen As New chkSentence
Dim tam As Boolean
    SentenceSplit = False
    If sen Is Nothing Then
        tam = False
    Else
        tam = sen.IgnoreGrammarcheck
    End If
    ll = lStart + lLength
    ls = lStart
    TB.SelStart = ls
    TB.SelLength = 1
    Do While TB.SelStart < ll
        If IsEndSentence(TB.SelText) Then
            If ls < TB.SelStart Then
                Set sen = New chkSentence
                vSen.IgnoreGrammarcheck = tam
                vSen.SentenceStart = ls
                vSen.SentenceLength = TB.SelStart - ls
                vSen.SentenceSign = TB.SelText
                e = sen.Copy(vSen)
                If e <> 0 Then Call ErrorHandle(e)
                lStart = lStart + vSen.SentenceLength + 1
                lLength = lLength - (vSen.SentenceLength + 1)
                Exit Function
            Else
                ls = ls + 1
            End If
        End If
        TB.SelStart = TB.SelStart + 1
        TB.SelLength = 1
    Loop
    If TB.SelStart > ls Then
        Set sen = New chkSentence
        vSen.IgnoreGrammarcheck = tam
        vSen.SentenceStart = ls
        vSen.SentenceLength = TB.SelStart - ls
        vSen.SentenceSign = ""
        e = sen.Copy(vSen)
        If e <> 0 Then Call ErrorHandle(e)
        lStart = lStart + vSen.SentenceLength + 1
        lLength = lLength - vSen.SentenceLength - 1
    Else
        SentenceSplit = True
    End If
End Function
