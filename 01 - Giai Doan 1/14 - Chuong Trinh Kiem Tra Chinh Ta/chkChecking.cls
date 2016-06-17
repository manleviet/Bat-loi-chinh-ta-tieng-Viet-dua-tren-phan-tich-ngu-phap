VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkChecking Class - a class to Vietnamese checking
'-------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private TB As MSForms.TextBox
Private par As chkParagraph
Private sen As chkSentence
Private lStart As Long
Private lLength As Long
Private IgnoreAll As chkIgnoreAll
Private Replace As chkReplaceAll
Private Ignore As chkIgnore
Private sStop As Boolean
'--------------------khoi tao--------------
Private Sub Class_Initialize()
    lStart = 0
    lLength = -1
End Sub

Public Sub Init(TTextBox)
    Set TB = TTextBox
End Sub
'-----------------tach doan------------------
Private Function IsEndParagraph(ByVal tu As String) As Boolean
    IsEndParagraph = False
    If tu <> "" Then
        Select Case AscW(tu)
            Case 13: IsEndParagraph = True
        End Select
    End If
End Function

Public Function ParagraphSplit() As Boolean
Dim ls As Long
Dim vPar As New chkParagraph
Dim e As Integer
    ParagraphSplit = False
    TB.SelStart = lStart + lLength + 1
    TB.SelLength = 1
    ls = TB.SelStart
    Do While TB.SelLength = 1
        If IsEndParagraph(TB.SelText) Then
            If ls < TB.SelStart Then
                Set par = New chkParagraph
                vPar.ParagraphStart = ls
                vPar.ParagraphLength = TB.SelStart - ls
                e = par.Copy(vPar)
                If e <> 0 Then Call ErrorHandle(e)
                lStart = vPar.ParagraphStart
                lLength = vPar.ParagraphLength
                Exit Function
            Else
                ls = ls + 1
            End If
        End If
        TB.SelStart = TB.SelStart + 1
        TB.SelLength = 1
    Loop
    If TB.SelStart > ls Then
        Set par = New chkParagraph
        vPar.ParagraphStart = ls
        vPar.ParagraphLength = TB.SelStart - ls
        e = par.Copy(vPar)
        If e Then Call ErrorHandle(e)
        lStart = vPar.ParagraphStart
        lLength = vPar.ParagraphLength
    Else
        ParagraphSplit = True
    End If
End Function
'------------------kiem tra-----------------
Public Function SyllableCheck() As Integer
On Error GoTo Result
    Set IgnoreAll = New chkIgnoreAll
    Set Replace = New chkReplaceAll
    Set Ignore = New chkIgnore
    sStop = False
    Do While Not ParagraphSplit
        Do While Not par.SentenceSplit(TB, sen)
            If sStop Then Exit Do
            If sen.SyllSplit(TB) <> 0 Then GoTo Result
            'Dim i As Long
            'For i = 1 To sen.SyllCount
            '    TB.SelStart = sen.Syllable(i).Start
            '    TB.SelLength = sen.Syllable(i).Length
            '    MsgBox TB.SelText
            'Next i
            If sen.SyllCount > 0 Then
                If sen.WordSplit(TB) <> 0 Then GoTo Result
                'Dim j As Long
                'For i = 1 To sen.SyllCount
                '    For j = 1 To 3
                '        TB.SelStart = sen.AW(j, i).Start
                '        TB.SelLength = sen.AW(j, i).Length
                '        MsgBox j & " " & i & " " & TB.SelText
                '    Next j
                'Next i
                If Not sen.SyllableCheck(TB, sStop, Ignore, IgnoreAll, Replace) Then
                    lLength = -1
                    Exit Do
                End If
            Else
                If sen.SentenceSign <> "" Then
                    TB.SelStart = sen.SentenceStart
                    TB.SelLength = sen.SentenceLength + 1
                    MsgBox "Cau '" & TB.SelText & "' la sai."
                End If
            End If
        Loop
    Loop
    Exit Function
Result:
    SyllableCheck = 1000
End Function
