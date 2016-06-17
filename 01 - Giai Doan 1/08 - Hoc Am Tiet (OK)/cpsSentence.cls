VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsSentence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsSentence Class - a class demonstrating a sentence
'-------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private lStart As Long
Private lLength As Long
Private sSign As String * 1
Private Syll As cpsSyllable
'Private AWord As cpsAWord
'Private cRS As cpsRS

Public Property Get Start() As Long
    Start = lStart
End Property

Public Property Let Start(ByVal i As Long)
    lStart = i
End Property

Public Property Get Length() As Long
    Length = lLength
End Property

Public Property Let Length(ByVal i As Long)
    lLength = i
End Property

Public Property Get Sign() As String
    Sign = sSign
End Property

Public Property Let Sign(ByVal s As String)
    sSign = s
End Property
'lay tung am tiet trong cau
Public Property Get Syllable(ByVal i As Long) As cpsWord
    Set Syllable = Syll.Syllable(i)
End Property
'so luong am tiet trong cau
Public Property Get WCount() As Long
    WCount = Syll.Count
End Property

'Public Property Get AW(ByVal i As Long, ByVal j As Long) As cpsWord
 '   Set AW = AWord.Word(i, j)
'End Property

'Public Property Get RSCount() As Long
'    RSCount = cRS.Count
'End Property

'Public Property Get RSSCount(ByVal i As Long) As Long
'    RSSCount = cRS.SCount(i)
'End Property

'Public Property Get RSItem(ByVal i As Long, ByVal j As Long) As cpsRSItem
'    Set RSItem = cRS.Word(i, j)
'End Property

'Private Sub AddUp(ByVal TB As MSForms.TextBox, ByVal i As Long, ByVal j As Long, ByVal st As String)
'Dim k As Integer
'    If (j = AWord.Count) And (i = 1) Then cRS.AddSentence st & "|"
'    If (j = AWord.Count - 1) And (i = 2) Then cRS.AddSentence st & "|"
'    If (j = AWord.Count - 2) And (i = 3) Then cRS.AddSentence st & "|"
'    If i + j <= AWord.Count Then
'        For k = 1 To 3
'            If AWord.Word(k, j + i).x <> 0 Then
'                AddUp TB, k, j + i, st & "|" & k & "|" & j + i
'            End If
'        Next k
'    End If
'End Sub

'Public Function AddUpWord(ByVal TB As MSForms.TextBox) As Integer
'Dim i As Long
'Dim st As String
'    For i = 1 To 3
'        If AWord.Word(i, 1).x <> 0 Then
'            st = i & "|" & 1
'            AddUp TB, i, 1, st
'        End If
'    Next i
'End Function

'Private Function AddUpSyllables(ByVal TB As MSForms.TextBox, ByVal i As Long, ByVal j As Long) As String
'Dim tu As String
'    If j <= 0 Or j > 3 Then Exit Function
'    TB.SelStart = Syll.Syllable(i).Start
'    TB.SelLength = Syll.Syllable(i).Length
'   'chua xu ly tb.seltext
'    tu = TB.SelText
'    If j >= 2 Then
'        TB.SelStart = Syll.Syllable(i + 1).Start
'        TB.SelLength = Syll.Syllable(i + 1).Length
'        tu = tu & " " & TB.SelText
'    End If
'    If j = 3 Then
'        TB.SelStart = Syll.Syllable(i + 2).Start
'        TB.SelLength = Syll.Syllable(i + 2).Length
'        tu = tu & " " & TB.SelText
'    End If
'    AddUpSyllables = tu
'End Function
'Tach tu dua tren cac am tiet da duoc tach
'Public Function WordSplit(TB As MSForms.TextBox) As Integer
'On Error GoTo Result
'Dim i As Long, j As Long, tu As String
'Dim vWord As New cpsWord
'Dim loca As New clsLocation
'    WordSplit = 0
'    If Me.WCount > 0 Then
'        AWord.Count = Me.WCount
'        For i = 1 To Me.WCount
'            vWord.Start = Syll.Syllable(i).Start
'            For j = 1 To 3
'                If i + j - 1 <= Me.WCount Then
'                    If j = 1 Then
'                        vWord.Length = Syll.Syllable(i).Length
'                    Else
'                        vWord.Length = Syll.Syllable(i + j - 1).Start - vWord.Start + Syll.Syllable(i + j - 1).Length
'                    End If
'                    tu = AddUpSyllables(TB, i, j)
'                    tu = UniLCase(tu)
'                    Set loca = Dic.FindWord(tu)
'                    If loca.ok = 0 Then
'                        vWord.x = loca.x
'                        vWord.y = loca.y
'                    Else
'                        vWord.x = 0
'                        vWord.y = 0
'                    End If
'                    If AWord.AddWord(j, i, vWord) <> 0 Then GoTo Result
'                End If
'            Next j
'        Next i
'    End If
'    Exit Function
'Result:
'    WordSplit = 1000
'End Function
'Tach cac am tiet
Public Function SyllSplit(TB As MSForms.TextBox) As Integer
On Error GoTo Result
Dim l As Long, ls As Long, ll As Long
Dim vWord As New cpsWord
    SyllSplit = 0
    l = lStart + lLength
    ls = lStart
    Do While ls < l
        ll = l
        If GetWord(TB, ls, ll) <> 0 Then GoTo Result
        If ll = 0 Then Exit Do
        vWord.Start = ls
        vWord.Length = ll
        If Syll.AddSyll(vWord) <> 0 Then GoTo Result
        ls = ls + ll
    Loop
    Exit Function
Result:
    SyllSplit = 1000
End Function

Public Function Copy(ByVal vSen As cpsSentence) As Integer
On Error GoTo Result
    Copy = 0
    lStart = vSen.Start
    lLength = vSen.Length
    sSign = vSen.Sign
    Exit Function
Result:
    Copy = 1000
End Function

Private Function IsBeginWord(ByVal st As String) As Boolean
    IsBeginWord = True
    If st <> "" Then
        Select Case AscW(st)
            Case 32: IsBeginWord = False 'khoang trong
            Case 9: IsBeginWord = False 'dau tab
            'Case 44: IsBeginWord = False 'dau phay
        End Select
    End If
End Function

Private Function IsEndWord(ByVal st As String) As Boolean
    IsEndWord = False
    If st <> "" Then
        Select Case AscW(st)
            Case 32: IsEndWord = True 'khoang trong
            Case 9: IsEndWord = True 'dau tab
            Case 44: IsEndWord = True 'dau phay
            Case 40: IsEndWord = True 'mo ngoac
            Case 41: IsEndWord = True 'dong ngoac
        End Select
    End If
End Function

Private Sub Class_Initialize()
    lStart = 0
    lLength = 0
    sSign = ""
    Set Syll = New cpsSyllable
    'Set AWord = New cpsAWord
    'Set cRS = New cpsRS
End Sub
'Lay tu
Private Function GetWord(TB As MSForms.TextBox, ls As Long, ll As Long) As Integer
On Error GoTo Result
    GetWord = 0
    Do While ls < ll
        TB.SelStart = ls
        TB.SelLength = 1
        If Not IsBeginWord(TB.SelText) Then
            ls = ls + 1
        Else
            Exit Do
        End If
    Loop
    If ls = ll Then
        ll = 0
        Exit Function
    Else
        TB.SelStart = ls
        TB.SelLength = 1
        If TB.SelText = "," Or TB.SelText = "(" Or TB.SelText = ")" Then
            ll = 1
            Exit Function
        End If
        TB.SelStart = ls + 1
        TB.SelLength = 1
        Do While TB.SelStart < ll
            If Not IsEndWord(TB.SelText) Then
                TB.SelStart = TB.SelStart + 1
                TB.SelLength = 1
            Else
                Exit Do
            End If
        Loop
        ll = TB.SelStart - ls
    End If
    Exit Function
Result:
    GetWord = 1000
End Function
