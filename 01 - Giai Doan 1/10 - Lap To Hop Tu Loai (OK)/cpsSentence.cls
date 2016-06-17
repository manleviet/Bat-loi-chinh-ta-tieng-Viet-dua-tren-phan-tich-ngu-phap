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
Private AWord As cpsAWord
Private cRS As cpsRS
Private cSetWC As cpsSetWC
'-----------------khoi tao--------------------
Private Sub Class_Initialize()
    lStart = 0
    lLength = 0
    sSign = ""
    Set Syll = New cpsSyllable
    Set AWord = New cpsAWord
    Set cRS = New cpsRS
    Set cSetWC = New cpsSetWC
End Sub
'------------------cau------------------------
Public Property Get SentenceStart() As Long
    SentenceStart = lStart
End Property

Public Property Let SentenceStart(ByVal i As Long)
    lStart = i
End Property

Public Property Get SentenceLength() As Long
    SentenceLength = lLength
End Property

Public Property Let SentenceLength(ByVal i As Long)
    lLength = i
End Property

Public Property Get SentenceSign() As String
    SentenceSign = sSign
End Property

Public Property Let SentenceSign(ByVal s As String)
    sSign = s
End Property

Public Function Copy(ByVal vSen As cpsSentence) As Integer
On Error GoTo Result
    Copy = 0
    lStart = vSen.SentenceStart
    lLength = vSen.SentenceLength
    sSign = vSen.SentenceSign
    Exit Function
Result:
    Copy = 1000
End Function
'------------------am tiet----------------------
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
'Lay tu
Private Function GetSyll(TB As MSForms.TextBox, ls As Long, ll As Long) As Integer
On Error GoTo Result
    GetSyll = 0
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
    GetSyll = 1000
End Function
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
        If GetSyll(TB, ls, ll) <> 0 Then GoTo Result
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
'lay tung am tiet trong cau
Public Property Get Syllable(ByVal i As Long) As cpsWord
    Set Syllable = Syll.Syllable(i)
End Property
'so luong am tiet trong cau
Public Property Get SyllCount() As Long
    SyllCount = Syll.SyllCount
End Property
'----------------tu----------------------
Public Property Get AWCount() As Long
    AWCount = Syll.SyllCount
End Property

Public Property Get AW(ByVal i As Long, ByVal j As Long) As cpsWord
    Set AW = AWord.Word(i, j)
End Property

Private Function AddUpSyllables(ByVal TB As MSForms.TextBox, ByVal i As Long, ByVal j As Long) As String
Dim tu As String
    If j <= 0 Or j > 3 Then Exit Function
    TB.SelStart = Syll.Syllable(i).Start
    TB.SelLength = Syll.Syllable(i).Length
    'chua xu ly tb.seltext
    tu = TB.SelText
    If j >= 2 Then
        TB.SelStart = Syll.Syllable(i + 1).Start
        TB.SelLength = Syll.Syllable(i + 1).Length
        tu = tu & " " & TB.SelText
    End If
    If j = 3 Then
        TB.SelStart = Syll.Syllable(i + 2).Start
        TB.SelLength = Syll.Syllable(i + 2).Length
        tu = tu & " " & TB.SelText
    End If
    AddUpSyllables = tu
End Function
'Tach tu dua tren cac am tiet da duoc tach
Public Function AWSplit(TB As MSForms.TextBox) As Integer
'On Error GoTo Result
Dim i As Long, j As Long, tu As String
Dim vWord As New cpsWord
Dim loca As New clsLocation
    AWSplit = 0
    If Me.AWCount > 0 Then
        AWord.Count = Me.AWCount
        For i = 1 To Me.AWCount
            vWord.Start = Syll.Syllable(i).Start
            For j = 1 To 3
                If i + j - 1 <= Me.AWCount Then
                    If j = 1 Then
                        vWord.Length = Syll.Syllable(i).Length
                    Else
                        vWord.Length = Syll.Syllable(i + j - 1).Start - vWord.Start + Syll.Syllable(i + j - 1).Length
                    End If
                    tu = AddUpSyllables(TB, i, j)
                    tu = UniLCase(tu)
                    Set loca = Dic.FindWord(tu)
                    If loca.ok = 0 Then
                        vWord.x = loca.x
                        vWord.y = loca.y
                    Else
                        vWord.x = 0
                        vWord.y = 0
                    End If
                    If AWord.AddWord(j, i, vWord) <> 0 Then GoTo Result
                End If
            Next j
        Next i
    End If
    Exit Function
Result:
    AWSplit = 1000
End Function
'-----------------------gop tu------------------------
Public Property Get RSCount() As Long
    RSCount = cRS.Count
End Property

Public Property Get RSSCount(ByVal i As Long) As Long
    RSSCount = cRS.SCount(i)
End Property

Public Property Get RS(ByVal i As Long, ByVal j As Long) As cpsRSItem
    Set RS = cRS.Word(i, j)
End Property

Private Sub AddUp(ByVal TB As MSForms.TextBox, ByVal i As Long, ByVal j As Long, ByVal st As String)
Dim k As Integer
    If (j = AWord.Count) And (i = 1) Then cRS.AddSentence st & "|"
    If (j = AWord.Count - 1) And (i = 2) Then cRS.AddSentence st & "|"
    If (j = AWord.Count - 2) And (i = 3) Then cRS.AddSentence st & "|"
    If i + j <= AWord.Count Then
        For k = 1 To 3
            If AWord.Word(k, j + i).x <> 0 Then
                AddUp TB, k, j + i, st & "|" & k & "|" & j + i
            End If
        Next k
    End If
End Sub

Public Function GenerateRS(ByVal TB As MSForms.TextBox) As Integer
On Error GoTo Result
Dim i As Long
Dim st As String
    GenerateRS = 0
    For i = 1 To 3
        If AWord.Word(i, 1).x <> 0 Then
            st = i & "|" & 1
            AddUp TB, i, 1, st
        End If
    Next i
    Exit Function
Result:
    GenerateRS = 1000
End Function
'-----------------to hop tu loai---------------------
Public Property Get WCSCount() As Long
    WCSCount = cSetWC.Count
End Property

Public Property Get WCSRSItem(ByVal i As Long) As Long
    WCSRSItem = cSetWC.RSItem(i)
End Property

Public Property Get WCS(ByVal i As Long) As String
    WCS = cSetWC.WCS(i)
End Property

Private Sub Generate(ByVal st As String, ByVal r As Long, ByVal i As Long, ByVal l As Long, ByVal temp)
Dim p As Long
Dim tu As String
    If i > l Then
        Call cSetWC.AddWCS(st, r)
    Else
        Do While Len(temp(i)) <> 0
            p = InStr(1, temp(i), "|")
            tu = Left(temp(i), p - 1)
            temp(i) = Mid(temp(i), p + 1)
            If st = "" Then
                Generate tu, r, i + 1, l, temp
            Else
                Generate st & " " & tu, r, i + 1, l, temp
            End If
        Loop
    End If
End Sub

Public Function GenerateWCS() As Integer
On Error GoTo Result
Dim temp() As String
Dim iItem As Long
Dim i As Long, j As Long
    GenerateWCS = 0
    For i = 1 To cRS.Count
        iItem = cRS.SCount(i)
        ReDim temp(iItem)
        For j = 1 To iItem
            temp(j) = Dic.WordClass(AWord.Word(cRS.Word(i, j).x, cRS.Word(i, j).y).x, AWord.Word(cRS.Word(i, j).x, cRS.Word(i, j).y).y)
        Next j
        Generate "", i, 1, iItem, temp
    Next i
    Exit Function
Result:
    GenerateWCS = 1000
End Function
