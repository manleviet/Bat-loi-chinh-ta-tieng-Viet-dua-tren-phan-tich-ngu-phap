VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chkSentence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: chkSentence Class - a class demonstrating a sentence
'-------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private lStart As Long
Private lLength As Long
Private sSign As String * 1
Private Syll As chkSyllable
Private AWord As chkAWord
Private bIgnoreGrammarCheck As Boolean
Private cRS As chkRS
Private cSetWC As chkSetWC
Private cEarley As cepEarlyParse
'---------------------khoi tao---------------------
Private Sub Class_Initialize()
    lStart = 0
    lLength = 0
    sSign = ""
    bIgnoreGrammarCheck = False
    Set Syll = New chkSyllable
    Set AWord = New chkAWord
    Set cEarley = New cepEarlyParse
End Sub

Public Property Get IgnoreGrammarcheck() As Boolean
    IgnoreGrammarcheck = bIgnoreGrammarCheck
End Property

Public Property Let IgnoreGrammarcheck(b As Boolean)
    bIgnoreGrammarCheck = b
End Property
' sao chep cau
Public Function Copy(ByVal vSen As chkSentence) As Integer
On Error GoTo Result
    Copy = 0
    lStart = vSen.SentenceStart
    lLength = vSen.SentenceLength
    sSign = vSen.SentenceSign
    bIgnoreGrammarCheck = vSen.IgnoreGrammarcheck
    Exit Function
Result:
    Copy = 1000
End Function
'---------------cau-----------------------
Public Property Get SentenceStart() As Long
    SentenceStart = lStart
End Property

Public Property Let SentenceStart(ByVal i As Long)
    lStart = i
End Property
'do dai cau
Public Property Get SentenceLength() As Long
    SentenceLength = lLength
End Property

Public Property Let SentenceLength(ByVal i As Long)
    lLength = i
End Property
'ky tu ket thuc cau
Public Property Get SentenceSign() As String
    SentenceSign = sSign
End Property
'ky tu ket thuc cau
Public Property Let SentenceSign(ByVal s As String)
    sSign = s
End Property
'--------------------am tiet--------------------------
'kiem tra ky tu bat dau am tiet
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
'kiem tra ky tu ket thuc am tiet
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
'Tach cac am tiet
Public Function SyllSplit(TB As MSForms.TextBox) As Integer
On Error GoTo Result
Dim l As Long, ls As Long, ll As Long
Dim vWord As New chkWord
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
'lay tung am tiet trong cau
Public Property Get Syllable(ByVal i As Long) As chkWord
    Set Syllable = Syll.Syllable(i)
End Property
'so luong am tiet trong cau
Public Property Get SyllCount() As Long
    SyllCount = Syll.Count
End Property
'-------------------tu------------------------------
'lay tung phan tu cua mang AWord
Public Property Get AW(ByVal i As Long, ByVal j As Long) As chkWord
    Set AW = AWord.Word(i, j)
End Property
'gop mot,hai hay ba am tiet thanh mot tu
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
Public Function WordSplit(TB As MSForms.TextBox) As Integer
On Error GoTo Result
Dim i As Long, j As Long, tu As String
Dim vWord As New chkWord
Dim loca As New clsLocation
    WordSplit = 0
    If Me.SyllCount > 0 Then
        AWord.Count = Me.SyllCount
        For i = 1 To Me.SyllCount
            vWord.Start = Syll.Syllable(i).Start
            For j = 1 To 3
                If i + j - 1 <= Me.SyllCount Then
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
    WordSplit = 1000
End Function
'-----------kiem tra chinh ta---------------
'kiem tra loi am tiet
Private Function CheckError(ByVal i As Long) As Boolean
    CheckError = False
    Select Case i
        Case 1
            If AWord.Word(1, i).x = 0 And AWord.Word(2, i).x = 0 And AWord.Word(3, i).x = 0 Then
                CheckError = True
            End If
        Case 2:
            If AWord.Word(1, i).x = 0 And AWord.Word(2, i).x = 0 And AWord.Word(3, i).x = 0 Then
                If i - 1 <> 0 And AWord.Word(2, i - 1).x = 0 And AWord.Word(3, i - 1).x = 0 Then
                    CheckError = True
                End If
            End If
        Case Else:
            If AWord.Word(1, i).x = 0 And AWord.Word(2, i).x = 0 And AWord.Word(3, i).x = 0 Then
                If i - 1 <> 0 And AWord.Word(2, i - 1).x = 0 And AWord.Word(3, i - 1).x = 0 Then
                    If i - 2 <> 0 And AWord.Word(3, i - 2).x = 0 Then
                        CheckError = True
                    End If
                End If
            End If
    End Select
End Function

Private Function SpecialCaseProcessing(ByVal st As String) As eSpecialCharacter
    Select Case st
        Case ",": SpecialCaseProcessing = eSpecialCharacter.DauPhay
        Case "(": SpecialCaseProcessing = eSpecialCharacter.MoNgoac
        Case ")": SpecialCaseProcessing = eSpecialCharacter.DongNgoac
        Case Else: SpecialCaseProcessing = eSpecialCharacter.None
    End Select
End Function
'kiem tra am tiet
Public Function SyllableCheck(ByVal TB As MSForms.TextBox, bStop As Boolean, Ignore As chkIgnore, IgnoreAll As chkIgnoreAll, Replace As chkReplaceAll) As Boolean
Dim i As Long, tu As String
Dim loca As New clsLocation
Dim frmForm As New frmGoiY
    SyllableCheck = True
    For i = 1 To AWord.Count
        If CheckError(i) Then
            TB.SelStart = AWord.Word(1, i).Start
            TB.SelLength = AWord.Word(1, i).Length
            tu = TB.SelText
            '**************don dep khoang trong
            'kiem tra xem co phai dau phay, mo ngoac, dong ngoac khong
            Select Case SpecialCaseProcessing(tu)
                Case 0:
                    Set loca = Dic.FindSWord(tu) 'tim kiem trong tu dien
                    If loca.ok <> 0 Then
                        'kiem tra trong hai mang Ignore va Replace
                        'can chu y gia tri bIgnoreGrammarCheck
                        If Not Ignore.CheckIgnore(tu, i) Then
                            If IgnoreAll.CheckIgnore(tu) Then
                                bIgnoreGrammarCheck = True
                            Else
                                If Replace.CheckReplace(tu) Then
                                    TB.SelStart = AWord.Word(1, i).Start
                                    TB.SelLength = AWord.Word(1, i).Length
                                    TB.SelText = tu
                                    SyllableCheck = False
                                    Exit Function
                                Else
                                    'goi y sua loi
                                    TB.SelStart = lStart
                                    TB.SelLength = lLength + 1
                                    frmForm.txtNguCanh.Text = TB.SelText
                                    frmForm.txtAmTiet.Text = tu
                                    frmForm.SimilarSyllables
                                    frmForm.Show vbModal, frmKiemTra
                                    Select Case frmForm.Result
                                        Case eGoiY.BoQua:
                                            bIgnoreGrammarCheck = True
                                            If Ignore.AddIgnore(tu, i) <> 0 Then Call ErrorHandle(1000)
                                            SyllableCheck = False
                                            Exit Function
                                        Case eGoiY.BoQuaHet:
                                            bIgnoreGrammarCheck = True
                                            If IgnoreAll.AddIgnore(tu) <> 0 Then Call ErrorHandle(1000)
                                            SyllableCheck = False
                                            Exit Function
                                        Case eGoiY.Dung:
                                            bStop = True
                                            Exit Function
                                        Case eGoiY.ThayThe:
                                            TB.SelStart = AWord.Word(1, i).Start
                                            TB.SelLength = AWord.Word(1, i).Length
                                            TB.SelText = frmForm.txtAmTiet.Text
                                            SyllableCheck = False
                                            Exit Function
                                        Case eGoiY.ThayTheHet:
                                            If Replace.AddReplace(tu, frmForm.txtAmTiet.Text) <> 0 Then Call ErrorHandle(1000)
                                            TB.SelStart = AWord.Word(1, i).Start
                                            TB.SelLength = AWord.Word(1, i).Length
                                            TB.SelText = frmForm.txtAmTiet.Text
                                            SyllableCheck = False
                                            Exit Function
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Case 1, 2, 3:
                    tu = "!!vch!! " & tu
                    Set loca = Dic.FindWord(tu) 'tim kiem trong tu dien
                    If loca.ok <> 0 Then 'loi
                        Call ErrorHandle(1000)
                    Else
                        AWord.Word(1, i).x = loca.x
                        AWord.Word(1, i).y = loca.y
                    End If
            End Select
        End If
    Next i
    Set Ignore = New chkIgnore
End Function
'------------kiem tra loi chinh ta dua tren cu phap---------------
Public Function GrammarCheck(ByVal TB As MSForms.TextBox, bStop As Boolean, Ignore As chkIgnore, IgnoreAll As chkIgnoreAll, Replace As chkReplaceAll, ByVal cRule As crlRule) As Boolean
Dim e As Integer
Dim lLoRight As Long, i As Long, Max As Double, iMax As Long
Dim frmForm As New frmGoiYCuPhap
Dim tu As String, r As Double, xx As Long, yy As Long
    'can chu y gia tri bIgnoreGrammarCheck
    GrammarCheck = True
    If Not bIgnoreGrammarCheck Then
        Set cRS = New chkRS
        If GenerateRS <> 0 Then 'lap RS
            Exit Function
        End If
        Set cSetWC = New chkSetWC
        If GenerateWCS <> 0 Then ' lap WCS
            Exit Function
        End If
        lLoRight = cEarley.Parse(cRule, cSetWC)
        If lLoRight = -1 Then
            Max = 0
            For i = 1 To cSetWC.Count 'tim gia tri lon nhat theo ti le
                r = cSetWC.RNum(i) / Len(cSetWC.WCS(i))
                If r > Max Then
                    Max = r
                    iMax = i
                End If
            Next i
            If iMax = 0 Then iMax = 1
            xx = cRS.Word(cSetWC.RSItem(iMax), cSetWC.RNum(iMax) + 1).x
            yy = cRS.Word(cSetWC.RSItem(iMax), cSetWC.RNum(iMax) + 1).y
            TB.SelStart = AWord.Word(xx, yy).Start
            TB.SelLength = AWord.Word(xx, yy).Length
            tu = TB.SelText
            '***********don dep khoang trong
            If IgnoreAll.CheckIgnore(tu) Then
                Exit Function
            End If
            If Replace.CheckReplace(tu) Then
                TB.SelStart = AWord.Word(xx, yy).Start
                TB.SelLength = AWord.Word(xx, yy).Length
                TB.SelText = tu
                GrammarCheck = False
                Exit Function
            End If
            TB.SelStart = lStart
            TB.SelLength = lLength + 1
            frmForm.txtNguCanh.Text = TB.SelText
            frmForm.txtAmTiet.Text = tu
            frmForm.SimilarWord
            frmForm.Show vbModal, frmKiemTra
            Select Case frmForm.Result
                Case eGoiY.BoQua:
                    Exit Function
                Case eGoiY.BoQuaHet:
                    If IgnoreAll.AddIgnore(tu) <> 0 Then Call ErrorHandle(1000)
                    Exit Function
                Case eGoiY.Dung:
                    bStop = True
                    Exit Function
                Case eGoiY.ThayThe:
                    TB.SelStart = AWord.Word(xx, yy).Start
                    TB.SelLength = AWord.Word(xx, yy).Length
                    TB.SelText = frmForm.txtAmTiet.Text
                    GrammarCheck = False
                    Exit Function
                Case eGoiY.ThayTheHet:
                    If Replace.AddReplace(tu, frmForm.txtAmTiet.Text) <> 0 Then Call ErrorHandle(1000)
                    TB.SelStart = AWord.Word(xx, yy).Start
                    TB.SelLength = AWord.Word(xx, yy).Length
                    TB.SelText = frmForm.txtAmTiet.Text
                    GrammarCheck = False
                    Exit Function
            End Select
        End If
    End If
End Function
'--------------gop tu-----------------
'Public Property Get RSCount() As Long
'    RSCount = cRS.Count
'End Property

'Public Property Get RSSCount(ByVal i As Long) As Long
'    RSSCount = cRS.SCount(i)
'End Property

'Public Property Get RS(ByVal i As Long, ByVal j As Long) As cpsRSItem
'    Set RS = cRS.Word(i, j)
'End Property

Private Sub AddUp(ByVal i As Long, ByVal j As Long, ByVal st As String)
Dim k As Integer
    If (j = AWord.Count) And (i = 1) Then cRS.AddSentence st & "|"
    If (j = AWord.Count - 1) And (i = 2) Then cRS.AddSentence st & "|"
    If (j = AWord.Count - 2) And (i = 3) Then cRS.AddSentence st & "|"
    If i + j <= AWord.Count Then
        For k = 1 To 3
            If AWord.Word(k, j + i).x <> 0 Then
                AddUp k, j + i, st & "|" & k & "|" & j + i
            End If
        Next k
    End If
End Sub

Private Function GenerateRS() As Integer
On Error GoTo Result
Dim i As Long
Dim st As String
    GenerateRS = 0
    For i = 1 To 3
        If AWord.Word(i, 1).x <> 0 Then
            st = i & "|" & 1
            AddUp i, 1, st
        End If
    Next i
    Exit Function
Result:
    GenerateRS = 1000
End Function
'-----------------to hop tu loai---------------------
'Public Property Get WCSCount() As Long
'    WCSCount = cSetWC.Count
'End Property

'Public Property Get WCSRSItem(ByVal i As Long) As Long
'    WCSRSItem = cSetWC.RSItem(i)
'End Property

'Public Property Get WCS(ByVal i As Long) As String
'    WCS = cSetWC.WCS(i)
'End Property

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

Private Function GenerateWCS() As Integer
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



