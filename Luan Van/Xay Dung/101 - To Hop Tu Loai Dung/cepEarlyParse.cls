VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cepEarlyParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cepEarlyParse Class - a class demonstrating a early parser
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private Early As cepEarlyTable

Private Sub Class_Initialize()
    Set Early = New cepEarlyTable
End Sub
'phan tich cu phap
Public Function Parse(ByVal Rule As crlRule, WCS As cpsSetWC) As Long 'tra lai -1 neu sai; >0: vi tri dung
Dim DN As String, tu As String
Dim WCSet As String, SK As String
Dim i As Long, n As Long
    Parse = -1
    SK = ""
    For i = 1 To WCS.Count
        WCSet = AddEndSet(WCS.WCS(i))
        n = NSK(SK, WCSet)
        If n = 0 Then
            Early.Init
        Else
            SK = Left(SK, n * 3)
            WCSet = Mid(WCSet, n * 3 + 1)
            Early.Reset n + 1
        End If
        Do While True
            SK = SK & Left(WCSet, 3)
            DN = GetPredict(WCSet)
            If DN = "$$$" Then 'gia tri doan nhan la gia tri ket thuc chuoi
                If KiemTraKetThucDung(Rule) Then 'ket thuc dung
                    WCS.Right(i) = True
                    Parse = i
                    Exit Function
                Else 'ket thuc khong dung
                    WCS.Right(i) = False
                    WCS.RNum(i) = Early.Scan - 1
                    SK = Left(SK, Len(SK) - 3)
                    Exit Do
                End If
            Else 'chua het chuoi nhap
                'If Early.IsEmpty Then 'khong the tiep tuc
                '    WCS.Right(i) = False
                '    WCS.RNum(i) = Early.Scan - 1
                '    SK = Left(SK, Len(SK) - 3)
                '    Exit Do
                'End If
                Do While Not Early.IsEndCol
                    tu = Early.WordAfterDot
                    If tu <> "" Then
                        If IsNonTerminal(tu) Then
                            If DoanNhan(DN, tu, Rule) Then Exit Do
                        End If
                    Else
                        HoanThien
                    End If
                Loop
                If Early.IsEndCol Then 'khi het cot ma khong the sang cot khac
                    WCS.Right(i) = False
                    WCS.RNum(i) = Early.Scan - 1
                    SK = Left(SK, Len(SK) - 3)
                    Exit Do
                End If
            End If
        Loop
    Next i
End Function

Private Function DoanNhan(ByVal DN As String, ByVal tu As String, ByVal Rule As crlRule) As Boolean
Dim TRule() As String, iTRule As Long
Dim TNRule() As String, iTNRule As Long
Dim ok As Boolean
Dim i As Long, j As Long
    GetTRule Rule, tu, TRule
    iTRule = UBound(TRule)
    GetTNRule Rule, tu, TNRule
    iTNRule = UBound(TNRule)
    ok = False
    For i = 1 To iTRule
        If Right(TRule(i), 3) = DN Then
            ok = True
            Exit For
        End If
    Next i
    If ok Then
        For j = 1 To iTNRule
            If Mid(TNRule(j), 5, 3) = tu Then Early.AddRule 1, TNRule(j)
        Next j
        Early.SoKhop TRule(i)
        DoanNhan = True
    Else
        For i = 1 To iTNRule
            Early.AddRule 1, TNRule(i)
        Next i
        Early.IncScan
        DoanNhan = False
    End If
End Function

Private Sub HoanThien()
Dim VT As String, Ori As String
Dim i As Long
Dim aRule() As String
    VT = Early.RuleLeft
    If VT <> "ROT" Then
        Ori = Early.Origin
        Early.GetRuleWithRuleLeft Ori, VT, aRule
        For i = 1 To UBound(aRule)
            Early.AddRule 0, aRule(i)
        Next i
    End If
    Early.IncScan
End Sub

Private Function KiemTraKetThucDung(ByVal Rule As crlRule) As Boolean
Dim tu As String
Dim TNRule() As String
Dim iTNRule As Long
Dim i As Long
Dim VT As String
Dim Ori As String
Dim aRule() As String
    Do While Not Early.IsEndCol
        tu = Early.WordAfterDot
        If tu <> "" Then
            If IsNonTerminal(tu) Then 'doan nha
                GetTNRule Rule, tu, TNRule
                iTNRule = UBound(TNRule)
                For i = 1 To iTNRule
                    Early.AddRule 1, TNRule(i)
                Next i
            End If
        Else 'hoan thien
            VT = Early.RuleLeft
            If VT <> "ROT" Then
                Ori = Early.Origin
                Early.GetRuleWithRuleLeft Ori, VT, aRule
                For i = 1 To UBound(aRule)
                    Early.AddRule 0, aRule(i)
                Next i
            Else
                KiemTraKetThucDung = True
                Exit Function
            End If
        End If
        Early.IncScan
    Loop
    KiemTraKetThucDung = False
End Function
'lay cac luat chua ky tu ket thuc trong tu dien luat
Private Function GetTRule(ByVal Rule As crlRule, ByVal tu As String, r() As String) As Integer
On Error GoTo Result
Dim loca As New clsLocation
Dim TRule() As String
Dim iTRule As Long
Dim luat As String
Dim i As Long
    GetTRule = 0
    Set loca = Rule.FindRuleLeft(tu) ' tim trong chi muc luat
    If loca.ok <> 0 Then GoTo Result
    iTRule = 0
    ReDim TRule(iTRule)
    For i = 1 To Rule.RiPCount(loca.x) 'duyet qua trang luat
        luat = Rule.Rule(loca.x, i) 'lay mot luat
        If Len(luat) = 7 Then 'kiem tra luat co ky tu ket thuc
            If IsTerminal(Mid(luat, 5, 3)) Then
                iTRule = iTRule + 1
                ReDim Preserve TRule(iTRule)
                TRule(iTRule) = luat
            End If
        End If
    Next i
    r = TRule
    Exit Function
Result:
    GetTRule = 1000
End Function
'lay cac luat chua ky tu khong ket thuc trong tu dien luat
Private Function GetTNRule(ByVal Rule As crlRule, ByVal tu As String, r() As String) As Integer
On Error GoTo Result
Dim loca As New clsLocation
Dim TNRule() As String
Dim iTNRule As Long
Dim luat As String
Dim i As Long
    GetTNRule = 0
    Set loca = Rule.FindRuleLeft(tu) 'tim trong tu dien chi muc luat
    If loca.ok <> 0 Then GoTo Result
    iTNRule = 0
    ReDim TNRule(iTNRule)
    For i = 1 To Rule.RiPCount(loca.x) 'duyet qua trang luat
        luat = Rule.Rule(loca.x, i) 'lay mot luat
        If IsNonTerminal(Mid(luat, 5, 3)) Then 'kiem tra luat
            iTNRule = iTNRule + 1
            ReDim Preserve TNRule(iTNRule)
            TNRule(iTNRule) = luat
        End If
    Next i
    r = TNRule
    Exit Function
Result:
    GetTNRule = 1000
End Function

Private Function IsNonTerminal(ByVal tu As String) As Boolean
Dim p As Long
    IsNonTerminal = False
    p = InStr(1, NTer, Left(tu, 1))
    If p <> 0 Then IsNonTerminal = True
End Function

Private Function IsTerminal(ByVal tu As String) As Boolean
Dim p As Long
    IsTerminal = False
    p = InStr(1, Ter, Left(tu, 1))
    If p <> 0 Then IsTerminal = True
End Function

Private Function AddEndSet(ByVal st As String) As String
Dim p As Long
    p = InStr(1, st, " ")
    Do While p <> 0
        st = Left(st, p - 1) & Mid(st, p + 1)
        p = InStr(p, st, " ")
    Loop
    AddEndSet = st & "$$$"
End Function

Private Function GetPredict(st As String) As String
    GetPredict = Left(st, 3)
    st = Mid(st, 4)
End Function

Private Function NSK(ByVal SK As String, ByVal WCS As String) As Long
Dim p As Long
Dim tu As String
    NSK = 0
    tu = Left(SK, 3)
    SK = Mid(SK, 4)
    Do While tu <> ""
        p = InStr(1, WCS, tu)
        If p = 1 Then
            NSK = NSK + 1
            tu = Left(SK, 3)
            SK = Mid(SK, 4)
            WCS = Mid(WCS, 4)
        Else
            Exit Do
        End If
    Loop
End Function
