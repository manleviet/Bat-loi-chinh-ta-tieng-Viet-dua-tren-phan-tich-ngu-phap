VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: clsParagraph Class - a class demonstrating a paragraph
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private lStart As Long
Private lLength As Long
Private Item() As cpsSentence
Private iItem As Long
'---------------------doan-----------------
Public Function Copy(ByVal vPar As cpsParagraph) As Integer
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
Public Function SentenceSplit(ByVal TB As MSForms.TextBox) As Integer
On Error GoTo Result
Dim ls As Long, ll As Long
Dim vSen As New cpsSentence
    'Thuc hien tach cau
    SentenceSplit = 0
    ll = lStart + lLength
    ls = lStart
    TB.SelStart = ls
    TB.SelLength = 1
    ls = TB.SelStart
    Do While TB.SelStart < ll
        If IsEndSentence(TB.SelText) Then
            If ls < TB.SelStart Then
                iItem = iItem + 1
                ReDim Preserve Item(iItem)
                Set Item(iItem) = New cpsSentence
                vSen.SentenceStart = ls
                vSen.SentenceLength = TB.SelStart - ls
                vSen.SentenceSign = TB.SelText
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
        vSen.SentenceStart = ls
        vSen.SentenceLength = TB.SelStart - ls
        vSen.SentenceSign = ""
        If Item(iItem).Copy(vSen) <> 0 Then GoTo Result
    End If
    Exit Function
Result:
    SentenceSplit = 1000
End Function

Public Property Get SentenceCount() As Long
    SentenceCount = iItem
End Property

Public Property Get Sentence(ByVal i As Long) As cpsSentence
    Set Sentence = Item(i)
End Property

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
'---------------am tiet-----------------
Public Function SyllSplit(ByVal i As Long, ByVal TB As MSForms.TextBox) As Integer
    SyllSplit = Item(i).SyllSplit(TB)
End Function

Public Property Get SyllCount(ByVal i As Long) As Long
    SyllCount = Item(i).SyllCount
End Property

Public Property Get Syllable(ByVal i As Long, ByVal j As Long) As cpsWord
    Set Syllable = Item(i).Syllable(j)
End Property
'---------------tu----------------------
Public Function AWSplit(ByVal i As Long, ByVal TB As MSForms.TextBox) As Integer
    AWSplit = Item(i).AWSplit(TB)
End Function

Public Property Get AWCount(ByVal i As Long) As Long
    AWCount = Item(i).AWCount
End Property

Public Property Get AW(ByVal i As Long, ByVal j As Long, ByVal k As Long) As cpsWord
    Set AW = Item(i).AW(j, k)
End Property
'---------------gop tu--------------
Public Function MakeRS(ByVal i As Long, ByVal TB As MSForms.TextBox) As Integer
    MakeRS = Item(i).GenerateRS(TB)
End Function

Public Property Get RSCount(ByVal i As Long) As Long
    RSCount = Item(i).RSCount
End Property

Public Property Get RSSCount(ByVal i As Long, ByVal j As Long) As Long
    RSSCount = Item(i).RSSCount(j)
End Property

Public Property Get RS(ByVal i As Long, ByVal j As Long, ByVal k As Long) As cpsRSItem
    Set RS = Item(i).RS(j, k)
End Property
