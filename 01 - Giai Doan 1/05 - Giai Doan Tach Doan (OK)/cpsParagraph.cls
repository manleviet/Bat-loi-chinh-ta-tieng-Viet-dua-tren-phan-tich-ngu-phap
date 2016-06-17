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
'Private Item() As cpsSentence
'Private iItem As Long

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

'Public Function SentenceSplit() As Integer
'On Error GoTo Result
'Dim ls As Long
'Dim vSen As New cpsSentence
'    'Thuc hien tach cau
'    SentenceSplit = 0
'    TB.SelStart = 0
'    TB.SelLength = 1
'    ls = TB.SelStart
'    Do While TB.SelLength = 1
'        If IsEndSentence(TB.SelText) Then
'            If ls < TB.SelStart Then
'                iItem = iItem + 1
'                ReDim Preserve Item(iItem)
'                Set Item(iItem) = New cpsSentence
'                vSen.Start = ls
'                vSen.Length = TB.SelStart - ls
'                vSen.Sign = TB.SelText
'                If Item(iItem).Copy(vSen) <> 0 Then GoTo Result
'            End If
'            ls = TB.SelStart + 1
'        End If
'        TB.SelStart = TB.SelStart + 1
'        TB.SelLength = 1
'    Loop
'    If TB.SelStart > ls Then
 '       iItem = iItem + 1
'        ReDim Preserve Item(iItem)
'        Set Item(iItem) = New cpsSentence
'        vSen.Start = ls
'        vSen.Length = TB.SelStart - ls
'        vSen.Sign = ""
'        If Item(iItem).Copy(vSen) <> 0 Then GoTo Result
'    End If
'    Exit Function
'Result:
'    SentenceSplit = 1000
'End Function

'Public Property Get SentenceCount() As Long
'    SentenceCount = iItem
'End Property

'Public Property Get Sentence(ByVal i As Long) As cpsSentence
'    Set Sentence = Item(i)
'End Property

'Private Function IsEndSentence(ByVal st As String) As Boolean
'    IsEndSentence = False
'    If st <> "" Then
'        Select Case AscW(st)
'            Case 46: IsEndSentence = True 'dau cham
'            Case 63: IsEndSentence = True 'dau cham hoi
'            Case 33: IsEndSentence = True 'dau cham cam
'            Case 13: IsEndSentence = True 'ky tu xuong dong
'        End Select
'    End If
'End Function
