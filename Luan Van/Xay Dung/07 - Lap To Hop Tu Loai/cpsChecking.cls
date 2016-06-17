VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpsChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Project: Vietnamese Checking
'Description: cpsChecking Class
'--------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Private TB As MSForms.TextBox
Private Item() As cpsParagraph
Private iItem As Long
'--------------khoi tao-----------------
Private Sub Class_Initialize()
    iItem = 0
    ReDim Item(iItem)
End Sub

Public Sub Init(TTextBox)
    Set TB = TTextBox
End Sub

Public Function Clear() As Integer
On Error GoTo Result
    Clear = 0
    iItem = 0
    ReDim Item(iItem)
Result:
    Clear = 1000
End Function
'---------------doan--------------------
Private Function IsEndParagraph(ByVal tu As String) As Boolean
    IsEndParagraph = False
    If tu <> "" Then
        Select Case AscW(tu)
            Case 13: IsEndParagraph = True
        End Select
    End If
End Function

Public Function ParagraphSplit() As Integer
On Error GoTo Result
Dim ls As Long
Dim vPar As New cpsParagraph
    ParagraphSplit = 0
    TB.SelStart = 0
    TB.SelLength = 1
    ls = TB.SelStart
    Do While TB.SelLength = 1
        If IsEndParagraph(TB.SelText) Then
            If ls < TB.SelStart Then
                iItem = iItem + 1
                ReDim Preserve Item(iItem)
                Set Item(iItem) = New cpsParagraph
                vPar.ParagraphStart = ls
                vPar.ParagraphLength = TB.SelStart - ls
                If Item(iItem).Copy(vPar) <> 0 Then GoTo Result
            End If
            ls = TB.SelStart + 1
        End If
        TB.SelStart = TB.SelStart + 1
        TB.SelLength = 1
    Loop
    If TB.SelStart > ls Then
        iItem = iItem + 1
        ReDim Preserve Item(iItem)
        Set Item(iItem) = New cpsParagraph
        vPar.ParagraphStart = ls
        vPar.ParagraphLength = TB.SelStart - ls
        If Item(iItem).Copy(vPar) <> 0 Then GoTo Result
    End If
    Exit Function
Result:
    ParagraphSplit = 1000
End Function

Public Property Get ParagraphCount() As Long
    ParagraphCount = iItem
End Property

Public Property Get Paragraph(ByVal i As Long) As cpsParagraph
    Set Paragraph = Item(i)
End Property
'----------------cau----------------------
Public Function SentenceSplit(ByVal i As Long) As Integer
    SentenceSplit = Item(i).SentenceSplit(TB)
End Function

Public Property Get SentenceCount(ByVal i As Long) As Long
    SentenceCount = Item(i).SentenceCount
End Property

Public Property Get Sentence(ByVal i As Long, ByVal j As Long) As cpsSentence
    Set Sentence = Item(i).Sentence(j)
End Property
'---------------am tiet-----------------
Public Function SyllSplit(ByVal i As Long, ByVal j As Long) As Integer
    SyllSplit = Item(i).SyllSplit(j, TB)
End Function

Public Property Get SyllCount(ByVal i As Long, ByVal j As Long) As Long
    SyllCount = Item(i).SyllCount(j)
End Property

Public Property Get Syllable(ByVal i As Long, ByVal j As Long, ByVal k As Long) As cpsWord
    Set Syllable = Item(i).Syllable(j, k)
End Property
'----------------tu----------------------
Public Function AWSplit(ByVal i As Long, ByVal j As Long) As Integer
    AWSplit = Item(i).AWSplit(j, TB)
End Function

Public Property Get AWCount(ByVal i As Long, ByVal j As Long) As Long
    AWCount = Item(i).AWCount(j)
End Property

Public Property Get AW(ByVal i As Long, ByVal j As Long, ByVal k As Long, ByVal t As Long) As cpsWord
    Set AW = Item(i).AW(j, k, t)
End Property
'---------------go tu---------------------
Public Function MakeRS(ByVal i As Long, ByVal j As Long) As Integer
    MakeRS = Item(i).MakeRS(j, TB)
End Function

Public Property Get RSCount(ByVal i As Long, ByVal j As Long) As Long
    RSCount = Item(i).RSCount(j)
End Property

Public Property Get RSSCount(ByVal i As Long, ByVal j As Long, ByVal k As Long) As Long
    RSSCount = Item(i).RSSCount(j, k)
End Property

Public Property Get RS(ByVal i As Long, ByVal j As Long, ByVal k As Long, ByVal t As Long) As cpsRSItem
    Set RS = Item(i).RS(j, k, t)
End Property
'------------------to hop tu loai-----------------
Public Function MakeWCS(ByVal i As Long, ByVal j As Long) As Integer
    MakeWCS = Item(i).MakeWCS(j)
End Function

Public Property Get WCS(ByVal i As Long, ByVal j As Long, ByVal k As Long) As String
    WCS = Item(i).WCS(j, k)
End Property

Public Property Get WCSCount(ByVal i As Long, ByVal j As Long) As Long
    WCSCount = Item(i).WCSCount(j)
End Property

Public Property Get WCSRSItem(ByVal i As Long, ByVal j As Long, ByVal k As Long) As Long
    WCSRSItem = Item(i).WCSRSItem(j, k)
End Property
