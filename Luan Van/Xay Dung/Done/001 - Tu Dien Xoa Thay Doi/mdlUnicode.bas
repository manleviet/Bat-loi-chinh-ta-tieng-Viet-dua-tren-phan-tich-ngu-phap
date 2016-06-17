Attribute VB_Name = "mdlUnicode"
'Project: Vietnamese Checking
'Description: mdlUnicode Modul - a Modul processing Unicode
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public UVowels As String

Sub InitUnicode()
Dim TStr As String
    TStr = TStr & ChrW(&HE1) & ChrW(&HE0) & ChrW(&H1EA3) & ChrW(&HE3) & ChrW(&H1EA1) & ChrW(&H103) & ChrW(&H1EAF) & ChrW(&H1EB1) & ChrW(&H1EB3) & ChrW(&H1EB5) & ChrW(&H1EB7) & ChrW(&HE2) & ChrW(&H1EA5) & ChrW(&H1EA7) & ChrW(&H1EA9) & ChrW(&H1EAB) & ChrW(&H1EAD) & ChrW(&HE9) & ChrW(&HE8) & ChrW(&H1EBB)
    TStr = TStr & ChrW(&H1EBD) & ChrW(&H1EB9) & ChrW(&HEA) & ChrW(&H1EBF) & ChrW(&H1EC1) & ChrW(&H1EC3) & ChrW(&H1EC5) & ChrW(&H1EC7) & ChrW(&HED) & ChrW(&HEC) & ChrW(&H1EC9) & ChrW(&H129) & ChrW(&H1ECB) & ChrW(&HF3) & ChrW(&HF2) & ChrW(&H1ECF) & ChrW(&HF5) & ChrW(&H1ECD) & ChrW(&HF4) & ChrW(&H1ED1)
    TStr = TStr & ChrW(&H1ED3) & ChrW(&H1ED5) & ChrW(&H1ED7) & ChrW(&H1ED9) & ChrW(&H1A1) & ChrW(&H1EDB) & ChrW(&H1EDD) & ChrW(&H1EDF) & ChrW(&H1EE1) & ChrW(&H1EE3) & ChrW(&HFA) & ChrW(&HF9) & ChrW(&H1EE7) & ChrW(&H169) & ChrW(&H1EE5) & ChrW(&H1B0) & ChrW(&H1EE9) & ChrW(&H1EEB) & ChrW(&H1EED) & ChrW(&H1EEF)
    TStr = TStr & ChrW(&H1EF1) & ChrW(&HFD) & ChrW(&H1EF3) & ChrW(&H1EF7) & ChrW(&H1EF9) & ChrW(&H1EF5) & ChrW(&H111) & ChrW(&HC1) & ChrW(&HC0) & ChrW(&H1EA2) & ChrW(&HC3) & ChrW(&H1EA0) & ChrW(&H102) & ChrW(&H1EAE) & ChrW(&H1EB0) & ChrW(&H1EB2) & ChrW(&H1EB4) & ChrW(&H1EB6) & ChrW(&HC2) & ChrW(&H1EA4)
    TStr = TStr & ChrW(&H1EA6) & ChrW(&H1EA8) & ChrW(&H1EAA) & ChrW(&H1EAC) & ChrW(&HC9) & ChrW(&HC8) & ChrW(&H1EBA) & ChrW(&H1EBC) & ChrW(&H1EB8) & ChrW(&HCA) & ChrW(&H1EBE) & ChrW(&H1EC0) & ChrW(&H1EC2) & ChrW(&H1EC4) & ChrW(&H1EC6) & ChrW(&HCD) & ChrW(&HCC) & ChrW(&H1EC8) & ChrW(&H128) & ChrW(&H1ECA)
    TStr = TStr & ChrW(&HD3) & ChrW(&HD2) & ChrW(&H1ECE) & ChrW(&HD5) & ChrW(&H1ECC) & ChrW(&HD4) & ChrW(&H1ED0) & ChrW(&H1ED2) & ChrW(&H1ED4) & ChrW(&H1ED6) & ChrW(&H1ED8) & ChrW(&H1A0) & ChrW(&H1EDA) & ChrW(&H1EDC) & ChrW(&H1EDE) & ChrW(&H1EE0) & ChrW(&H1EE2) & ChrW(&HDA) & ChrW(&HD9) & ChrW(&H1EE6)
    TStr = TStr & ChrW(&H168) & ChrW(&H1EE4) & ChrW(&H1AF) & ChrW(&H1EE8) & ChrW(&H1EEA) & ChrW(&H1EEC) & ChrW(&H1EEE) & ChrW(&H1EF0) & ChrW(&HDD) & ChrW(&H1EF2) & ChrW(&H1EF6) & ChrW(&H1EF8) & ChrW(&H1EF4) & ChrW(&H110)
    UVowels = TStr
End Sub

Function LowerUniChar(ByVal Ch As String) As String
Dim Pos As Integer
    Pos = InStr(UVowels, Ch)
    If Pos > 67 Then
      LowerUniChar = Mid(UVowels, Pos - 67, 1)
    ElseIf Pos > 0 Then
      LowerUniChar = Ch
    Else
      LowerUniChar = LCase(Ch)
    End If
End Function

Function IsUpperUniChar(ByVal Ch As String) As Boolean
   IsUpperUniChar = (InStr(UVowels, Ch) > 67)
End Function

Public Function IsUpperChar(ByVal st As String) As Boolean
    Select Case AscW(st)
        Case 65 To 90: IsUpperChar = True
        Case Else: IsUpperChar = False
    End Select
End Function

Public Function UniLCase(ByVal st As String) As String
Dim thu As String, tam As String
Dim i As Long
    thu = ""
    For i = 1 To Len(st)
        tam = Mid(st, i, 1)
        If IsUpperChar(tam) Then tam = LCase(tam)
        If IsUpperUniChar(tam) Then tam = LowerUniChar(tam)
        thu = thu & tam
    Next i
    UniLCase = thu
End Function

Function UniStrToUTF16(UniString) As Integer()
Dim BArray() As Integer
Dim i As Long
Dim TLen As Long
Dim b1 As Byte
Dim b2 As Byte
Dim UTF16 As Long
   TLen = Len(UniString) 'Lay do dai chuoi
   ReDim BArray(TLen)
   If TLen = 0 Then
      UniStrToUTF16 = BArray 'Tra ra mot mang UTF-16
      Exit Function 'Neu chuoi rong thi khong lam
   End If
   For i = 1 To TLen
      ' Lay gia tri UTF16 tu ky tu Unicode
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Ket hop hai byte thanh Unicode UTF-16
      UTF16 = b2 'gan b2 vao UTF16 truoc khi nhan voi 256 de tranh tran
      UTF16 = UTF16 * 256 + b1
      
      BArray(i) = UTF16 'Dua gia tri vao mang
   Next
   UniStrToUTF16 = BArray 'Tra ra mot mang UTF-16
End Function

Function UTF16ToUniStr(BArray) As String
Dim i As Long
Dim TopIndex As Long
Dim TwoBytes(1) As Byte
Dim TStr As String
   TopIndex = UBound(BArray)
   If TopIndex = 0 Then
     UTF16ToUniStr = ""
     Exit Function
   End If
   For i = 1 To TopIndex
     TStr = TStr & ChrW(BArray(i))
   Next i
   UTF16ToUniStr = TStr
End Function
