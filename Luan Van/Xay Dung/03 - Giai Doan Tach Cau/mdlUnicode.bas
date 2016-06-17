Attribute VB_Name = "mdlUnicode"
'Project: Vietnamese Checking
'Description: mdlUnicode Modul - a Modul processing Unicode
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Public UVowels As String
' API to access VB6 String by pointer in order to copy memory
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Enum coEncoding
  coANSI = 0
  coUnicode = 1
  coUTF8 = 2
End Enum

Function IsUniChar(Ch) As Boolean
' Return True if Ch is a Unicode Vowel or dd, DD
   IsUniChar = (InStr(UVowels, Ch) > 0)
End Function


Function UpperUniChar(Ch) As String
' Return the Uppercase for a given vowel or dd
Dim Pos ' Position of character in Unicode vowel list
    ' Locate the character in list of Unicode  vowels
    Pos = InStr(UVowels, Ch)
    If (Pos > 67) Then
      UpperUniChar = Ch ' It's already uppercase - leave it alone
    ElseIf (Pos > 0) Then
      ' It's a Lowercase Unicode Vowel - so get the corresponding Uppercase vowel in the list
      UpperUniChar = Mid(UVowels, Pos + 67, 1)
    Else
      ' It's just a normal ANSI character
      UpperUniChar = UCase(Ch)
    End If
End Function

Function UpperUniStr(IPString) As String
' Convert a Unicode string to UpperCase
Dim i, TLen, TStr
  TStr = ""  ' Initialise the resultant string
  TLen = Len(IPString) ' get length of input Unicode string
  If TLen > 0 Then
    ' Iterate through each character of the Unicode string
    For i = 1 To TLen
      ' Convert each character to uppercase
      TStr = TStr & UpperUniChar(Mid(IPString, i, 1))
    Next
  End If
  UpperUniStr = TStr ' Return the resultant string
End Function
Function LowerUniStr(IPString) As String
' Convert a Unicode string to LowerCase
Dim i, TLen, TStr
  TStr = ""  ' Initialise the resultant string
  TLen = Len(IPString) ' get length of input Unicode string
  If TLen > 0 Then
    ' Iterate through each character of the Unicode string
    For i = 1 To TLen
      ' Convert each character to lowercase
      TStr = TStr & LowerUniChar(Mid(IPString, i, 1))
    Next
  End If
  LowerUniStr = TStr ' Return the resultant string
End Function

Function ToUTF8(ByVal UTF16 As Long) As Byte()
' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
Dim BArray() As Byte
  If UTF16 < &H80 Then
    ReDim BArray(0) ' one byte UTF-8
    BArray(0) = UTF16 ' Use number as is
  ElseIf UTF16 < &H800 Then
    ReDim BArray(1) ' two byte UTF-8
    BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
    UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
    BArray(0) = &HC0 + (UTF16 And &H1F) ' Use 5 remaining bits
  Else
    ReDim BArray(2) ' three byte UTF-8
    BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
    UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
    BArray(1) = &H80 + (UTF16 And &H3F) ' Use next 6 bits
    UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
    BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
  End If
  ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
Function ToUTF16(BArray) As Long
' Convert 2 or 3 UTF-8 bytes to a 16bit UTF-16BE
Dim IntUB
  IntUB = UBound(BArray) ' Find out how many bytes UTF-8 takes
  Select Case IntUB
  Case 0  ' one byte UTF-8.  Note that bArray starts with index=0
    ToUTF16 = BArray(0)  ' Use number as is
  Case 1  ' two byte UTF-8
    ToUTF16 = (BArray(0) And &H1F) * &H40 + (BArray(1) And &H3F)
  Case 2   ' three byte UTF-8
    ToUTF16 = (BArray(0) And &HF) * &H1000 + (BArray(1) And &H3F) * &H40 + (BArray(2) And &H3F)
  End Select
  
End Function
Function UniStrToUTF8(UniString) As Byte()
' Convert a Unicode string to a byte stream of UTF-8
Dim BArray() As Byte
Dim TempB() As Byte
Dim i As Long
Dim k As Long
Dim TLen As Long
Dim b1 As Byte
Dim b2 As Byte
Dim UTF16 As Long
Dim j
   TLen = Len(UniString) ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
    ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2 ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
        ReDim Preserve BArray(k)
        BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray ' Return the resultant UTF-8 byte array

End Function
Function UTF8ToUniStr(BArray) As String
' Convert a byte stream of UTF-8 to Unicode String
Dim i As Long
Dim TopIndex As Long
Dim TwoBytes(1) As Byte
Dim ThreeBytes(2) As Byte
Dim AByte As Byte
Dim TStr As String
   TopIndex = UBound(BArray)  ' Number of bytes equal TopIndex+1
   If TopIndex = 0 Then Exit Function ' get out if there's nothing to convert
   i = 0 ' Initialise pointer
   ' Iterate through the Byte Array
   Do While i <= TopIndex
     AByte = BArray(i) ' fetch a byte
     If AByte = &HE1 Then
        ' Start of 3 byte UTF-8 group for a character
        ' Copy 3 byte to ThreeBytes
        ThreeBytes(0) = BArray(i): i = i + 1
        ThreeBytes(1) = BArray(i): i = i + 1
        ThreeBytes(2) = BArray(i): i = i + 1
        ' Convert Byte array to UTF-16 then Unicode
        TStr = TStr & ChrW(ToUTF16(ThreeBytes))
     ElseIf (AByte >= &HC3) And (AByte <= &HC6) Then
        ' Start of 2 byte UTF-8 group for a character
        TwoBytes(0) = BArray(i): i = i + 1
        TwoBytes(1) = BArray(i): i = i + 1
        ' Convert Byte array to UTF-16 then Unicode
        TStr = TStr & ChrW(ToUTF16(TwoBytes))
     Else
        ' Normal ANSI character - use it as is
        TStr = TStr & Chr(AByte): i = i + 1 ' Increment byte array index
     End If
   Loop
   UTF8ToUniStr = TStr  ' Return the resultant string
End Function
Function HexDisplayOfFile(TFileName) As String
' Display the content of a text file in Hex format like:
'                FF FE 54 00 B0 01 DB 1E 63 00
   Dim Text1, MyChar, FileNum
   FileNum = FreeFile ' Obtain a File handle from the OS
   Open TFileName For Binary As #FileNum ' Open given Text file as binary
   ' Read all characters in the file.
   Do While Not EOF(FileNum)
      MyChar = Input(1, #FileNum)   ' Read a character as raw binary
      If MyChar <> "" Then
     ' Convert byte to Hex like 0A, 6B etc..
         Text1 = Text1 & HexOf(Asc(MyChar)) & " "
      End If
   Loop
   Close #FileNum  ' Close file
   HexDisplayOfFile = Text1  ' Return the Hex display string
End Function
Function GetFileEncoding(TFileName) As coEncoding
' Return the type of Text file : UTF16LE, UTF-8 or ANSI
Dim b1, FileNum
On Error Resume Next  ' Ignore error
FileNum = FreeFile ' Obtain a File handle from the OS
Open TFileName For Binary As #FileNum  ' Open given Textfile  as Binary
' Read all characters in the file.
   b1 = Input(1, #FileNum)   ' Read the first character.
   If Asc(b1) = &HFF Then
       GetFileEncoding = coUnicode  ' UTF-16LE
   ElseIf Asc(b1) = &HEF Then
       GetFileEncoding = coUTF8    ' UTF-8
   Else
       GetFileEncoding = coANSI    ' Normal ANSI
   End If
Close #FileNum  ' Close the file
End Function
Function ToUniDecimal(UniString As String) As String
' Return the HTML equivalent string of a Unicode string
Dim i As Integer  ' Must declare as integer for CopyMemory to work
Dim TLen, TStr
Dim b1 As Byte
Dim b2 As Byte
Dim UTF16 As Long
  TLen = Len(UniString)  ' Get Length of input Unicode string
  If TLen = 0 Then Exit Function ' Get out if null string
  ' Iterate through each character in the string
  For i = 1 To TLen
    If IsUniChar(Mid(UniString, i, 1)) Then
    ' Cast the String character to 2 bytes
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
    ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2 ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to format &#99999;  for HTML
      TStr = TStr & "&#" & Trim(CStr(UTF16)) & ";"
    Else
      ' Get here if it;s an ANSI character
      TStr = TStr & Mid(UniString, i, 1)
    End If
  Next
  ToUniDecimal = TStr  ' Return the HTML string
End Function
Function HexOf(ByVal AscNum As Integer) As String
  ' Return the 2 character Hex string of  AscNum, prefix extra "0" if necessary
  Dim TStr
  If AscNum > 255 Then AscNum = AscNum Mod 256
  TStr = Hex(AscNum)  ' Convert to Hex
  If Len(TStr) = 1 Then
    ' Attach "0" on the left
    TStr = "0" & TStr
  End If
  HexOf = TStr ' Return the 2 character Hex string
End Function







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
