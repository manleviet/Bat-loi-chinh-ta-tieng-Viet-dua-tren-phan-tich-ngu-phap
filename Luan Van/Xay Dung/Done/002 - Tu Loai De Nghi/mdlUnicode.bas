Attribute VB_Name = "mdlUnicode"
'Project: Vietnamese Checking
'Description: mdlUnicode Modul - a Modul processing Unicode
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

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
