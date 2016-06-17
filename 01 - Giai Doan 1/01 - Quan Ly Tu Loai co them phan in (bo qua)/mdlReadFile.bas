Attribute VB_Name = "mdlReadFile"
Option Explicit

Public Function ReadTextFile(FileName) As String
   Dim Fs As FileSystemObject
   Dim TS As TextStream
   '  Create a FileSystem Object
   Set Fs = CreateObject("Scripting.FileSystemObject")
   ' Open TextStream for Input
   Set TS = Fs.OpenTextFile(FileName, ForReading, False, TristateUseDefault)
   ReadTextFile = TS.ReadAll  ' Read the whole content of the text file in one stroke
   TS.Close ' Close the Text Stream
   Set Fs = Nothing  ' Dispose FileSystem Object
End Function
