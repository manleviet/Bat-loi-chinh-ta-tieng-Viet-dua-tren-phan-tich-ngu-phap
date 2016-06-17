Attribute VB_Name = "mdlInfoSystem"
'Project: Vietnamese Checking
'Description: mdlInfoSystem Modul - Informations System Declaration
'----------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
Public fIPage As String
Public fWordDic As String
Public fWClass As String

Public Sub GetInformation()
Dim FileName As String
    FileName = App.Path
    If Right(FileName, 1) <> "\" Then
        FileName = FileName & "\"
    End If
    
    fIPage = FileName & "Data\ipage.vch"
    fWordDic = FileName & "Data\worddic.vch"
    fWClass = FileName & "Data\wclass.vch"
End Sub
