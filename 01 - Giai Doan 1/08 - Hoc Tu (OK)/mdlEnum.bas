Attribute VB_Name = "mdlEnum"
'Project: Vietnamese Checking
'Description: mdlEnum Modul - Enums Declaration
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

Public Enum EError
    NoError = 0
    SortError = 1
    AddWordError = 2
    SwapError = 3
    DelError = 4
    NoHaveWord = 5
    SetWordError = 6
    AddPageError = 7
    LoadDicError = 8
    SaveDicError = 9
    GetWordError = 10
    CopyCPageError = 11
    ClearCPageError = 12
    CopyWordError = 13 ' loi trong viec copy clsWord
    TheSameWord = 14
End Enum

Public Enum eFILETYPE
    ipage = 0
    Word = 1
    WClass = 2
    Rule = 3
End Enum

