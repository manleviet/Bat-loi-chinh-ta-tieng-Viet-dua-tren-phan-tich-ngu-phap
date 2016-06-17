Attribute VB_Name = "mdlEnum"
'Project: Vietnamese Checking
'Description: mdlEnum Modul - Enums Declaration
'------------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit

Public Enum eError
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
End Enum

Public Enum eFILETYPE
    IPage = 0
    Word = 1
    WClass = 2
    Rule = 3
End Enum

Public Enum eGoiY
    BoQua = 0
    BoQuaHet = 1
    ThayThe = 2
    ThayTheHet = 3
    Dung = 4
End Enum

Public Enum eSpecialCharacter
    None = 0
    DauPhay = 1
    MoNgoac = 2
    DongNgoac = 3
End Enum
