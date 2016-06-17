Attribute VB_Name = "mdlType"
'Project: Vietnamese Checking
'Description: mdlType Modul - Types Declaration
'----------------------------
'Author: Le Viet Man
'   University of Hue
'   College of Sciences - IT Department

Option Explicit
'cau truc File Header
Public Type tFILEHEADER
    iType As Byte
    iSize As Long
End Type

Const NTer = "S"
Const Ter = "ABCD"
