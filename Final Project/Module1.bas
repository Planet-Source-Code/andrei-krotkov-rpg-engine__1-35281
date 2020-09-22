Attribute VB_Name = "modRPG"
 Option Explicit
Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" _
(ByVal Hdc As Long, ByVal x As Long, _
ByVal y As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public swa(999) As Integer
Public HP As Integer, MaxHP As Integer
Public MP As Integer, MaxMP As Integer
Public EHP As Integer, eMaxHP As Integer
Public Str As Integer, def As Integer, Estr As Integer, edef As Integer
Public Flevel As Integer, ilevel As Integer, hLevel As Integer
Public Damage As Integer, Weakness As Integer
Public Eanim As Anim
Public Canrun As Boolean
Public Bwin As Boolean
Public cash As Integer
Public Enum Anim
    
    Fire = 1
    Ice = 2
    Dark = 3

End Enum
'ITEMS:
Public strItems(999) As String
Public intItemsOwn(999) As Integer
