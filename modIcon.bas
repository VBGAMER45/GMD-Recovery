Attribute VB_Name = "modIcon"
'#############################################
' modIcon.bas
' vbgamer45
'##############################################
Option Explicit
'Used for drawing the icon in attach process
Public glLargeIcons() As Long
Public glSmallIcons() As Long
Public lIcons         As Long




Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16
Public Const DI_NORMAL = 3

Public Declare Function DrawIconEx Lib "User32" _
    (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long

Public Declare Function ExtractIconEx Lib "shell32.dll" _
    Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, _
    phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long





