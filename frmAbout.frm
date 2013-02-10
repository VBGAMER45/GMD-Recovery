VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "About GameMaker GMD Recovery"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrLight 
      Interval        =   250
      Left            =   840
      Top             =   720
   End
   Begin VB.Timer tmrIcon 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GameMaker Decompiler 5.x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   0
      Left            =   960
      Picture         =   "frmAbout.frx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   1
      Left            =   1440
      Picture         =   "frmAbout.frx":030A
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "frmAbout.frx":0614
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   3
      Left            =   2400
      Picture         =   "frmAbout.frx":091E
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   4
      Left            =   960
      Picture         =   "frmAbout.frx":0C28
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   7
      Left            =   2400
      Picture         =   "frmAbout.frx":0F32
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   6
      Left            =   1920
      Picture         =   "frmAbout.frx":123C
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFlame 
      Height          =   480
      Index           =   5
      Left            =   1440
      Picture         =   "frmAbout.frx":1546
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Type LightningBolt
    Center As Long
    Inner1 As Long
    Inner2 As Long
    Inner3 As Long
    Inner4 As Long
    Outer1 As Long
    Outer2 As Long
    Outer3 As Long
    Outer4 As Long
    
    Nicks As Long
    VelocityMin As Long
    VelocityMax As Long
    Height As Long
    
    SplitBolt As Boolean
    SameStart As Boolean
    ShowCloud As Boolean
End Type
Dim Bolt As LightningBolt
Dim IconNumber As Integer
Dim AniValue As Integer
Dim BackDown As Boolean
Dim tList(10) As String
Dim NextTitle As Integer
Private Sub Form_Load()
    tmrIcon.Interval = 100
    Bolt.Center = RGB(207, 207, 233)
    'Set Inner
    Bolt.Inner1 = RGB(176, 176, 207)
    Bolt.Inner2 = RGB(176, 176, 192)
    Bolt.Inner3 = RGB(160, 160, 176)
    Bolt.Inner4 = RGB(145, 145, 176)
    'Set Outer
    Bolt.Outer1 = RGB(90, 90, 110)
    Bolt.Outer2 = RGB(80, 80, 110)
    Bolt.Outer3 = RGB(65, 65, 80)
    Bolt.Outer4 = RGB(50, 50, 70)
    'Set Properties
    Bolt.Nicks = 40
    Bolt.VelocityMin = 3
    Bolt.VelocityMax = 15
    Bolt.Height = 590
    Bolt.SameStart = False
    Bolt.ShowCloud = True
    Bolt.SplitBolt = True
    BackDown = False
    AniValue = 0
    NextTitle = 0
    'Title list
    
    tList(0) = "GameMaker 5.x GMD Recovery"
    tList(1) = "By VisualBasicZone.com"
    tList(2) = "Decompiles for:"
    tList(3) = "5.0 5.1 5.2 5.3Beta 5.3 and 5.3a"
    tList(4) = "Contact at:"
    tList(5) = "gmdecompiler@yahoo.com"
    tList(6) = "Have fun!"
End Sub
Private Sub tmrIcon_Timer()

            Icon = imgFlame(Int(8 * Rnd)).Picture

            IconNumber = (IconNumber + 1) Mod 3
            
           If BackDown = False Then
                If AniValue < 253 Then
                    AniValue = AniValue + 2
                Else
                    BackDown = True
                End If
           Else
                If AniValue > 2 Then
                    AniValue = AniValue - 2
                Else
                    If NextTitle < 6 Then
                        NextTitle = NextTitle + 1
                    Else
                        NextTitle = 0
                    End If
                    lblTitle.Caption = tList(NextTitle)
                    BackDown = False
                End If
           
           End If
           
           lblTitle.ForeColor = RGB(AniValue, AniValue, AniValue)
           
           

End Sub

Public Function DrawBolt(Pic2 As Form)
Dim SM, i, LX, LY, SX, SY, LX2, LY2, SX2, SY2 As Integer
Dim Alter As Boolean
'Pic2.Cls
If Bolt.SameStart = True Then
    SM = Pic2.ScaleWidth / 2
Else
    SM = Int(Rnd * (Pic2.ScaleWidth / 2) + Pic2.ScaleWidth / 4)
End If
Pic2.ForeColor = Bolt.Center
SX = SM
If Bolt.ShowCloud = True Then
    SY = 10
Else
    SY = 0
    SetPixel Pic2.hdc, SM, 0, Bolt.Center
End If
i = Int(Rnd * 2)
If i = 0 Then
    Alter = False
    LX = SX + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
Else
    Alter = True
    LX = SX - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
End If
LY = Bolt.Height / Bolt.Nicks
For i = 0 To Bolt.Nicks
    Pic2.ForeColor = Bolt.Center
    Pic2.Line (SX, SY)-(LX, LY)
    Pic2.ForeColor = Bolt.Inner1
    Pic2.Line (SX - 1, SY)-((SX + LX) / 2 - 1, (SY + LY) / 2)
    Pic2.Line (SX + 1, SY)-((SX + LX) / 2 + 1, (SY + LY) / 2)
    Pic2.ForeColor = Bolt.Inner2
    Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX - 1, LY)
    Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX + 1, LY)
    Pic2.ForeColor = Bolt.Outer1
    Pic2.Line (SX - 2, SY)-((SX + LX) / 2 - 2, (SY + LY) / 2)
    Pic2.Line (SX + 2, SY)-((SX + LX) / 2 + 2, (SY + LY) / 2)
    Pic2.ForeColor = Bolt.Outer2
    Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX - 2, LY)
    Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX + 2, LY)
    If i >= Round((Bolt.Nicks / 2), 0) And Bolt.SplitBolt = True Then
        Pic2.ForeColor = Bolt.Center
        Pic2.Line (SX, SY)-(LX, LY)
        Pic2.ForeColor = Bolt.Inner1
        Pic2.Line (SX2 - 1, SY2)-((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)
        Pic2.Line (SX2 + 1, SY2)-((SX2 + LX2) / 2 + 1, (SY2 + LY2) / 2)
        Pic2.ForeColor = Bolt.Inner2
        Pic2.Line ((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)-(LX2 - 1, LY2)
        Pic2.Line ((SX2 + LX2) / 2 - 1, (SY2 + LY2) / 2)-(LX2 + 1, LY2)
        Pic2.ForeColor = Bolt.Outer1
        Pic2.Line (SX2 - 2, SY2)-((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)
        Pic2.Line (SX2 + 2, SY2)-((SX2 + LX2) / 2 + 2, (SY2 + LY2) / 2)
        Pic2.ForeColor = Bolt.Outer2
        Pic2.Line ((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)-(LX2 - 2, LY2)
        Pic2.Line ((SX2 + LX2) / 2 - 2, (SY2 + LY2) / 2)-(LX2 + 2, LY2)
    End If
    If i = Bolt.Nicks Then
        'Do Tail
        If Alter = True Then
            '-
            SX = LX
            SY = LY
            LX = LX - 1
            LY = LY + 2
        Else
            '+
            SX = LX
            SY = LY
            LX = LX + 1
            LY = LY + 2
        End If
        Pic2.ForeColor = Bolt.Inner1
        Pic2.Line (SX, SY)-(LX, LY)
        Pic2.ForeColor = Bolt.Inner3
        Pic2.Line (SX - 1, SY)-((SX + LX) / 2 - 1, (SY + LY) / 2)
        Pic2.Line (SX + 1, SY)-((SX + LX) / 2 + 1, (SY + LY) / 2)
        Pic2.ForeColor = Bolt.Inner4
        Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX - 1, LY)
        Pic2.Line ((SX + LX) / 2 - 1, (SY + LY) / 2)-(LX + 1, LY)
        Pic2.ForeColor = Bolt.Outer3
        Pic2.Line (SX - 2, SY)-((SX + LX) / 2 - 2, (SY + LY) / 2)
        Pic2.Line (SX + 2, SY)-((SX + LX) / 2 + 2, (SY + LY) / 2)
        Pic2.ForeColor = Bolt.Outer4
        Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX - 2, LY)
        Pic2.Line ((SX + LX) / 2 - 2, (SY + LY) / 2)-(LX + 2, LY)
        GoTo RefreshPic
    End If
    SX = LX
    SY = LY
    Alter = Int(Rnd * 2)
    If Alter = True Then
        Alter = False
        LX = LX + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
    Else
        Alter = True
        LX = LX - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
    End If
    LY = LY + Bolt.Height / Bolt.Nicks
    If i < Round((Bolt.Nicks / 2), 0) - 1 And Bolt.SplitBolt = True Then
        LX2 = SX
        LY2 = SY
    End If
    If i >= Round((Bolt.Nicks / 2), 0) - 1 And Bolt.SplitBolt = True Then
        SX2 = LX2
        SY2 = LY2
        Alter = Int(Rnd * 2)
        If Alter = True Then
            Alter = False
            LX2 = LX2 + Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
        Else
            Alter = True
            LX2 = LX2 - Int(Rnd * (Bolt.VelocityMax + 1 - Bolt.VelocityMin) + Bolt.VelocityMin)
        End If
        LY2 = LY2 + Bolt.Height / Bolt.Nicks
    End If
Next i
RefreshPic:

Pic2.Refresh
End Function

Private Sub TmrLight_Timer()
    Dim i As Integer
        Call DrawBolt(Me)
        'Draw rain
        For i = 0 To 100
            Me.PSet (Int(Rnd * Me.ScaleWidth), Int(Rnd * Me.ScaleHeight)), vbBlue
        Next
End Sub
