VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help for Game Maker 5.x GMD Recovery"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Tag             =   "                                        vbgamer45"
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "The encryption key is stored in the exe right above the gmd begin header."
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   $"frmHelp.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Use the guide and read the Read Me.txt if you are having troubles. I have tried to make this as easy as possible."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Bmp files are dumped in 5.0 in the temp directory. Just a couple of bytes added to the front of them."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmHelp.frx":0094
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
