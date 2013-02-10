VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameOverride 
      Caption         =   "Override Encryption key?"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
      Begin VB.CommandButton cmdBrowseNormal 
         Caption         =   "Browse"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkOverRideCharSet 
         Caption         =   "Override CharSet Normal"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdNormalKey 
         Caption         =   "Generate Normal"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chkOverrideKey 
         Caption         =   "Override Encryption Key?"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Encryption key needs to be a text file that is 256 bytes long. And needs to contain all characters from hex of 00 to FF."
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CheckBox chkShowWavoffset 
      Caption         =   "Show WAV offsets"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Shows the location of wav sounds inside a file."
      Top             =   240
      Width           =   1695
   End
   Begin VB.CheckBox chkShowOffsetBmp 
      Caption         =   "Show BMP offsets"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Shows the location of a bmp image inside a file"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################################
'frmOptions
'vbgamer45
'################################################
Option Explicit
Private Sub chkOverRideCharSet_Click()
    If chkOverRideCharSet.Value = vbChecked Then
        OverRideNormalKey = True
    Else
        OverRideNormalKey = False
    End If
End Sub

Private Sub chkOverrideKey_Click()
    If chkOverrideKey.Value = vbChecked Then
        OverRideEncryptKey = True
    Else
        OverRideEncryptKey = False
    End If
End Sub

Private Sub chkShowOffsetBmp_Click()
    If chkShowOffsetBmp.Value = vbChecked Then
        ShowBmpOffsets = True
    Else
        ShowBmpOffsets = False
    End If
End Sub

Private Sub chkShowWavoffset_Click()
    If chkShowWavoffset.Value = vbChecked Then
        ShowWavoffsets = True
    Else
        ShowWavoffsets = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog1.InitDir = App.Path
    CommonDialog1.DialogTitle = "Select Key File"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    KeyFileName = CommonDialog1.Filename
    txtKey.Text = KeyFileName
    frmMain.LoadKeyByFilename (CommonDialog1.Filename)
    
End Sub

Private Sub cmdBrowseNormal_Click()
    CommonDialog1.InitDir = App.Path
    CommonDialog1.DialogTitle = "Select Normal File"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    NormalFileName = CommonDialog1.Filename
    'txtKey.Text = KeyFileName
    frmMain.LoadNormalByFilename (CommonDialog1.Filename)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNormalKey_Click()
    Dim t As String
    Dim i As Long
    Dim f As Long
        For i = 0 To 255
            t = t & Chr(i)
        Next i
     f = FreeFile
     Open App.Path & "\NormalValues.txt" For Output As #f
        Print #f, t
     Close #f
     MsgBox "An example key generated. Shows all values edit it to your needs.  Check NormalValues.txt"
End Sub

Private Sub Form_Load()
    If ShowBmpOffsets = True Then
        chkShowOffsetBmp.Value = vbChecked
    Else
        chkShowOffsetBmp.Value = vbUnchecked
    End If
    If ShowWavoffsets = True Then
        chkShowWavoffset.Value = vbChecked
    Else
        chkShowWavoffset.Value = vbUnchecked
    End If
    If OverRideEncryptKey = True Then
        chkOverrideKey.Value = vbChecked
    Else
        chkOverrideKey.Value = vbUnchecked
    End If
    If OverRideNormalKey = True Then
        chkOverRideCharSet.Value = vbChecked
    Else
        chkOverRideCharSet.Value = vbUnchecked
    End If
    txtKey.Text = KeyFileName
End Sub

Private Sub txtKey_Change()
    KeyFileName = txtKey.Text
End Sub
