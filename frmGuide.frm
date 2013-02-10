VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGuide 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guide for GameMaker 5.x GMD Recovery"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameStep5 
      Caption         =   "Step 5"
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdExtractGmd 
         Caption         =   "Extract GMD "
         Height          =   615
         Left            =   960
         TabIndex        =   23
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Final Step now we need to find where the gmd starts.  This is will now make your file you need!"
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameStep4 
      Caption         =   "Step 4"
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdDecryptExe 
         Caption         =   "Decrypt Exe"
         Height          =   975
         Left            =   960
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Decrypting the exe.  This is going to take a while. Let it do its job.  When its done it will say Done. Then press Next."
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameStep3 
      Caption         =   "Step 3"
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdOverrideKey 
         Caption         =   "Override Key"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdFindKey 
         Caption         =   "Find Key"
         Height          =   615
         Left            =   1200
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Advanced option only!  Use this only if you know how.>>>"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   $"frmGuide.frx":0000
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame FrameStep2 
      Caption         =   "Step 2"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdGetEncryptKey 
         Caption         =   "Get Memory Dump"
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   $"frmGuide.frx":00C9
         Height          =   1335
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblGmVersion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame FrameStep1 
      Caption         =   "Step 1"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdOneStep 
         Caption         =   "One Step Decompile (Beta)"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   1680
         Width           =   2555
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Select your exe! The game will run. DO NOT CLOSE IT!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   $"frmGuide.frx":0206
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label lblVersion 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "frmGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Number As Integer
Dim MaxSteps As Integer

Private Sub cmdBrowse_Click()
    CommonDialog1.DialogTitle = "Select Exe"
    CommonDialog1.Filename = ""
    CommonDialog1.Filter = "GameMaker Exe Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    
    txtPath.Text = CommonDialog1.Filename
    
    'check version
    Call frmMain.VersionCheck(CommonDialog1.Filename, False)
    lblGmVersion.Caption = "Made in GameMaker: " & GameVersion

On Error Resume Next
    If GameVersion = "5.0" Then
    ChDir (Left(txtPath.Text, Len(txtPath.Text) - Len(CommonDialog1.FileTitle)))
    Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "5.1" Then

         Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "5.2" Then

         Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "5.3Beta" Then
         Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "5.3" Then
         Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "5.3a" Then
        Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "5.4" Then
         Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "6.0" Then
        Shell GetShortPathName(txtPath.Text)
    ElseIf GameVersion = "4.3" Then

        Shell GetShortPathName(txtPath.Text)
    Else
         MsgBox "Not a GM exe", vbExclamation
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDecryptExe_Click()
    Call frmMain.VersionCheck(txtPath.Text, True)
End Sub

Private Sub cmdExtractGmd_Click()
    Call frmMain.FindGMDStartOffset
End Sub

Private Sub cmdFindKey_Click()
    If GameVersion = "5.0" Then
        frmMain.Key50Find
    ElseIf GameVersion = "5.1" Then
        frmMain.Key51FIND
    ElseIf GameVersion = "5.2" Then
        frmMain.Key52Find
    ElseIf GameVersion = "5.3Beta" Then
        frmMain.Key53BetaFind
    ElseIf GameVersion = "5.3" Then
        frmMain.Key53Find
    ElseIf GameVersion = "5.3a" Then
        Call frmMain.Key53AFind
    ElseIf GameVersion = "5.4" Then
        Call frmMain.Key54Find
    ElseIf GameVersion = "6.0" Then
        Call frmMain.Key60Find
       ' MsgBox "6.0 Encryption bah still working on it"
    ElseIf GameVersion = "4.3" Then
        frmMain.Key43Find
    Else
        MsgBox "Not a GM exe"
       ' MsgBox "Not added for 4.3 or not a GM exe"
    End If
        
End Sub

Private Sub cmdGetEncryptKey_Click()
   ' frmSelectProcess.txtHigh.Text = "14594304"
   frmSelectProcess.txtHigh.Text = "8594304"
    frmSelectProcess.Show vbModal, Me
End Sub

Private Sub cmdNext_Click()
    If Number < MaxSteps Then
        Number = Number + 1
        cmdPrevious.Enabled = True
        Select Case Number
            
            Case 1
                FrameStep2.Visible = False
                FrameStep3.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = False
                FrameStep1.Visible = True
            Case 2
                FrameStep1.Visible = False
                FrameStep3.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = False
                FrameStep2.Visible = True
            Case 3
                FrameStep1.Visible = False
                FrameStep2.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = False
                FrameStep3.Visible = True
            Case 4
                FrameStep1.Visible = False
                FrameStep2.Visible = False
                FrameStep3.Visible = False
                FrameStep5.Visible = False
                FrameStep4.Visible = True
                
            Case 5
                FrameStep5.Visible = False
                FrameStep1.Visible = False
                FrameStep2.Visible = False
                FrameStep3.Visible = False
                FrameStep5.Visible = True
        End Select
        
    Else
        cmdPrevious.Enabled = True
        cmdNext.Enabled = False
    End If
End Sub

Private Sub cmdOneStep_Click()
    Dim Response As String
    Response = MsgBox("This is in beta. I still suggest you use the guide! Do you want to continue?", vbYesNo + vbInformation)
    If Response = vbYes Then
        CommonDialog1.Filename = ""
        CommonDialog1.DialogTitle = "Select Exe"
        CommonDialog1.Filter = "GameMaker Exe Files (*.exe)|*.exe"
        CommonDialog1.ShowOpen
    
        If CommonDialog1.Filename = "" Then Exit Sub

        Call frmMain.OneStepDecompile(CommonDialog1.Filename)
    End If
End Sub

Private Sub cmdOverrideKey_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub cmdPrevious_Click()
    If Number > 1 Then
        Number = Number - 1
        cmdNext.Enabled = True
        Select Case Number
            
            Case 1
                FrameStep2.Visible = False
                FrameStep3.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = False
                FrameStep1.Visible = True
            Case 2
                FrameStep1.Visible = False
                FrameStep3.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = False
                FrameStep2.Visible = True
            Case 3
                FrameStep1.Visible = False
                FrameStep2.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = False
                FrameStep3.Visible = True
            Case 4
                FrameStep1.Visible = False
                FrameStep2.Visible = False
                FrameStep3.Visible = False
                FrameStep5.Visible = False
                FrameStep4.Visible = True
            Case 5
                FrameStep1.Visible = False
                FrameStep2.Visible = False
                FrameStep3.Visible = False
                FrameStep4.Visible = False
                FrameStep5.Visible = True
        End Select
        
    Else
        cmdPrevious.Enabled = False
        cmdNext.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Ver: " & Version
    cmdPrevious.Enabled = False
    Number = 1
    MaxSteps = 5
End Sub



