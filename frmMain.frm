VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Game Maker 5.x GMD Recovery VisualBasicZone.com"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8475
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":27A2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtResults 
      Height          =   4815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileGuide 
         Caption         =   "&Show Guide"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileAttach 
         Caption         =   "&Attach to Exe"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenExe 
         Caption         =   "&Decrypt Exe"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileNoMemDecrypt 
         Caption         =   "Beta Decrypt no Memory"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileForceDecrypt 
         Caption         =   "&Force Decrypt"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileBruteForce 
         Caption         =   "&Brute Force Key Search"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileExtractGmd 
         Caption         =   "E&xtract Gmd from Game.enc"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileOneStep 
         Caption         =   "One Step Decompile (Beta)"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuExtractBmp 
         Caption         =   "Bmp Extractor"
      End
      Begin VB.Menu mnuWavExtract 
         Caption         =   "Wav Extractor"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGet43Key 
         Caption         =   "Get 4.3 Key"
      End
      Begin VB.Menu mnuGert50Key 
         Caption         =   "Get 5.0 Key"
      End
      Begin VB.Menu mnuGet51Key 
         Caption         =   "Get 5.1 Key"
      End
      Begin VB.Menu mnuGet52Key 
         Caption         =   "Get 5.2 Key"
      End
      Begin VB.Menu mnuGet53Beta 
         Caption         =   "Get 5.3 Key Beta Edtion"
      End
      Begin VB.Menu mnuGet53Key 
         Caption         =   "Get 5.3 Key"
      End
      Begin VB.Menu mnuGetKey53a 
         Caption         =   "Get 5.3a Key"
      End
      Begin VB.Menu mnuGet54Key 
         Caption         =   "Get 5.4 Key"
      End
      Begin VB.Menu mnuGet60Key 
         Caption         =   "Get 6.0 Key"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlugin 
         Caption         =   "Load Plugins"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuPluginList 
      Caption         =   "&Plugin List"
      Visible         =   0   'False
      Begin VB.Menu mnuPluginArray 
         Caption         =   "(none)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUse 
         Caption         =   "How to use"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###################################
'# Game Maker 5.x GMD Recovery
'# By: vbgamer45
'# September 27, 2004
'###################################
Dim Filesize As Long
Dim Unencrypted As String
Dim key As String * 256
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub MakeUnEncrypted()
    Dim t As String
    Dim i As Long

    For i = 0 To 255
        t = t & Chr(i)
    Next i
    Unencrypted = t
End Sub

Private Sub Form_Load()
    Dim R2 As String
    R2 = MsgBox("By running this program you agree to only use it on GameMaker exe files that you made! If not you are required quit this program now and delete it now!", vbYesNo + vbInformation, "GameMaker 5.x GMD Recovery")

    If R2 = vbNo Then End


    ShowBmpOffsets = False
    ShowWavoffsets = False
    'Set the Version
    Version = ".22"
    
   
    On Error Resume Next
    'Make the plugins directory
    MkDir (App.Path & "\Plugins\")
    'Print the read me
    Call modGlobals.PrintReadMe
    Me.Show
    
    If Command <> "" Then
        'Process cmd line arguements
        If Command = "about" Then
            frmAbout.Show vbModal, Me
        End If
        
        Exit Sub
    End If
    
    'Show the guide
    frmGuide.Show vbModal, Me
                                                                                        Me.Print "v b g a m e r 4  5"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    txtResults.Width = Me.Width - 100
    txtResults.Height = Me.Height - 700
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExtractBmp_Click()
'vbgamer45
    Dim larray() As Long
    'make the directories
    Call MakeDir
    ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    
    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find("BM", oldpos + 2)
        larray(UBound(larray)) = Pos

        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)

    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Extracting bmps " & UBound(larray) & vbCrLf
    
    
    If ShowBmpOffsets = True Then
        txtResults.Text = txtResults.Text & "Possible Bmp offsets" & vbCrLf
    End If
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            If ShowBmpOffsets = True Then
                txtResults.Text = txtResults.Text & larray(i) & vbCrLf
            End If
        GetBitmap CommonDialog1.Filename, (larray(i) + 1)
        End If
    Next
MsgBox "Done: Check " & App.Path & "\dump\images\", vbInformation
End Sub

Private Sub mnuFileAttach_Click()
    frmSelectProcess.Show vbModal, Me

End Sub

Private Sub mnuFileBruteForce_Click()
    frmBruteForce.Show
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileExtractGmd_Click()
     Call frmMain.FindGMDStartOffset
End Sub

Private Sub mnuFileForceDecrypt_Click()
On Error GoTo errHandle
    Dim k As String
    k = InputBox("Enter GameMaker Version(5.0 or 5.1 5.2 or 5.3):")
    If k = "" Then Exit Sub

    GameVersion = k
    CommonDialog1.Filename = ""
    CommonDialog1.DialogTitle = "Select Exe"
    CommonDialog1.Filter = "GameMaker Exe Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    
    Call Me.DecompileExe(CommonDialog1.Filename)
Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuFileForceDecrypt: " & Err.Number & " " & Err.Description
End Sub

Private Sub mnuFileGuide_Click()
    frmGuide.Show vbModal, Me
End Sub

Private Sub mnuFileNoMemDecrypt_Click()
    frmExeDecrypt.Show
End Sub

Private Sub mnuFileOneStep_Click()
    CommonDialog1.Filename = ""
    CommonDialog1.DialogTitle = "Select Exe"
    CommonDialog1.Filter = "GameMaker Exe Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    
    Call OneStepDecompile(CommonDialog1.Filename)
End Sub

Private Sub mnuGert50Key_Click()
    Call Key50Find

End Sub
Sub Key43Find()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    Dim key43 As String * 14
    Open App.Path & "\key43.txt" For Binary Access Read As #23
        Get #23, , key43
    Close #23


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key43, oldpos + Len(key43))
        larray(UBound(larray)) = Pos
       oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 4.3 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          modExtractor.Get43Key CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub

Sub Key50Find()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    Dim key50 As String * 14
    Open App.Path & "\key50.txt" For Binary Access Read As #23
        Get #23, , key50
    Close #23


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key50, oldpos + Len(key50))
        larray(UBound(larray)) = Pos
       oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.0 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          modExtractor.Get50Key CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub

Private Sub mnuGet43Key_Click()
  Call Key43Find
End Sub

Private Sub mnuGet51Key_Click()
    Call Key51FIND
End Sub
Sub Key51FIND()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
   ' Dim key51 As String * 6
   ' Open App.Path & "\key51.txt" For Binary Access Read As #23
    '    Get #23, , key51
   ' Close #23
    Dim key51 As String * 7
    Open App.Path & "\key51a.txt" For Binary Access Read As #23
        Get #23, , key51
    Close #23

    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key51, oldpos + Len(key51))
        larray(UBound(larray)) = Pos
       oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.1 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          Get51Key CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub

Private Sub mnuGet52Key_Click()
    Call Key52Find
End Sub
Sub Key52Find()
Dim larray() As Long, f As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    'Dim key52 As String * 6
    'Open App.Path & "\key52.txt" For Binary Access Read As #23
      '  Get #23, , key52
    'Close #23
    Dim key52 As String * 7
    f = FreeFile
    Open App.Path & "\key52a.txt" For Binary Access Read As #f
        Get #f, , key52
    Close #f

    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key52, oldpos + Len(key52))
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.2 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          Get52Key CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next

MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub

Private Sub mnuGet53Beta_Click()
    Call Key53BetaFind
End Sub
Sub Key53BetaFind()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    Dim key53 As String * 6
    Open App.Path & "\key53Beta.txt" For Binary Access Read As #23
        Get #23, , key53
    Close #23


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key53, oldpos + Len(key53))
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.3Beta Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          modExtractor.Get53BetaKey CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub
Sub Key53AFind()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    Dim key53a As String * 6
    Open App.Path & "\key53Beta.txt" For Binary Access Read As #23
        Get #23, , key53a
    Close #23


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key53a, oldpos + Len(key53a))
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.3A Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          modExtractor.Get53AKey CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub
Private Sub mnuGet53Key_Click()
    Call Key53Find
End Sub
Sub Key53Find()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    Dim key53 As String * 6
    Open App.Path & "\key53Beta.txt" For Binary Access Read As #23
        Get #23, , key53
    Close #23


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key53, oldpos + Len(key53))
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.3 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
         modExtractor.Get53BetaKey CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub

Private Sub mnuGet54Key_Click()
    Call Key54Find
End Sub
Sub Key54Find()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    Dim key54 As String * 7
    Open App.Path & "\key54.txt" For Binary Access Read As #23
        Get #23, , key54
    Close #23


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key54, oldpos + Len(key54))
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 5.4 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
         modExtractor.Get54Key CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub

Private Sub mnuGet60Key_Click()
    Call Key60Find
End Sub
Sub Key60Find()
Dim larray() As Long
Dim f As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    Close
    f = FreeFile
    Dim key60 As String * 7
    Open App.Path & "\key60.txt" For Binary Access Read As #f
        Get #f, , key60
    Close #f


    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(key60, oldpos + Len(key60))
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding 6.0 Key Location UBound:" & UBound(larray) & vbCrLf
'Dim KeyFound As Boolean
KeyFound = False
    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
          txtResults.Text = txtResults.Text & larray(i) & vbCrLf
         
          modExtractor.Get60Key CommonDialog1.Filename, (larray(i) + 1)
          KeyFound = True
          Exit For
        End If
    Next
MsgBox "Done: Keyfound=" & KeyFound & " If false means you are probably not using WinXp", vbInformation
End Sub
Private Sub mnuGetKey53a_Click()
    Call Key53AFind
'    MsgBox "I can find the key easy....Just don't get how to use it yet..."
End Sub

Private Sub mnuHowToUse_Click()
    frmHelp.Show vbModal, Me
End Sub

Private Sub mnuOpenExe_Click()
On Error GoTo errHandle
    CommonDialog1.Filename = ""
    CommonDialog1.DialogTitle = "Select Exe"
    CommonDialog1.Filter = "GameMaker Exe Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    
    Call VersionCheck(CommonDialog1.Filename, True)
Exit Sub
errHandle:
    MsgBox "Error_frmMain_mnuOpenExe: " & Err.Number & " " & Err.Description
End Sub
Public Sub VersionCheck(Filename As String, OpenExe As Boolean)
On Error GoTo errHandle

    VersionDetected = False
    txtResults.Text = ""
    
    If Filename <> "" Then
    txtResults.Text = txtResults.Text & "Opening " & Filename & vbCrLf
    txtResults.Text = txtResults.Text & "Checking GameMaker Version..." & vbCrLf
    
    Call AddLog(Filename)
    'Check if its gamemaker 5.0 version
    Dim Srch As String
    Srch = GameVer50
    Dim srch2 As String * 14
    Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
        Get #1, 477553, srch2
    
    Close #1 'Closes file after all of it is read
    If srch2 = Srch Then
        txtResults.Text = txtResults.Text & "GameMaker Version 5.0" & vbCrLf
        GameVersion = "5.0"
        VersionDetected = True
    End If
    'Check if its gamemaker 5.1 version
    If VersionDetected = False Then
        Srch = GameVer51
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            Get #1, 640029, srch2

       Close #1 'Closes file after all of it is read
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 5.1" & vbCrLf
            GameVersion = "5.1"
            VersionDetected = True
        End If
    End If
     'Check if its gamemaker 5.2 version
    If VersionDetected = False Then
        Srch = GameVer52
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            Get #1, 669457, srch2

       Close #1 'Closes file after all of it is read
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 5.2" & vbCrLf
            GameVersion = "5.2"
            VersionDetected = True
        End If
    End If
    If VersionDetected = False Then
        Srch = GameVer43
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            Get #1, 553105, srch2

       Close #1 'Closes file after all of it is read
     
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 4.3" & vbCrLf
            GameVersion = "4.3"
            VersionDetected = True
        End If
    End If
    If VersionDetected = False Then
    '5.3 check
        Srch = GameVer53Beta
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            Get #1, 679641, srch2

       Close #1 'Closes file after all of it is read
     
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 5.3 Beta" & vbCrLf
            GameVersion = "5.3Beta"
            VersionDetected = True
        End If
    End If
    If VersionDetected = False Then
    '5.3 Normal key check
        Srch = GameVer53
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            Get #1, 656857, srch2
            'Get #1, 679793, srch2
        
       Close #1 'Closes file after all of it is read
    
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 5.3" & vbCrLf
            GameVersion = "5.3"
            VersionDetected = True
        End If
    End If
    If VersionDetected = False Then
    '5.3a
        Srch = GameVer53
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            
            Get #1, 679793, srch2
      
       Close #1 'Closes file after all of it is read
    
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 5.3a" & vbCrLf
            txtResults.Text = txtResults.Text & "5.3a Encrypted with new encryption scheme" & vbCrLf
            GameVersion = "5.3a"
            VersionDetected = True
        End If
    End If
    If VersionDetected = False Then
    '5.4
        Srch = GameVer54
        sTemp = ""

        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            
            Get #1, 631505, srch2
      
       Close #1 'Closes file after all of it is read
    
        If srch2 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 5.4" & vbCrLf
            txtResults.Text = txtResults.Text & "5.4 Encrypted with a totaly new encryption scheme" & vbCrLf
            GameVersion = "5.4"
            VersionDetected = True
        End If
    End If
    
    If VersionDetected = False Then
    '6.0
        Srch = GameVer60
        sTemp = ""
        Dim srch3 As String * 11
        Open Filename For Binary Access Read Lock Read As #1 'Opens path specified by 'Filename' so VB can read it...
            
            Get #1, 557270, srch3
      
       Close #1 'Closes file after all of it is read
        If srch3 = Srch Then
            txtResults.Text = txtResults.Text & "GameMaker Version 6.0" & vbCrLf
            txtResults.Text = txtResults.Text & "6.0 Encrypted with a harder encryption scheme" & vbCrLf
            GameVersion = "6.0"
            VersionDetected = True
        End If
    End If
        If VersionDetected = True Then
            If OpenExe = True Then
                If KeyFound = False Then
                    Dim Response As String
                    Response = MsgBox("The encyrpiton key is not found! Do you want to use BruteForce to find the key? Press Yes to use brute force, No to continue decompiling, or Cancel to stop", vbYesNoCancel + vbInformation)
                    If Response = vbYes Then
                        frmBruteForce.Show
                    End If
                    If Response = vbNo Then
                        Call DecompileExe(Filename)
                    End If
                Else
                    Call DecompileExe(Filename)
                End If
            End If
        Else
            GameVersion = "none"
            MsgBox "Not a GameMaker 4.3 5.0 5.1 5.2 5.3Beta 5.3 5.3a 6.0 program", vbCritical
        End If
        
    End If
Exit Sub
errHandle:
    MsgBox "Error_frmMain_VersionCheck: " & Err.Number & " " & Err.Description
End Sub

Sub DecompileExe(Filename As String)
    Dim endKey As String
    Dim EndKey2 As String
    Dim BeginKey As String
    Dim BeginKey2 As String

    Dim NewArray(255) As Byte
    
    Call MakeUnEncrypted
    If OverRideNormalKey = True Then
        Call frmMain.LoadNormalByFilename(NormalFileName)
    Else
        Call MakeUnEncrypted
    
    End If
    If OverRideEncryptKey = True Then
        frmMain.LoadKeyByFilename (KeyFileName)
    Else
        If GameVersion = "5.0" Then
            key = Finalkey50
        ElseIf GameVersion = "5.1" Then
            key = Finalkey51
        ElseIf GameVersion = "5.2" Then
            key = modExtractor.Finalkey52
        ElseIf GameVersion = "5.3Beta" Then
            key = modExtractor.Finalkey53Beta
        ElseIf GameVersion = "4.3" Then
            key = modExtractor.Finalkey43
        ElseIf GameVersion = "5.3" Then
            key = modExtractor.Finalkey53Beta
        ElseIf GameVersion = "5.3a" Then
            key = modExtractor.Finalkey53A
          

            For i = 0 To 36
                endKey = endKey + Mid$(key, 220 + i, 1)
            Next
            For i = 1 To 37
                BeginKey2 = BeginKey2 + Mid$(key, i, 1)
            Next
            For i = 38 To 256
                EndKey2 = EndKey2 + Mid$(key, i, 1)
            Next
            f = FreeFile
            Open App.Path & "\endkey.txt" For Binary Access Write Lock Write As f
                Put f, , endKey
            Close f
           ' f = FreeFile
            'Open App.Path & "\endkey2.txt" For Binary Access Write Lock Write As f
              '  Put f, , EndKey2
            'Close f
            For i = 2 To 219
                BeginKey = BeginKey + Mid$(key, i, 1)
            Next
            f = FreeFile
            Open App.Path & "\beginkey.txt" For Binary Access Write Lock Write As f
                Put f, , BeginKey
            Close f
           ' f = FreeFile
            'Open App.Path & "\beginkey2.txt" For Binary Access Write Lock Write As f
              '  Put f, , BeginKey2
          '  Close f
            key = Mid$(key, 1, 1) + BeginKey + endKey
            f = FreeFile
            Open App.Path & "\bothkey.txt" For Binary Access Write Lock Write As f
               Put f, , key
            Close f
        ElseIf GameVersion = "5.4" Then
            key = modExtractor.Finalkey54
          

            For i = 0 To 36
                endKey = endKey + Mid$(key, 220 + i, 1)
            Next
            For i = 1 To 37
                BeginKey2 = BeginKey2 + Mid$(key, i, 1)
            Next
            For i = 38 To 256
                EndKey2 = EndKey2 + Mid$(key, i, 1)
            Next
            f = FreeFile
            Open App.Path & "\endkey.txt" For Binary Access Write Lock Write As f
                Put f, , endKey
            Close f
           ' f = FreeFile
            'Open App.Path & "\endkey2.txt" For Binary Access Write Lock Write As f
              '  Put f, , EndKey2
            'Close f
            For i = 2 To 219
                BeginKey = BeginKey + Mid$(key, i, 1)
            Next
            f = FreeFile
            Open App.Path & "\beginkey.txt" For Binary Access Write Lock Write As f
                Put f, , BeginKey
            Close f
           ' f = FreeFile
            'Open App.Path & "\beginkey2.txt" For Binary Access Write Lock Write As f
              '  Put f, , BeginKey2
          '  Close f
            key = Mid$(key, 1, 1) + BeginKey + endKey
                        f = FreeFile
            Open App.Path & "\bothkey.txt" For Binary Access Write Lock Write As f
               Put f, , key
            Close f
        ElseIf GameVersion = "6.0" Then
            key = modExtractor.Finalkey60
        End If
    End If
    txtResults.Text = txtResults.Text & "Begining Decompile..." & vbCrLf
       For i = 0 To 255
            NewArray(Asc(Mid$(key, i + 1, 1))) = Asc(Mid$(Unencrypted, i + 1, 1))
       Next
       
    On Error Resume Next
    Call MakeDir
    'Kill old copy
    Kill (App.Path & "\temp\copy.dat")
    Kill (App.Path & "\temp\game.enc")
    Kill (App.Path & "\temp\game.gmd")

    'copyfile
    FileCopy Filename, App.Path & "\temp\copy.dat"
    

    Dim Temp As Byte ' * 1

    Dim Temp2 As Byte

    If GameVersion = "4.3" Then
   
        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 1229823
                Do Until EOF(2)
                 
                    Get #2, , Temp
                    Temp2 = NewArray(Temp)
                    'For g = 1 To 256
                        'If Temp = Asc(Mid(key, g, 1)) Then
                         ' Temp2 = Asc(Mid(Unencrypted, g, 1))
                         ' Exit For
                        ' End If
                    'Next g
                 
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2
End If
    
    
    If GameVersion = "5.0" Then

        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 1229823
                Do Until EOF(2)
                 
                    Get #2, , Temp
                    Temp2 = NewArray(Temp)
                    'For g = 1 To 256
                        'If Temp = Asc(Mid(key, g, 1)) Then
                          'Temp2 = Asc(Mid(Unencrypted, g, 1))
                         ' Exit For
                         'End If
                    'Next g
                 
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2
     

        
        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf
        
     
    End If
    If GameVersion = "5.1" Then
    'Extract the game part of the file
    'Now Decrypt the file
              '  Debug.Print key
           ' Debug.Print "Normal: " & Unencrypted
     txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
              'Seek #2, OffsetVer51
              Seek #2, 1402368
              
            '  Dim middle As Integer
           '  Dim low As Integer, high As Integer
                 'low = 1
                 'high = 256
              
              Do Until EOF(2)
               
                 Get #2, , Temp
                Temp2 = NewArray(Temp)
                'low = 1
               ' high = 255
                'middle = 256
                 '############################
                   ' For g = 1 To 256
                       'If Temp = Asc(Mid(key, g, 1)) Then
                        ' Temp2 = Asc(Mid(Unencrypted, g, 1))
                        ' Exit For
                      ' End If
                   ' Next g
                    '########################
                    'For g = 128 To 256
                       ' If Temp = Asc(Mid(key, g, 1)) Then
                        '  Temp2 = Asc(Mid(Unencrypted, g, 1))
                        '  Exit For
                        'End If
                       ' If Temp = Asc(Mid(key, g - 1, 1)) Then
                        '  Temp2 = Asc(Mid(Unencrypted, g - 1, 1))
                       ''   Exit For
                      '  End If
                  '  Next g
                    
                   ' Do While (low <= high)
                      '  middle = (low + high) \ 2
 
                       ' If (Temp = Asc(Mid(key, middle, 1))) Then
                          '  Temp2 = Asc(Mid(Unencrypted, middle, 1))
                            
                          ' low = high + 300
                          ' Exit Do
                        'ElseIf Temp < Asc(Mid(key, middle, 1)) Then
                           ' high = middle - 1
                        'Else
                         '  low = middle + 1
                        ' End If
       
                   '  Loop
                       'Do
                         '   middle = (low + high) \ 2

                        ' Select Case StrComp(Mid(key, middle, 1), Chr(Temp), vbBinaryCompare)
                               ' Case -1: low = middle + 1
                                'Case 1: high = middle - 1
                               ' Case 0
                               ' Temp2 = Asc(Mid(Unencrypted, middle, 1))
                              '  low = high + 1
                              '  Exit Do
                           ' End Select
                  '     Loop Until low > high
                 
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
             
             
            Close #3
        Close #2


       
         

    End If
   
    If GameVersion = "5.2" Then
    'Extract the game part of the file
    txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
         Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            'Now decrypt the file

            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 1440255 'OffsetVer52
              Do Until EOF(2)
                 
                 Get #2, , Temp
                 Temp2 = NewArray(Temp)
                   ' For g = 1 To 256
                        'If Temp = Asc(Mid(key, g, 1)) Then
                         ' Temp2 = Asc(Mid(Unencrypted, g, 1))
                         
                          'Exit For
                        ' End If
                    'Next g
  
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3

            Filesize = LOF(2)
        Close #2
         
        
  

        'Extratcting sounds
        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf

        f = FreeFile
         

        

    End If
    If GameVersion = "5.3Beta" Then

        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 1482241
                Do Until EOF(2)
                 
                    Get #2, , Temp
                    Temp2 = NewArray(Temp)
                   ' For g = 1 To 256
                       ' If Temp = Asc(Mid(key, g, 1)) Then
                        '  Temp2 = Asc(Mid(Unencrypted, g, 1))
                        '  Exit For
                        ' End If
                  '  Next g
                 
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2
     

        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf

     
    End If
    If GameVersion = "5.3" Then

        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file

            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 1483264
                Do Until EOF(2)
                 
                    Get #2, , Temp
                    

                     Temp2 = NewArray(Temp)
                   'For g = 1 To 256
                        'If Temp = Asc(Mid(key, g, 1)) Then
                        ' Temp2 = Asc(Mid(Unencrypted, g, 1))
                        '  Exit For
                        ' End If
                  ' Next g
                    
            
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2

        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf

    End If
   If GameVersion = "5.3a" Then

        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
       

       
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 1484286
                Do Until EOF(2)
                 
                    Get #2, , Temp
                    

                Temp2 = NewArray(Temp)
                       
                  ' For g = 1 To 256
                      '  If Temp = Asc(Mid(key, g, 1)) Then
                       '  Temp2 = Asc(Mid(Unencrypted, g, 1))
                       '   Exit For
                     '    End If
                  ' Next g
                    
            
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2

        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf

    End If
  If GameVersion = "5.4" Then
        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
              ' Seek #2, 1465344  'example1
              
                Do Until EOF(2)
                    Get #2, , Temp
                Temp2 = NewArray(Temp)
                       
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2
        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf
    End If

  If GameVersion = "6.0" Then
        txtResults.Text = txtResults.Text & "Decrypting exe..." & vbCrLf
        Open App.Path & "\temp\copy.dat" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
            Open App.Path & "\temp\game.enc" For Binary Access Write Lock Write As #3
                Seek #2, 568320  'example1
              
                Do Until EOF(2)
                    Get #2, , Temp
                Temp2 = NewArray(Temp)
                       
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp2
                 End If
             Loop
            Close #3
        Close #2
        txtResults.Text = txtResults.Text & "FileSize: " & Filesize & vbCrLf
    End If
    
    txtResults.Text = txtResults.Text & "Decrytping done!" & vbCrLf

    txtResults.Text = txtResults.Text & "Done!" & vbCrLf
End Sub


Private Sub mnuOptions_Click()
    frmOptions.Show vbModal, Me
End Sub
Sub MakeDir()
On Error Resume Next
    MkDir (App.Path & "\dump")
    MkDir (App.Path & "\dump\images")
    MkDir (App.Path & "\dump\sounds")
    MkDir (App.Path & "\dump\scripts\")
    MkDir (App.Path & "\temp\")
End Sub

Private Sub mnuPlugin_Click()

    mnuPluginList.Visible = True
    Call Load_Plugins(frmMain)
End Sub

Private Sub mnuPluginArray_Click(Index As Integer)
On Error GoTo Err_Handle
LoadPlugin mnuPluginArray(Index).Caption + ".dll", mnuPluginArray(Index).Caption + ".Main", Me
Exit Sub
Err_Handle:

MsgBox App.Title & " Has Caused A Error And It Will Now Close." & vbCrLf & vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description, vbCritical, "Error - " & Err.Description

End Sub

Private Sub mnuWavExtract_Click()
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long

    Do Until Pos = -1
        Pos = RichTextBox1.Find("RIFF", oldpos + 4)
        larray(UBound(larray)) = Pos

        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
   

    Loop
    txtResults.Text = ""
    
    txtResults.Text = txtResults.Text & "Extracting Wav: " & UBound(larray) & vbCrLf
    If ShowWavoffsets = True Then
        txtResults.Text = txtResults.Text & "Possible wav offsets:"
    End If
    
    For i = 0 To UBound(larray)
    If larray(i) <> -1 Then
        If ShowWavoffsets = True Then
            txtResults.Text = txtResults.Text & larray(i) & vbCrLf
        End If
        
      GetWav CommonDialog1.Filename, (larray(i) + 1)
    End If
    Next
MsgBox "Done: Check " & App.Path & "\dump\sounds\", vbInformation
End Sub
Sub PluginExtract(ByVal RegisterFile As String, ByVal Dll_Open, ByVal Form As Form)
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    CommonDialog1.DialogTitle = "Select Dump (dump.txt)"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "Dump Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.Filename = "" Then Exit Sub
    Close
    RichTextBox1.LoadFile CommonDialog1.Filename, 1
    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long

    Do Until Pos = -1
        Pos = RichTextBox1.Find(pSearchString, oldpos + 4)
        larray(UBound(larray)) = Pos

       oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
   

    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Extracting " & RegisterFile & ": " & UBound(larray) & vbCrLf


    For i = 0 To UBound(larray)
    If larray(i) <> -1 Then
        
      Call CreateObject(Dll_Open).ProcessFile(CommonDialog1.Filename, (larray(i) + 1), App.Path)
    End If
    Next

End Sub
Public Sub LoadNormalByFilename(Filename As String)
On Error GoTo nofile:
d = FreeFile
Open Filename For Binary Access Read Lock Read As d
    Get d, , Unencrypted
Close d

Exit Sub
nofile:
Exit Sub
End Sub
Public Sub LoadKeyByFilename(Filename As String)
On Error GoTo nofile:
k = FreeFile
Open Filename For Binary Access Read As k
    Get k, , key
Close k


Exit Sub
nofile:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source
Exit Sub
End Sub
Public Sub FindGMDStartOffset()
On Error GoTo nofile:
Dim larray() As Long
'make the directories
Call MakeDir
ReDim larray(0)
    Close
    RichTextBox1.LoadFile App.Path & "\temp\game.enc", 1
    
    Dim SearchGmd As String * 3
    
    SearchGmd = Chr(145) & Chr(213) & Chr(18)

    Dim Pos As Long
    Pos = 0
    Dim oldpos As Long
  
    Do Until Pos = -1
        Pos = RichTextBox1.Find(SearchGmd, oldpos + Len(SearchGmd))
        
        larray(UBound(larray)) = Pos
        oldpos = Pos
       
        ReDim Preserve larray(UBound(larray) + 1)
    Loop
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "Finding GMD Start Offset Location UBound:" & UBound(larray) & vbCrLf

    
    For i = 0 To UBound(larray) - 1
        If larray(i) <> -1 Then
            
             
                modExtractor.ExtractGmd (larray(i) + 1)
                txtResults.Text = txtResults.Text & larray(i) + 1 & vbCrLf
                txtResults.Text = txtResults.Text & "End of GmdStartOffset List" & vbCrLf
              Exit For
        End If
    Next
    Dim Response As String
Response = MsgBox("Done! Extracted to: " & App.Path & "\temp\game.gmd  Do you want to open this file now?", vbYesNo + vbInformation, "Open extracted gmd?")
If Response = vbYes Then
    ShellExecute Me.hWnd, vbNullString, App.Path & "\temp\game.gmd", vbNullString, "C:\", SW_SHOWNORMAL

End If

Exit Sub
nofile:
    MsgBox Err.Description
Exit Sub
End Sub
Sub OneStepDecompile(Filename As String)
    Dim ProcessId As Long
    Dim memHandle As Long
    Dim buffer As String * 1000
    Dim readlen As Long
    Dim addr As Long
    Dim i As Long
    Dim f As Long, g As Long
    'First Check version
    Call VersionCheck(Filename, False)
    If GameVersion = "none" Then
        
        Exit Sub
    End If

    ProcessId = Shell(modGlobals.GetShortPathName(Filename), vbMinimizedFocus)
    txtResults.Text = ""
    txtResults.Text = txtResults.Text & "One Step Mode" & vbCrLf
    txtResults.Text = txtResults.Text & "GameMaker Version: " & GameVersion & vbCrLf
    txtResults.Text = txtResults.Text & "Filename: " & Filename & vbCrLf
    txtResults.Text = txtResults.Text & "ProcessID: " & ProcessId & vbCrLf
  
    'Wait a while
    For i = 0 To 30000
        g = g + 1
    Next i
    Sleep (6000)
    
    txtResults.Text = txtResults.Text & "Opening Process for Memory Reading" & vbCrLf
    memHandle = OpenProcess(&H1F0FFF, False, ProcessId)

    If (memHandle = 0) Then
         txtResults.Text = txtResults.Text & "Memory Reading Failed" & vbCrLf
    Else
         Call frmSelectProcess.KillOldDump
         txtResults.Text = txtResults.Text & "Reading Memory" & vbCrLf
        f = FreeFile
        Open App.Path & "\dump.txt" For Binary Access Write Lock Write As #f
        For addr = 4194304 To 14594304 Step 1000
            Call ReadProcessMemory(memHandle, addr, buffer, 1000, readlen)
            Put #f, , buffer
        Next

        Close #f
        txtResults.Text = txtResults.Text & "Memory Dump Completed" & vbCrLf
        TerminateProcess memHandle, 0
        CloseHandle (memHandle)
        
        
        CommonDialog1.Filename = "dump.txt"
        txtResults.Text = txtResults.Text & "Searching for Encryption Key" & vbCrLf
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
        ElseIf GameVersion = "4.3" Then
            frmMain.Key43Find
        Else
            MsgBox "Not a GM exe", vbExclamation
        End If
        If KeyFound = True Then
            txtResults.Text = txtResults.Text & "Encryption Key found!" & vbCrLf
        Else
            txtResults.Text = txtResults.Text & "Encryption Key NOT Found! Means you don't have WinXP Pro" & vbCrLf
        End If
        'Decrypt exe
        txtResults.Text = txtResults.Text & "Decrypting Exe" & vbCrLf
            Call frmMain.DecompileExe(Filename)
        txtResults.Text = txtResults.Text & "Decrypting Done" & vbCrLf
        txtResults.Text = txtResults.Text & "Extracting Gmd" & vbCrLf
            Call frmMain.FindGMDStartOffset
        txtResults.Text = txtResults.Text & "Decompile Finished!" & vbCrLf
    End If
    

    
End Sub
Public Function MakeGM6Key(strKey As String) As String
    Dim i As Long
    Dim strData As String
    For i = 0 To 255
        strData = Chr(Asc(Mid$(strKey, i + 1, 1)))
    Next

End Function
Private Function SwithPos()

End Function
