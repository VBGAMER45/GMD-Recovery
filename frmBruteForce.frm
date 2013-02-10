VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBruteForce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brute Force Search for Encryption Key"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "St&art"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Possible Key Offsets"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label lblFilename 
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Select Memory Dump. Normally dump.txt"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmBruteForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumArray(255) As Byte
Private Sub cmdBrowse_Click()
        CommonDialog1.Filename = ""
        CommonDialog1.DialogTitle = "Select Dump file"
        CommonDialog1.Filter = "Dump Files (*.*)|*.*"
        CommonDialog1.ShowOpen
    
        If CommonDialog1.Filename = "" Then Exit Sub
        
        lblFilename.Caption = CommonDialog1.Filename
        
        cmdStart.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
List1.Clear
    Dim Fpos As Long
    Dim Num As Long
    Dim Found As Boolean
    Dim i As Long
    Dim Response As String
    Dim Count As Long
    If lblFilename.Caption <> "" Then
    
    F = FreeFile
    Found = False
        Open lblFilename.Caption For Binary Access Read Lock Read As F
            Seek F, 500000
            Count = LOF(F) '500000
            Fpos = 500000
            lblProgress.Caption = Count
            Do While Fpos < LOF(F) And Found = False
            
            Seek F, Fpos
            Get F, Fpos, NumArray
            'Test loop
            For i = 0 To 255
                Num = Num + NumArray(i)
            Next
                If Num = 32640 Then
                    'MsgBox "Keyfound!! at Offset: " & Fpos
                    If vaildKey = False Then
                        'MsgBox "Not vaild key!"
                        lblProgress.Refresh
                    Else
                    
                    'Found = True
                         lblProgress.Refresh
                         Response = MsgBox("Vaild Key at offset:" & Fpos & " Do you want to continue searching?", vbYesNo + vbInformation)
                         If Response = vbNo Then Found = True
                         
                    
                        List1.AddItem Fpos
                    End If
                End If
            
                Fpos = Fpos + 1
                Count = Count - 1
                lblProgress.Caption = Count
                Num = 0
            Loop
    
        Close F
        If Found = False Then
            MsgBox "No KeyFound....Dam it..", vbCritical
        End If
        If Found = True Then
            MsgBox "Ok Search done. Click on an offset in the list to load an encryption key you can only load one at a time. Usally its the second on the list...", vbInformation
        End If
        
    End If
   
    Label3.Caption = "Click on a list item to use that key to decompile"
End Sub
Function vaildKey() As Boolean
Dim Counter(255) As Byte
    For i = 0 To 255
        Counter(NumArray(i)) = Counter(NumArray(i)) + 1
    Next
    For i = 0 To 255
        If Counter(i) > 1 Then
            vaildKey = False
            Exit Function
            
        End If
         If Counter(i) = 0 Then
            vaildKey = False
            Exit Function
            
        End If
        
    Next
    
    vaildKey = True
End Function

Private Sub List1_Click()
    If List1.Text <> "" Then
    F = FreeFile
    Open lblFilename.Caption For Binary Access Read Lock Read As F
        Offset = List1.Text
        Offset = Offset + 1
        Get F, Offset, modExtractor.Finalkey53A
        Get F, Offset, modExtractor.Finalkey53Beta
        Get F, Offset, modExtractor.Finalkey52
        Get F, Offset, modExtractor.Finalkey51
        Get F, Offset, modExtractor.Finalkey50
   
    Close F
    Open App.Path & "\bruteforcekey.txt" For Binary Access Write Lock Write As #10
        Put #10, , modExtractor.Finalkey52
    Close #10
    
    KeyFound = True
    End If
    MsgBox "Encryption key loaded continue onto next step of the guide.", vbInformation
End Sub
