VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExeDecrypt 
   Caption         =   "Exe Decrypt Beta"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenExe 
      Caption         =   "&Open Exe"
      Default         =   -1  'True
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"frmExeDecrypt.frx":0000
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmExeDecrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOpenExe_Click()
    CommonDialog1.Filename = ""
    CommonDialog1.DialogTitle = "Select Exe"
    CommonDialog1.Filter = "GameMaker Exe Files (*.exe)|*.exe"
    CommonDialog1.ShowOpen
    
    
    If CommonDialog1.Filename = "" Then Exit Sub
    'check version
    Call frmMain.VersionCheck(CommonDialog1.Filename, False)
    
    'Make unencrypted key
    Dim unArray(255) As Byte
    Dim i As Integer
    Dim F As Long
        For i = 0 To 255
            unArray(i) = i
        Next i
    
    F = FreeFile
    
    
    If GameVersion = "5.0" Then
        Open CommonDialog1.Filename For Binary Access Read Lock Read As F
            Seek F, eOffset50 + 1
        Close F
        Open App.Path & "\testexe.txt" For Binary Access Write As #14
            Put #14, , unArray
        Close #14
    ElseIf GameVersion = "5.1" Then
        Open CommonDialog1.Filename For Binary Access Read Lock Read As F
            Seek F, eOffset51 + 1
        Close F
        
        Open App.Path & "\testexe.txt" For Binary Access Write As #14
            Put #14, , unArray
        Close #14
    ElseIf GameVersion = "5.2" Then
        Dim eList(37) As Byte
        Open CommonDialog1.Filename For Binary Access Read Lock Read As F
            Seek F, eOffset52 + 1
            frmMain.txtResults.Text = frmMain.txtResults.Text & "Getting Encryption key" & vbCrLf
            Get F, , eList
            'For i = 0 To UBound(eList)
               ' MsgBox eList(i)
            'Next
           ' MsgBox eList(4)
        Close F
        'Do decyrpt
        Dim Temp As Byte
        Temp = unArray(eList(4) - 1)
        unArray(eList(4) - 1) = eList(4)
        unArray(eList(4)) = Temp
       ' MsgBox eList(8)
        Temp = unArray(eList(8) + 1)
        unArray(eList(8) + 1) = eList(8)
        unArray(eList(8)) = Temp
        
        Temp = unArray(eList(9) + 1)
        unArray(eList(9) + 1) = eList(9)
        unArray(eList(9)) = Temp
        
        Open App.Path & "\testexe.txt" For Binary Access Write As #14
            Put #14, , unArray
        Close #14
    ElseIf GameVersion = "5.3Beta" Then
        Open CommonDialog1.Filename For Binary Access Read Lock Read As F
            Seek F, eOffset53beta + 1
        Close F
        
        Open App.Path & "\testexe.txt" For Binary Access Write As #14
            Put #14, , unArray
        Close #14
    ElseIf GameVersion = "5.3" Then
        Open CommonDialog1.Filename For Binary Access Read Lock Read As F
            Seek F, eOffset53 + 1
        Close F
    ElseIf GameVersion = "5.3a" Then
        Open CommonDialog1.Filename For Binary Access Read Lock Read As F
            Seek F, eOffset53a + 1
        Close F
        
        Open App.Path & "\testexe.txt" For Binary Access Write As #14
            Put #14, , unArray
        Close #14
    
    Else
        MsgBox "Not a Gm exe or not supported version...", vbInformation
    End If
End Sub

