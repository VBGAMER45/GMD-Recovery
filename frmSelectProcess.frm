VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelectProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a Process"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Tag             =   "                                                               v bg am er45 "
   Begin VB.PictureBox picLarge 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2520
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picSmall 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      Picture         =   "frmSelectProcess.frx":0000
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dump Memory"
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtHigh 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Text            =   "4594304"
         ToolTipText     =   "UpperLimit of Memory that is going to be dumped."
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtLow 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "4194304"
         ToolTipText     =   "Base Address of the process. For GameMaker i have it set already."
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdDump 
         Caption         =   "Dump Memory"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Upper Limit"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Lower Limit"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ListBox lstProcess 
      Height          =   2010
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ListView AppRun 
      Height          =   3375
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5953
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectProcess.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectProcess.frx":08A4
            Key             =   "Test"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmSelectProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################
'frmSelectProcess
'vbgamer45
'#################################################
Dim myHandle As Long
Dim Lowerlimit As Long
Dim HighLimit As Long
Dim Number As Long
Dim itemx As ListItem
Dim ItemName As String

Function InitProcessCheater(pid As Long)
Dim pHandle As Long
pHandle = OpenProcess(&H1F0FFF, False, pid)

If (pHandle = 0) Then
    InitProcessCheater = False
    myHandle = 0
Else
    InitProcessCheater = True
    myHandle = pHandle

    
End If

End Function

Private Sub cmdClose_Click()
    Unload Me
'    Frame1.Visible = False
End Sub

Private Sub cmdClose2_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Dim myProcess As PROCESSENTRY32
    Dim mySnapshot As Long

    'first clear our listbox
    lstProcess.Clear

    myProcess.dwSize = Len(myProcess)

    'create snapshot
    mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

    'get first process
    ProcessFirst mySnapshot, myProcess
    lstProcess.AddItem myProcess.szexeFile ' set exe name
    PIDs(lstProcess.ListCount - 1) = myProcess.th32ProcessID ' set PID
Dim k As Integer
    'while there are more processes
    While ProcessNext(mySnapshot, myProcess)
        lstProcess.AddItem myProcess.szexeFile ' set exe name
       ' Debug.Print modMem.ExePathFromProcessId(myProcess.th32ProcessID)
        Call Load_Filename(modMem.ExePathFromProcessId(myProcess.th32ProcessID), myProcess.szexeFile, myProcess.th32ProcessID, Str(k))
        PIDs(lstProcess.ListCount - 1) = myProcess.th32ProcessID ' ' store PID
    k = k + 1
    Wend

End Sub

Private Sub cmdSelect_Click()

' `Get the Selected the Item
    For Each itemx In AppRun.ListItems
        If itemx.Selected = True Then
                ItemName = itemx.Text

                Number = itemx.Tag
                txtLow.Text = modMem.BaseModuleHandleFromProcessId(Number)
                Exit For
        End If
        
    Next itemx
    

 
    If ItemName = "System" Then MsgBox "Please select a process to open.", vbCritical, "Select Process": Exit Sub
    
Frame1.Visible = True
End Sub

Private Sub cmdDump_Click()
    Dim buffer As String * 1000
    Dim readlen As Long
    Dim addr As Long
    Dim f As Long
    Lowerlimit = txtLow.Text
    HighLimit = txtHigh.Text
    
    Call KillOldDump
    'init cheater
    If Not InitProcessCheater(Number) Then MsgBox "Could not open process. sorry :(", vbCritical, "GameMaker Decompiler": Exit Sub
    f = FreeFile
     Open App.Path & "\dump.txt" For Binary Access Write Lock Write As #f
     For addr = Lowerlimit To HighLimit Step 1000
        Call ReadProcessMemory(myHandle, addr, buffer, 1000, readlen)
       Put #f, , buffer
    
    Next
    
    Close #f
    
    TerminateProcess myHandle, 0
    
    MsgBox "DONE Open " & App.Path & "\dump.txt   When you find the key for the encryption", vbInformation

End Sub
Sub KillOldDump()
On Error Resume Next
    Kill (App.Path & "\dump.txt")
End Sub



Private Sub Form_Load()
    cmdRefresh_Click
End Sub

Private Sub txtHigh_Change()
    If IsNumeric(txtHigh.Text) = False Then txtHigh.Text = 0
End Sub

Private Sub txtLow_Change()
    If IsNumeric(txtLow.Text) = False Then txtLow.Text = 0
End Sub
Private Sub Load_Filename(sExeName As String, ExeTitle As String, pid As Long, Optional KeyNumber As String)
'Load the Icon into the List View

ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)

On Error GoTo ErrFound


Dim lIndex

lIndex = "0"

'Get Icon from the File

Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)

With picLarge
    Set .Picture = LoadPicture("")
     .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)

     .Refresh
End With

Mykey = sExeName & "(" & "-" & KeyNumber & ")"
    If glLargeIcons(lIndex) <> 0 Then
        ImageList1.ListImages.Add , Mykey, picLarge.Image
    Else
        ImageList1.ListImages.Add , Mykey, picSmall.Image
    End If
    
txtMax = sExeName
Dim t As ListItem
' Add Icon to Listview
Set t = AppRun.ListItems.Add(, txtMax, ExeTitle, Mykey)
t.Tag = pid

ErrFound:

End Sub

