Attribute VB_Name = "modRegPlug"
'##################################
'modRegPlug.bas
'vbgamer45
'##################################
'Main use in the plugins

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0
Private Const ERROR_AHHHHHH = &HF
Global pSearchString As String
Public Function RegisterServer(hWnd As Long, DllServerPath As String, bRegister As Boolean)
On Error Resume Next
Dim lb As Long, pa As Long
    lb = LoadLibrary(DllServerPath)
    If bRegister Then
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If
    If CallWindowProc(pa, hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
        RegisterServer = ERROR_SUCCESS
    Else
        RegisterServer = ERROR_AHHHHHH
    End If
FreeLibrary lb
End Function
'This works by: File To Regiter,Dll to Open (plugin),and the form that is opening it
Sub LoadPlugin(ByVal RegisterFile As String, ByVal Dll_Open, ByVal Form As Form)
RegisterFile = App.Path + "\Plugins\" + RegisterFile
RegisterServer Form.hWnd, RegisterFile, True
'Change This Here To Open A Diffrent Sub In The Class Module In The Dll
CreateObject(Dll_Open).Load
pSearchString = CreateObject(Dll_Open).GetSearchString()
pSearchString = Trim(pSearchString)
'MsgBox SearchString

Call frmMain.PluginExtract(RegisterFile, Dll_Open, Form)

RegisterServer Form.hWnd, RegisterFile, False

End Sub

Sub ShowAbout(ByVal RegisterFile As String, ByVal Dll_Open, ByVal Form As Form)
RegisterFile = App.Path + "\Plugins\" + RegisterFile
RegisterServer Form.hWnd, RegisterFile, True
'Change This Here To Open A Diffrent Sub In The Class Module In The Dll
CreateObject(Dll_Open).About

RegisterServer Form.hWnd, RegisterFile, False
End Sub

Sub Load_Plugins(Form As Form)
Dim x As Integer
On Error Resume Next
x = 0
Dim menu As New FileSystemObject
Set Folder = menu.GetFolder(App.Path + "\Plugins\")
For Each File In Folder.Files
    If Right(File.Name, 4) = ".dll" Then Else GoTo nextfile

    Load Form.mnuPluginArray(x)
    With Form.mnuPluginArray(x)
        .Caption = Mid$(File.Name, 1, Len(File.Name) - 4)
        .Tag = File
        .Visible = True
        .Enabled = True
    End With
    x = x + 1
nextfile:
    Next
End Sub


