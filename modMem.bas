Attribute VB_Name = "modMem"
'###############################################
' modMem.bas
' vbgamer45
'###############################################
'Holds all the memory api's i use.
Option Explicit
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
Declare Function VirtualQueryEx Lib "kernel32.dll" ( _
     ByVal hProcess As Long, _
     lpAddress As Any, _
     lpBuffer As MEMORY_BASIC_INFORMATION, _
     ByVal dwLength As Long) As Long
Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Declare Function EnumProcessModules Lib "psapi" _
   (ByVal ProcessId As Long, hModule As Long, ByVal cbSize As Long, _
    cbReturned As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi" _
   Alias "GetModuleFileNameExA" (ByVal hProcess As Long, _
   ByVal hModule As Long, ByVal lpFileName As String, _
   ByVal nSize As Long) As Long
Private Const MAX_PATHLEN = &H104
Private Type PROCESSENTRY31
      dwSize As Long
      cntUsage  As Long
      th32ProcessID As Long
      th32DefaultHeapID As Long
      th32ModuleID As Long
      cntThreads As Long
      th32ParentProcessID As Long
      pcPriClassBase As Long
      dwFlags As Long
      exeFilename(1 To MAX_PATHLEN) As Byte
End Type
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
   (ByVal dwFlags As Long, ByVal dprocess As Long) As Long
Private Declare Function Process32First Lib "kernel32" _
   (ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY31) As Long

Private Declare Function Process32Next Lib "kernel32" _
   (ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY31) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Global PIDs(1000) As Long
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION

Function BaseModuleHandleFromProcessId(ProcessId As Long)
   Dim hModule As Long, hProcess As Long, bReturned As Long
   hProcess = ProcessHandleFromProcessId(ProcessId)
   Call EnumProcessModules(hProcess, hModule, 4&, bReturned)
   BaseModuleHandleFromProcessId = hModule
End Function
Function ExePathFromProcessId(ByVal ProcessId As Long) As String
   Dim hProcess As Long
   '
  
   If Not IsWindowsNT Then   ' Win2000 is included
      
      Dim Process As PROCESSENTRY31, hSnap As Long, F As Long, Exename$
      hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
      If hSnap = -1 Then Exit Function
      Process.dwSize = Len(Process)
      F = Process32First(hSnap, Process)
      Do While F <> 0
         If Process.th32ProcessID = ProcessId Then
            GoSub GetExeName
            ExePathFromProcessId = Exename ' strZtoStr(process.exeFilename)
            Call CloseHandle(hSnap)
            Exit Function
            End If
         Process.dwSize = Len(Process)
         F = Process32Next(hSnap, Process)
      Loop
   Else

      Dim s As String, c As Long, hModule As Long
      Const cMaxPath = 1023
      s = String$(cMaxPath, 0)
      hProcess = ProcessHandleFromProcessId(ProcessId)
      hModule = BaseModuleHandleFromProcessId(ProcessId)
      c = GetModuleFileNameEx(hProcess, hModule, s, cMaxPath)
      If c Then ExePathFromProcessId = Left$(s, c)
      End If
    
   Exit Function

GetExeName:
   Dim i&, cb As Byte
   Exename = ""
   Do While i < MAX_PATHLEN
      i = i + 1
      cb = Process.exeFilename(i)
      If cb = 0 Then Return
      Exename = Exename & Chr$(cb)
      Loop
     
   Return
End Function

Private Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Function ProcessHandleFromProcessId(ProcessId As Long)
   ProcessHandleFromProcessId = _
      OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessId)
End Function
