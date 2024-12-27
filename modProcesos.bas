Attribute VB_Name = "modProcesos"
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

Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
ByVal lFlags As Long, lProcessID As Long) As Long

Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
 

Public Sub KillProcess(NameProcess As String)
Const PROCESS_ALL_ACCESS = 1
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapshot As Long
Dim SzExename As String
Dim ExitCode As Long
Dim MyProcess As Long
Dim AppKill As Boolean
Dim AppCount As Integer
Dim i As Integer
Dim WinDirEnv As String

If NameProcess <> "" Then
AppCount = 0

uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
RProcessFound = ProcessFirst(hSnapshot, uProcess)

Do
i = InStr(1, uProcess.szexeFile, Chr(0))
SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
WinDirEnv = Environ("Windir") + "\"
WinDirEnv = LCase$(WinDirEnv)

If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
AppCount = AppCount + 1
MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
AppKill = TerminateProcess(MyProcess, ExitCode)
Call CloseHandle(MyProcess)
End If
RProcessFound = ProcessNext(hSnapshot, uProcess)
Loop While RProcessFound
Call CloseHandle(hSnapshot)
End If

End Sub

Public Function estacorriendo(ByVal NombreDelProceso As String) As Boolean
Const MAX_PATH As Long = 260
Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
Dim sName As String
NombreDelProceso = UCase$(NombreDelProceso)
ReDim lProcesses(1023) As Long
If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
For N = 0 To (lRet \ 4) - 1
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
If hProcess Then
ReDim lModules(1023)
If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
sName = String$(MAX_PATH, vbNullChar)
GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
sName = Left$(sName, InStr(sName, vbNullChar) - 1)
If Len(sName) = Len(NombreDelProceso) Then
If NombreDelProceso = UCase$(sName) Then estacorriendo = True: Exit Function
End If
End If
End If
CloseHandle hProcess
Next N
End If
End Function


