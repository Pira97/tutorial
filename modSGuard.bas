Attribute VB_Name = "modSGuard"
Option Explicit

Private Type tDef
    MainCName As String
    strHush As String * 32
End Type

Private Type tArrProc
    strName As String
    ID As Long
    strMD5 As String
    bStatus As Boolean
End Type

Private Definition_List() As tDef

Private arrProc() As tArrProc
Private lngProc As Long

Private lngDef As Long

Private Declare Function Process32First Lib "kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function CloseHandle Lib "Kernel32.dll" _
   (ByVal handle As Long) As Long

Private Declare Function OpenProcess Lib "Kernel32.dll" _
  (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
      ByVal dwProcId As Long) As Long

Private Declare Function EnumProcesses Lib "psapi.dll" _
   (ByRef lpidProcess As Long, ByVal cb As Long, _
      ByRef cbNeeded As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" _
   (ByVal hProcess As Long, ByVal hModule As Long, _
      ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" _
   (ByVal hProcess As Long, ByRef lphModule As Long, _
      ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
   ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Private Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long           ' This process
   th32DefaultHeapID As Long
   th32ModuleID As Long            ' Associated exe
   cntThreads As Long
   th32ParentProcessID As Long     ' This process's parent process
   pcPriClassBase As Long          ' Base priority of process threads
   dwFlags As Long
   szExeFile As String * 260       ' MAX_PATH
End Type

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS = &H2&
Private Const hNull = 0

Public Function Load_Definitions() As Boolean

On Error GoTo ErrorHandler

Dim N As Integer
Dim i As Long
Dim strLine As String

If Extract_File(Scripts, App.Path & "\Recursos", "ggdef.dat", Windows_Temp_Dir) Then
    
    lngDef = General_Get_Line_Count(Windows_Temp_Dir & "ggdef.dat")
    If lngDef <= 0 Then Exit Function
    
    ReDim Definition_List(1 To lngDef) As tDef
    
    N = FreeFile()
    
    Open Windows_Temp_Dir & "ggdef.dat" For Input As #N
    
    For i = 1 To lngDef
        Line Input #N, strLine
        Definition_List(i).MainCName = General_Field_Read(1, strLine, ";")
        Definition_List(i).strHush = General_Field_Read(2, strLine, ";")
    Next i
    
    Close #N
    
    Delete_File Windows_Temp_Dir & "ggdef.dat"
    
    Load_Definitions = True
End If

Exit Function

ErrorHandler:

End Function

Public Function Main_Logic() As Long

Dim strRet As String

strRet = GetStrProc

If LenB(strRet) > 0 Then
    Call Clienttcp.Send_Data(GameMain_Logic, strRet)
    Call General_Sleep(3.5)
    Call MsgBox(Locale_GUI_Frase(346), vbExclamation, Locale_GUI_Frase(347))
    EndGame
End If

End Function

Private Function GetStrProc() As String

On Error GoTo ErrorHandler

Dim i As Long, j As Long
Dim bCAppFound As Boolean

If Not General_Windows_Is_2000XP Then
   
   Dim f As Long, sname As String
   Dim hSnap As Long, proc As PROCESSENTRY32
   hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
   If hSnap = hNull Then Exit Function
   proc.dwSize = Len(proc)
   ' Iterate through the processes
   f = Process32First(hSnap, proc)
   Do While f
     sname = StrZToStr(proc.szExeFile)
     If Not pExist(proc.th32ProcessID) Then
        lngProc = lngProc + 1
        ReDim Preserve arrProc(1 To lngProc)
        arrProc(lngProc).strName = sname
        arrProc(lngProc).ID = proc.th32ProcessID
        arrProc(lngProc).strMD5 = MD5File(arrProc(lngProc).strName)
     End If
     f = Process32Next(hSnap, proc)
   Loop

Else

   Dim cb As Long
   Dim cbNeeded As Long
   Dim NumElements As Long
   Dim ProcessIDs() As Long
   Dim cbNeeded2 As Long
   Dim NumElements2 As Long
   Dim Modules(1 To 200) As Long
   Dim lRet As Long
   Dim ModuleName As String
   Dim nSize As Long
   Dim hProcess As Long
   'Get the array containing the process id's for each process object
   cb = 8
   cbNeeded = 96
   Do While cb <= cbNeeded
      cb = cb * 2
      ReDim ProcessIDs(cb / 4) As Long
      lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
   Loop
   NumElements = cbNeeded / 4

   For i = 1 To NumElements
      'Get a handle to the Process
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
         Or PROCESS_VM_READ, 0, ProcessIDs(i))
      'Got a Process handle
      If hProcess <> 0 Then
          'Get an array of the module handles for the specified
          'process
          lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                       cbNeeded2)
          'If the Module Array is retrieved, Get the ModuleFileName
          If lRet <> 0 Then
             ModuleName = Space(MAX_PATH)
             nSize = 500
             lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                             ModuleName, nSize)
            
            If Not pExist(ProcessIDs(i)) Then
                lngProc = lngProc + 1
                ReDim Preserve arrProc(1 To lngProc)
                arrProc(lngProc).strName = Left$(ModuleName, lRet)
                arrProc(lngProc).ID = ProcessIDs(i)
                arrProc(lngProc).strMD5 = MD5File(arrProc(lngProc).strName)
            End If
            
          End If
      End If
    'Close the handle to the process
   lRet = CloseHandle(hProcess)
   Next

End If

If lngProc <= 0 Then GoTo ErrorHandler

For i = 1 To lngProc
    
    If arrProc(i).bStatus = False Then
        For j = 1 To lngDef
            If Definition_List(j).strHush = arrProc(i).strMD5 Then
                GetStrProc = Definition_List(j).MainCName & " (PID" & arrProc(i).ID & ")"
                Exit Function
            End If
        Next j
        
        arrProc(i).bStatus = True
        
    End If
    
    If MD5HushYo = arrProc(i).strMD5 Then bCAppFound = True
    
Next i

If Not bCAppFound And App.LogMode = 1 Then GoTo ErrorHandler

Exit Function

ErrorHandler:
    Call MsgBox(Locale_GUI_Frase(346), vbExclamation, Locale_GUI_Frase(347))
    EndGame

End Function

Private Function pExist(ByVal pID As Long) As Boolean

Dim j As Long

For j = 1 To lngProc
    If pID = arrProc(j).ID Then
        pExist = True
        Exit Function
    End If
Next j

End Function

Private Function StrZToStr(s As String) As String

StrZToStr = Left$(s, Len(s) - 1)

End Function


