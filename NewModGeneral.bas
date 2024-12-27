Attribute VB_Name = "NewModGeneral"
Option Explicit
'Seguridad
Public MacAdress        As String
Public HDserial         As Long
'Nueva seguridad
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'get mac adress

 
' // Inicializaciones, Launcher, InGame
Public Resource_Path As String
Public DirSounds     As String
Public DirMidi       As String
Public DirInit       As String


Public Win2kXP As Boolean
Public Musica        As Boolean
Public SonidoHabilitado  As Boolean
Public Efectos3D As Boolean

Public VSYNC     As Boolean
Public RecordarCuentaIni As Boolean
Public UserAccountRecorded As String
Public RunWindowed As Boolean

Public DeviceIndex As Byte

Public Pixels As Byte

Public FXVolume As Long
Public MusicVolume As Long
 

Public NombreSkin As String

Public NickModerno As Boolean
Public NombresModernos As Byte

Public PrimeraVez    As Boolean
Public Desvanecimiento As Boolean
 
 
Public CursorHabilitado As Boolean
Public MouseSpeed As Byte
Public BloqueoAlCaminar As Boolean
Public CantidadEnMacros As Boolean

'AutoUsar
Public AutoUsarActivado As Boolean
Public accionMouseUno As Byte
Public accionMousedos As Byte
Public MouseUno As String
Public CargamosMouse As Byte
Public MouseDos As String
Public CargamosMouseDos As Byte
Public Rendimiento As Byte
'Fin Autousar

Public HabilitarMensajesGlobales As Boolean

'End Configuraciones Cliente'//

'Consola transparente
Private Const GWL_EXSTYLE As Long = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_EX_TRANSPARENT As Long = &H20&
'Move_Drag
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

''
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const LWA_ALPHA = &H2
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000

'XP
'To get OS version
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      ' Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32_NT As Long = 2&

Private Declare Function GetOSVersion Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private OSInfo As OSVERSIONINFO

'¿Show cursor?
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long
'Loading pictures from byte arrays
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Function General_Get_Line_Count(ByVal FileName As String) As Long
On Error GoTo ErrorHandler
    Dim N As Integer, tmpStr As String
    If LenB(FileName) Then
        N = FreeFile()
        
        Open FileName For Input As #N
            Do While Not EOF(N)
                General_Get_Line_Count = General_Get_Line_Count + 1
                Line Input #N, tmpStr
            Loop
        Close N
    End If
    Exit Function

ErrorHandler:
    Resume Next
    
End Function

Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function General_Field_Count(ByVal Text As String, ByVal delimiter As Byte) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Count the number of fields in a delimited string
'*****************************************************************
    'If string is empty there aren't any fields
    If Len(Text) = 0 Then
        Exit Function
    End If

    Dim i As Long
    Dim FieldNum As Long
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
        End If
    Next i
    General_Field_Count = FieldNum + 1
End Function

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = vbNullString Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Sub General_Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, File
End Sub

Public Function General_Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = vbNullString
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        General_Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        General_Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function


Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal first As Long, ByVal last As Long)
'**************************************************************
'Author: juan Martín Sotuyo Dodero
'Last Modify Date: 3/03/2005
'Good old QuickSort algorithm :)
'**************************************************************
    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
    
    Low = first
    High = last
    List_Separator = SortArray((first + last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If first < High Then General_Quick_Sort SortArray, first, High
    If Low < last Then General_Quick_Sort SortArray, Low, last
End Sub

Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

Public Function General_Load_Picture_From_Resource(ByVal picture_file_name As String) As IPicture
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Loads a picture from a resource file and returns it
'**************************************************************

'On Error GoTo ErrorHandler

picture_file_name = picture_file_name & ".jpg"

If Extract_File(Interface, App.Path & "\Recursos\", picture_file_name, Resource_Path, False) Then
    Set General_Load_Picture_From_Resource = LoadPicture(Resource_Path & picture_file_name)
    Call Delete_File(Resource_Path & picture_file_name)
Else
    Set General_Load_Picture_From_Resource = Nothing
End If

Exit Function

ErrorHandler:
    If General_File_Exists(Resource_Path & picture_file_name, vbNormal) Then
        Call Delete_File(Resource_Path & picture_file_name)
    End If

End Function

Public Function General_Load_Picture_From_Resource_Ex(ByVal picture_file_name As String) As IPicture
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 2/2/2006
'Loads a picture from a resource file loaded in memory and returns it
'**************************************************************

On Error GoTo ErrorHandler

Dim bytArr() As Byte
    
picture_file_name = picture_file_name & ".jpg"
    
If Extract_File_Ex(Interface, App.Path & "\Recursos\", picture_file_name, bytArr()) Then
    Set General_Load_Picture_From_Resource_Ex = General_Load_Picture_From_BArray(bytArr())
Else
    Set General_Load_Picture_From_Resource_Ex = Nothing
End If

Exit Function

ErrorHandler:

End Function

Public Function General_Load_Picture_From_BArray(ByRef bytArr() As Byte) As IPicture
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 2/2/2006
'Loads a picture from a byte array
'**************************************************************

On Error GoTo ErrorHandler

Dim LowerBound As Long
Dim ByteCount As Long
Dim hMem As Long
Dim lpMem As Long
Dim IID_IPicture(15) As Long
Dim istm As stdole.IUnknown
    
LowerBound = LBound(bytArr)
ByteCount = (UBound(bytArr) - LowerBound) + 1
hMem = GlobalAlloc(&H2, ByteCount)
If hMem <> 0 Then
    lpMem = GlobalLock(hMem)
    If lpMem <> 0 Then
        MoveMemory ByVal lpMem, bytArr(LowerBound), ByteCount
        Call GlobalUnlock(hMem)
        If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
            If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), General_Load_Picture_From_BArray)
            End If
        End If
    End If
End If

Exit Function

ErrorHandler:

End Function

Public Function General_Load_Skin_Picture_From_Resource_Ex(ByVal picture_file_name As String) As IPicture
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 2/2/2006
'Loads a picture from a resource skin file loaded in memory and returns it
'**************************************************************

On Error GoTo ErrorHandler

Dim bytArr() As Byte

picture_file_name = picture_file_name & ".jpg"

If Extract_File_Ex(Skins, App.Path & "\Skins\", picture_file_name, bytArr()) Then
    Set General_Load_Skin_Picture_From_Resource_Ex = General_Load_Picture_From_BArray(bytArr)
Else
    If Extract_File_Ex(Interface, App.Path & "\Recursos\", picture_file_name, bytArr()) Then
        Set General_Load_Skin_Picture_From_Resource_Ex = General_Load_Picture_From_BArray(bytArr)
    Else
        Set General_Load_Skin_Picture_From_Resource_Ex = Nothing
    End If
End If

Exit Function

ErrorHandler:

End Function

Public Function General_Get_Skin_Author() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'**************************************************************
        
If Extract_File(Skins, App.Path & "\Skins\", "autor.jpg", Resource_Path) Then
    General_Get_Skin_Author = Trim$(General_Get_File_Contents(Resource_Path & "autor.jpg"))
    Delete_File Resource_Path & "autor.jpg"
Else
    General_Get_Skin_Author = "Desconocido"
End If

End Function

Public Function General_Get_File_Contents(ByVal FileName As String) As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
On Error GoTo ErrorHandler

Dim N As Integer, tmpStr As String

If LenB(FileName) Then
    N = FreeFile()
    
    Open FileName For Input As #N
    
    Do While Not EOF(N)
        Line Input #N, tmpStr
        General_Get_File_Contents = General_Get_File_Contents & tmpStr
    Loop
    
    Close N

End If

Exit Function

ErrorHandler:

End Function

'I miss old void functions :( I dont want to call em subsss
Public Sub General_Cursor_Render(ByVal bRender As Boolean)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: Unknown
'
'**************************************************************

Call ShowCursor(IIf(bRender, 1, 0))
End Sub

Public Sub MensajeAdvertencia(ByVal Mensaje As String)
Call MsgBox(Mensaje, vbInformation + vbOKOnly, Locale_GUI_Frase(351))
End Sub

Public Function SuperMid(ByVal strMain As String, str1 As String, str2 As String, Optional reverse As Boolean) As String

    'DESCRIPTION: Extract the portion of a string between the two substrings defined in str1 and str2.
    'DEVELOPER: Ryan Wells (wellsr.com)
    'HOW TO USE: - Pass the argument your main string and the 2 strings you want to find in the main string.
    ' - This function will extract the values between the end of your first string and the beginning
    ' of your next string.
    ' - If the optional boolean "reverse" is true, an InStrRev search will occur to find the last
    ' instance of the substrings in your main string.
    Dim i As Integer, J As Integer, temp As Variant

    On Error GoTo ErrHandler:

    If reverse = True Then
        i = InStrRev(strMain, str1)
        J = InStrRev(strMain, str2)

        If Abs(J - i) < Len(str1) Then J = InStrRev(strMain, str2, i)
        If i = J Then 'try to search 2nd half of string for unique match
            J = InStrRev(strMain, str2, i - 1)

        End If

    Else
        i = InStr(1, strMain, str1)
        J = InStr(1, strMain, str2)

        If Abs(J - i) < Len(str1) Then J = InStr(i + Len(str1), strMain, str2)
        If i = J Then 'try to search 2nd half of string for unique match
            J = InStr(i + 1, strMain, str2)

        End If

    End If

    If i = 0 And J = 0 Then GoTo ErrHandler:
    If J = 0 Then J = Len(strMain) + Len(str2) 'just to make it arbitrarily large
    If i = 0 Then i = Len(strMain) + Len(str1) 'just to make it arbitrarily large
    If i > J And J <> 0 Then 'swap order
        temp = J
        J = i
        i = temp
        temp = str2
        str2 = str1
        str1 = temp

    End If

    i = i + Len(str1)
    SuperMid = mid(strMain, i, J - i)
    Exit Function
ErrHandler:
    SuperMid = "A"

    'MsgBox "Error extracting strings. Check your input" & vbNewLine & vbNewLine & "Aborting", , "Strings not found"
End Function

Public Function General_Windows_Is_2000XP() As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: Unknown
'
'**************************************************************

On Error GoTo ErrorHandler

Dim RetVal As Long

OSInfo.dwOSVersionInfoSize = Len(OSInfo)
RetVal = GetOSVersion(OSInfo)

If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT And OSInfo.dwMajorVersion >= 5 Then
    General_Windows_Is_2000XP = True
Else
    General_Windows_Is_2000XP = False
End If

Exit Function

ErrorHandler:
    General_Windows_Is_2000XP = False

End Function
Public Sub Make_Transparent_Richtext(ByVal hwnd As Long)

If Win2kXP Then Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub

Public Sub Make_Transparent_Form(ByVal hwnd As Long, Optional ByVal bytOpacity As Byte = 128)

If Win2kXP Then
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwnd, 0, bytOpacity, LWA_ALPHA)
End If

End Sub

Public Sub UnMake_Transparent_Form(ByVal hwnd As Long)

If Win2kXP Then Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And (Not WS_EX_TRANSPARENT))

End Sub


 ''
' Removes all text from the console and dialogs
Public Sub Auto_Drag(ByVal hwnd As Long)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub
 



Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
    
    On Error GoTo GetDriveSerialNumber_Err
    

    '***************************************************
    'Author: Nahuel Casas (Zagen)
    'Last Modify Date: 07/12/2009
    ' 07/12/2009: Zagen - Convertì las funciones, en formulas mas fàciles de modificar.
    '***************************************************
    

    Dim fso As Object, Drv As Object, DriveSerial As Long
         
    'Creamos el objeto FileSystemObject.
    Set fso = CreateObject("Scripting.FileSystemObject")
         
    'Asignamos el driver principal.
    If DriveLetter <> "" Then
        Set Drv = fso.GetDrive(DriveLetter)
    Else
        Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))

    End If
     
    With Drv

        If .IsReady Then
            DriveSerial = Abs(.SerialNumber)
        Else    '"Si el driver no està como para empezar ..."
            DriveSerial = -1

        End If

    End With
         
    'Borramos y limpiamos.
    Set Drv = Nothing
    Set fso = Nothing
    'Seteamos :)
    GetDriveSerialNumber = DriveSerial
         
    
    Exit Function

GetDriveSerialNumber_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLadder.GetDriveSerialNumber", Erl)
    Resume Next
    
End Function

Public Function GetMacAddress() As String
    
    On Error GoTo GetMacAddress_Err
    

    Const OFFSET_LENGTH As Long = 400

    Dim lSize           As Long

    Dim baBuffer()      As Byte

    Dim lIdx            As Long

    Dim sRetVal         As String
    
    Call GetAdaptersInfo(ByVal 0, lSize)

    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)

        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next

    End If

    GetMacAddress = sRetVal

    
    Exit Function

GetMacAddress_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLadder.GetMacAddress", Erl)
    Resume Next
    
End Function

