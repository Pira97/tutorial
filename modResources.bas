Attribute VB_Name = "modResources"
'*****************************************************************
'modResources.bas - v1.0.0
'
'All methods to handle resource files.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Juan Mart�n Sotuyo Dodero (juansotuyo@hotmail.com) - 10/13/2004
'   - First Release
'*****************************************************************
Option Explicit

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    intNumFiles As Integer              'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 32      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum resource_file_type
    Graphics
    Midi
    MP3
    Wav
    Scripts
    Patch
    Interface
    Maps
    Skins
End Enum

Private Const GRAPHIC_PATH As String = "\Graficos\"
Private Const MIDI_PATH As String = "\Midi\"
Private Const MP3_PATH As String = "\Mp3\"
Private Const wav_path As String = "\Wav\"
Private Const map_path As String = "\Mapas\"
Private Const INTERFACE_PATH As String = "\Interface\"
Private Const SCRIPT_PATH As String = "\Init\"
Private Const SKIN_PATH As String = "\Skins\"
Private Const PATCH_PATH As String = "\Patches\"
Private Const OUTPUT_PATH As String = "\Output\"

Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public strCurSkinName As String

Public Sub Set_Skin_Name(ByVal strSkinName As String)

strCurSkinName = strSkinName

End Sub

Public Sub Compress_Data(ByRef data() As Byte)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim loopc As Long
    
    'Get size of the uncompressed data
    Dimensions = UBound(data) + 1
    
    'Prepare a buffer 1.06 times as big as the original size
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    
    'Compress data using zlib
    Compress BufTemp(0), DimBuffer, data(0), Dimensions
    
    'Deallocate memory used by uncompressed data
    Erase data
    
    'Get rid of unused bytes in the compressed data buffer
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    'Copy the compressed data buffer to the original data array so it will return to caller
    data = BufTemp
    
    'Deallocate memory used by the temp buffer
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    data(0) = data(0) Xor data(1)
    data(2) = data(2) Xor data(3)
    data(UBound(data)) = data(UBound(data)) Xor data(4)
    
End Sub

Public Sub Decompress_Data_A(ByRef data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    data(0) = data(0) Xor 189
    
    UnCompress BufTemp(0), OrigSize, data(0), UBound(data) + 1
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Public Sub Decompress_Data_B(ByRef data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    data(0) = data(0) Xor data(1)
    data(2) = data(2) Xor data(3)
    data(UBound(data)) = data(UBound(data)) Xor data(4)
    
    UnCompress BufTemp(0), OrigSize, data(0), UBound(data) + 1
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Public Sub Encrypt_File_Header_A(ByRef FileHead As FILEHEADER)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts Normal data or turns encrypted data back to Normal
'*****************************************************************
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .intNumFiles = .intNumFiles Xor 15943
        .lngFileSize = .lngFileSize Xor 275932183
    End With
End Sub

Public Sub Encrypt_Info_Header_A(ByRef InfoHead As INFOHEADER)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts Normal data or turns encrypted data back to Normal
'*****************************************************************
    Dim EncryptedFileName As String
    Dim loopc As Long
    
    For loopc = 1 To Len(InfoHead.strFileName)
        If loopc Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr$(Asc(mid$(InfoHead.strFileName, loopc, 1)) Xor 159)
        Else
            EncryptedFileName = EncryptedFileName & Chr$(Asc(mid$(InfoHead.strFileName, loopc, 1)) Xor 96)
        End If
    Next loopc
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 37895489
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 1564645854
        .lngFileStart = .lngFileStart Xor 15997846
        .strFileName = EncryptedFileName
    End With
End Sub

Public Sub Encrypt_File_Header_B(ByRef FileHead As FILEHEADER)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts Normal data or turns encrypted data back to Normal
'*****************************************************************
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .intNumFiles = .intNumFiles Xor 29222
        .lngFileSize = .lngFileSize Xor (56732 + .intNumFiles)
    End With
End Sub

Public Sub Encrypt_Info_Header_B(ByRef InfoHead As INFOHEADER, ByRef FileHead As FILEHEADER)
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts Normal data or turns encrypted data back to Normal
'*****************************************************************
    Dim EncryptedFileName As String
    Dim loopc As Long
    
    For loopc = 1 To Len(InfoHead.strFileName)
        If loopc Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr$(Asc(mid$(InfoHead.strFileName, loopc, 1)) Xor 220)
        Else
            EncryptedFileName = EncryptedFileName & Chr$(Asc(mid$(InfoHead.strFileName, loopc, 1)) Xor 15)
        End If
    Next loopc
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 45464
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 563345
        .lngFileStart = .lngFileStart Xor 4366443
        .strFileName = EncryptedFileName
    End With
End Sub

Public Function Extract_All_Files(ByVal file_type As resource_file_type, ByVal Resource_Path As String, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Graphics
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Graficos.cao"
            Else
                SourceFilePath = Resource_Path & "\Graficos.cao"
            End If
            OutputFilePath = Resource_Path & GRAPHIC_PATH
            
        Case Midi
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Sounds.cao"
            Else
                SourceFilePath = Resource_Path & "\Sounds.cao"
            End If
            OutputFilePath = Resource_Path & MIDI_PATH
        
        Case MP3
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "MP3.cao"
            Else
                SourceFilePath = Resource_Path & "\MP3.cao"
            End If
            OutputFilePath = Resource_Path & MP3_PATH
        
        Case Wav
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Sounds.cao"
            Else
                SourceFilePath = Resource_Path & "\Sounds.cao"
            End If
            OutputFilePath = Resource_Path & wav_path
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Init.cao"
            Else
                SourceFilePath = Resource_Path & "\Init.cao"
            End If
            OutputFilePath = Resource_Path & SCRIPT_PATH
        
        Case Interface
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Interface.cao"
            Else
                SourceFilePath = Resource_Path & "\Interface.cao"
            End If
            OutputFilePath = Resource_Path & INTERFACE_PATH
        
        Case Skins
            If LenB(strCurSkinName) = 0 Then Exit Function
        
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & strCurSkinName & ".ias"
            Else
                SourceFilePath = Resource_Path & "\" & strCurSkinName & ".ias"
            End If
            OutputFilePath = Resource_Path & SKIN_PATH
            
        Case Else
            Exit Function
    End Select
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt FILEHEADER
    Encrypt_File_Header_B FileHead
        
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
        
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each INFOHEADER before accessing the data
        Encrypt_Info_Header_B InfoHead(loopc), FileHead
        
        'Check if there is enough memory
        If InfoHead(loopc).lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
            MsgBox "No hay espacio suficiente en el disco."
            Exit Function
        End If
        
        'Resize the byte data array
        ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
        
        'Decompress all data
        Decompress_Data_B SourceData, InfoHead(loopc).lngFileSizeUncompressed
        
        'Get a free handler
        handle = FreeFile
        
        'Create a new file and put in the data
        Open OutputFilePath & InfoHead(loopc).strFileName For Binary As handle
        
        Put handle, , SourceData
        
        Close handle
        
        Erase SourceData
        
        DoEvents
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    
    Extract_All_Files = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "No se logr� extraer los archivos. Motivo: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_Patch(ByVal Resource_Path As String, ByVal file_name As String) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim loopc As Long
    Dim loopc2 As Long
    Dim LoopC3 As Long
    Dim OutputFile As Integer
    Dim UpdatedFile As Integer
    Dim SourceFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim ResFileHead As FILEHEADER
    Dim ResFileHeadUnCr As FILEHEADER
    Dim ResInfoHead() As INFOHEADER
    Dim UpdatedInfoHead As INFOHEADER
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    Dim FileExtension As String
    Dim DataOffset As Long
    Dim OutputFilePath As String
    
    'Done flags
    Dim bmp_done As Boolean
    Dim wav_done As Boolean
    Dim mid_done As Boolean
    Dim mp3_done As Boolean
    Dim exe_done As Boolean
    Dim gui_done As Boolean
    Dim ind_done As Boolean
    Dim dat_done As Boolean
    Dim ini_done As Boolean
    Dim map_done As Boolean
    
    '************************************************************************************************
    'This is similar to Extract, but has some small differences to make sure what is being updated
    '************************************************************************************************
'Set up the error handler
On Local Error GoTo ErrHandler
    
    'Open the binary file
    SourceFile = FreeFile
    SourceFilePath = file_name
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    Encrypt_File_Header_B FileHead
        
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each Info Header before accessing the data
        Encrypt_Info_Header_B InfoHead(loopc), FileHead
        RequiredSpace = RequiredSpace + InfoHead(loopc).lngFileSizeUncompressed
    Next loopc
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Erase InfoHead
        MsgBox "�No hay espacio suficiente para extraer el archivo!", , "Error"
        Exit Function
    End If
    
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead())
        'Check the extension of the file
        Select Case LCase$(Right$(Trim$(InfoHead(loopc).strFileName), 3))
            Case Is = "bmp"
                If bmp_done Then GoTo EndMainLoop
                FileExtension = "bmp"
                OutputFilePath = Resource_Path & "\Graficos.cao"
                bmp_done = True
            Case Is = "mid"
                If mid_done Then GoTo EndMainLoop
                FileExtension = "mid"
                OutputFilePath = Resource_Path & "\Sounds.cao"
                mid_done = True
            Case Is = "mp3"
                If mp3_done Then GoTo EndMainLoop
                FileExtension = "mp3"
                OutputFilePath = Resource_Path & "\MP3.cao"
                mp3_done = True
            Case Is = "wav"
                If wav_done Then GoTo EndMainLoop
                FileExtension = "wav"
                OutputFilePath = Resource_Path & "\Sounds.cao"
                wav_done = True
            Case Is = "jpg"
                If gui_done Then GoTo EndMainLoop
                FileExtension = "jpg"
                OutputFilePath = Resource_Path & "\Interface.cao"
                gui_done = True
            Case Is = "ind"
                If ind_done Then GoTo EndMainLoop
                FileExtension = "ind"
                OutputFilePath = Resource_Path & "\Init.cao"
                ind_done = True
            Case Is = "dat"
                If dat_done Then GoTo EndMainLoop
                FileExtension = "dat"
                OutputFilePath = Resource_Path & "\Init.cao"
                dat_done = True
            Case Is = "ini"
                If ini_done Then GoTo EndMainLoop
                FileExtension = "ini"
                OutputFilePath = Resource_Path & "\Init.cao"
                ini_done = True
            Case Is = "csm"
                If map_done Then GoTo EndMainLoop
                FileExtension = "csm"
                OutputFilePath = Resource_Path & "\Mapas.cao"
                map_done = True
        End Select
        
        OutputFile = FreeFile
        Open OutputFilePath For Binary Access Read Lock Write As OutputFile
        
        'Get file header
        Get OutputFile, 1, ResFileHead
        
        'Desencrypt file header
        Encrypt_File_Header_B ResFileHead
        
        ResFileHeadUnCr = ResFileHead
        
        'Resize the Info Header array
        ReDim ResInfoHead(ResFileHead.intNumFiles - 1)
        
        'Load the info header
        Get OutputFile, , ResInfoHead
        
        'Desencrypt all Info Headers
        For loopc2 = 0 To UBound(ResInfoHead())
            Encrypt_Info_Header_B ResInfoHead(loopc2), ResFileHead
        Next loopc2
        
        'Check how many of the files are new, and how many are replacements
        For loopc2 = loopc To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(loopc2).strFileName), 3)) = FileExtension Then
                'Look for same name in the resource file
                For LoopC3 = 0 To UBound(ResInfoHead())
                    If ResInfoHead(LoopC3).strFileName = InfoHead(loopc2).strFileName Then
                        Exit For
                    End If
                Next LoopC3
                
                'Update the File Head
                If LoopC3 > UBound(ResInfoHead()) Then
                    'Update number of files and size
                    ResFileHead.intNumFiles = ResFileHead.intNumFiles + 1
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize + Len(InfoHead(0)) + InfoHead(loopc2).lngFileSize
                Else
                    'We substract the size of the old file and add the one of the new one
                    ResFileHead.lngFileSize = ResFileHead.lngFileSize - ResInfoHead(LoopC3).lngFileSize + InfoHead(loopc2).lngFileSize
                End If
                
                DoEvents
                
            End If
        Next loopc2
        
        'Get the offset of the compressed data
        DataOffset = CLng(ResFileHead.intNumFiles) * Len(ResInfoHead(0)) + Len(FileHead) + 1
        
        'Encrypt file Header
        Encrypt_File_Header_B ResFileHead
        
        'Now we start saving the updated file
        UpdatedFile = FreeFile
        Open OutputFilePath & "2" For Binary Access Write Lock Read As UpdatedFile
        
        'Store the filehead
        Put UpdatedFile, 1, ResFileHead
        
        'Start storing the Info Heads
        loopc2 = loopc
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While loopc2 <= UBound(InfoHead())
                If LCase$(ResInfoHead(LoopC3).strFileName) < LCase$(InfoHead(loopc2).strFileName) Then Exit Do
                If LCase$(Right$(Trim$(InfoHead(loopc2).strFileName), 3)) = FileExtension Then
                    'Copy the info head data
                    UpdatedInfoHead = InfoHead(loopc2)
                    
                    'Set the file start pos and update the offset
                    UpdatedInfoHead.lngFileStart = DataOffset
                    DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                    
                    'Encrypt the info header and save it
                    Encrypt_Info_Header_B UpdatedInfoHead, ResFileHeadUnCr
                    
                    Put UpdatedFile, , UpdatedInfoHead
                    
                    DoEvents
                    
                End If
                loopc2 = loopc2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If loopc2 Then
                If LCase$(ResInfoHead(LoopC3).strFileName) = LCase$(InfoHead(loopc2 - 1).strFileName) Then GoTo EndLoop
            End If
            
            'Copy the info head data
            UpdatedInfoHead = ResInfoHead(LoopC3)
            
            'Set the file start pos and update the offset
            UpdatedInfoHead.lngFileStart = DataOffset
            DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
            
            'Encrypt the info header and save it
            Encrypt_Info_Header_B UpdatedInfoHead, ResFileHead
            
            Put UpdatedFile, , UpdatedInfoHead
            
EndLoop:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the list we put it now
        For loopc2 = loopc2 To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(loopc2).strFileName), 3)) = FileExtension Then
                'Copy the info head data
                UpdatedInfoHead = InfoHead(loopc2)
                
                'Set the file start pos and update the offset
                UpdatedInfoHead.lngFileStart = DataOffset
                DataOffset = DataOffset + UpdatedInfoHead.lngFileSize
                
                'Encrypt the info header and save it
                Encrypt_Info_Header_B UpdatedInfoHead, ResFileHeadUnCr
                
                Put UpdatedFile, , UpdatedInfoHead
                
                DoEvents
                
            End If
        Next loopc2
        
        'Now we start adding the compressed data
        loopc2 = loopc
        For LoopC3 = 0 To UBound(ResInfoHead())
            Do While loopc2 <= UBound(InfoHead())
                If LCase$(ResInfoHead(LoopC3).strFileName) < LCase$(InfoHead(loopc2).strFileName) Then Exit Do
                If LCase$(Right$(Trim$(InfoHead(loopc2).strFileName), 3)) = FileExtension Then
                    'Get the compressed data
                    ReDim SourceData(InfoHead(loopc2).lngFileSize - 1)
                    
                    Get SourceFile, InfoHead(loopc2).lngFileStart, SourceData
                    
                    Put UpdatedFile, , SourceData
                    
                    DoEvents
                    
                End If
                loopc2 = loopc2 + 1
            Loop
            
            'If the file was replaced in the patch, we skip it
            If loopc2 Then
                If LCase$(ResInfoHead(LoopC3).strFileName) = LCase$(InfoHead(loopc2 - 1).strFileName) Then GoTo EndLoop2
            End If
            
            'Get the compressed data
            ReDim SourceData(ResInfoHead(LoopC3).lngFileSize - 1)
            
            Get OutputFile, ResInfoHead(LoopC3).lngFileStart, SourceData
            
            Put UpdatedFile, , SourceData
            
            DoEvents
            
EndLoop2:
        Next LoopC3
        
        'If there was any file in the patch that would go in the bottom of the lsit we put it now
        For loopc2 = loopc2 To UBound(InfoHead())
            If LCase$(Right$(Trim$(InfoHead(loopc2).strFileName), 3)) = FileExtension Then
                'Get the compressed data
                ReDim SourceData(InfoHead(loopc2).lngFileSize - 1)
                
                Get SourceFile, InfoHead(loopc2).lngFileStart, SourceData
                
                Put UpdatedFile, , SourceData
                
                DoEvents
                
            End If
        Next loopc2
        
        'We are done updating the file
        Close UpdatedFile
        
        'Close and delete the old resource file
        Close OutputFile
        Kill OutputFilePath
        
        'Rename the new one
        Name OutputFilePath & "2" As OutputFilePath
        
        'Deallocate the memory used by the data array
        Erase SourceData

EndMainLoop:
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    Erase ResInfoHead
    
    Extract_Patch = True

Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead

End Function

Public Function Batch_Update(ByVal Resource_Path As String, ByVal resource_name As String) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim loopc As Long
    Dim loopc2 As Long
    Dim OutputFile As Integer
    Dim UpdatedFile As Integer
    Dim SourceFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim ResFileHead As FILEHEADER
    Dim ResInfoHead() As INFOHEADER
    Dim UpdatedInfoHead() As INFOHEADER
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    Dim OutputFilePath As String
        
'Set up the error handler
On Local Error GoTo ErrHandler
    
        OutputFilePath = Resource_Path & "\" & resource_name
        
        OutputFile = FreeFile
        Open OutputFilePath For Binary Access Read Lock Write As OutputFile
        
        'Get file header
        Get OutputFile, 1, ResFileHead
        FileHead = ResFileHead
        
        'Desencrypt file header
        Encrypt_File_Header_A ResFileHead
        
        'Resize the Info Header array
        ReDim ResInfoHead(ResFileHead.intNumFiles - 1)
        ReDim UpdatedInfoHead(ResFileHead.intNumFiles - 1)
        
        'Load the info header
        Get OutputFile, , ResInfoHead
        
        'Desencrypt all Info Headers
        For loopc2 = 0 To UBound(ResInfoHead())
            Encrypt_Info_Header_A ResInfoHead(loopc2)
            UpdatedInfoHead(loopc2) = ResInfoHead(loopc2)
        Next loopc2
                        
        'Encrypt file Header
        Encrypt_File_Header_B ResFileHead
        
        'Now we start saving the updated file
        UpdatedFile = FreeFile
        Open OutputFilePath & "2" For Binary Access Write Lock Read As UpdatedFile
        
        'Store the filehead
        Put UpdatedFile, 1, ResFileHead
        
        'Info header...
        For loopc = 0 To UBound(ResInfoHead())
            'Get the compressed data
            ReDim SourceData(ResInfoHead(loopc).lngFileSize - 1)
            
            Get OutputFile, ResInfoHead(loopc).lngFileStart, SourceData
            
            Encrypt_Info_Header_B ResInfoHead(loopc), FileHead
            
            Put UpdatedFile, , ResInfoHead(loopc)
            
            DoEvents
        Next loopc
        
        'Data...
        For loopc = 0 To UBound(UpdatedInfoHead())
            'Get the compressed data
            ReDim SourceData(UpdatedInfoHead(loopc).lngFileSize - 1)
            
            Get OutputFile, UpdatedInfoHead(loopc).lngFileStart, SourceData
            
            Decompress_Data_A SourceData, UpdatedInfoHead(loopc).lngFileSizeUncompressed
            Compress_Data SourceData
            
            Put UpdatedFile, , SourceData
            
            DoEvents
        Next loopc
        
        'We are done updating the file
        Close UpdatedFile
        
        'Close and delete the old resource file
        Close OutputFile
        Kill OutputFilePath
        
        'Rename the new one
        Name OutputFilePath & "2" As OutputFilePath
        
        'Deallocate the memory used by the data array
        Erase SourceData
    
        Batch_Update = True

Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead

End Function

Public Function Compress_Files(ByVal file_type As resource_file_type, ByVal Resource_Path As String, ByVal dest_path As String) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopc As Long
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Graphics
            SourceFilePath = Resource_Path & GRAPHIC_PATH
            SourceFileExtension = ".bmp"
            OutputFilePath = dest_path & "Graficos.cao"
        
        Case Midi
            SourceFilePath = Resource_Path & MIDI_PATH
            SourceFileExtension = ".mid"
            OutputFilePath = dest_path & "Sounds.cao"
        
        Case MP3
            SourceFilePath = Resource_Path & MP3_PATH
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "MP3.cao"
        
        Case Wav
            SourceFilePath = Resource_Path & wav_path
            SourceFileExtension = ".wav"
            OutputFilePath = dest_path & "Sounds.cao"
                
        Case Scripts
            SourceFilePath = Resource_Path & SCRIPT_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Init.cao"
        
        Case Patch
            SourceFilePath = Resource_Path & PATCH_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Patch.cao"
    
        Case Interface
            SourceFilePath = Resource_Path & INTERFACE_PATH
            SourceFileExtension = ".jpg"
            OutputFilePath = dest_path & "Interface.cao"

        Case Maps
            SourceFilePath = Resource_Path & map_path
            SourceFileExtension = ".csm"
            OutputFilePath = dest_path & "Mapas.cao"
    
        Case Skins
            If LenB(strCurSkinName) = 0 Then Exit Function
            SourceFilePath = Resource_Path & SKIN_PATH
            SourceFileExtension = ".jpg"
            OutputFilePath = dest_path & strCurSkinName & ".ias"
    
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir$(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> vbNullString
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "No hay archivos con la extenci�n " & SourceFileExtension & " en " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
        
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> vbNullString Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopc = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopc) For Binary Access Read Lock Write As SourceFile
        
        'Store file name
        InfoHead(loopc).strFileName = FileNames(loopc)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        
        'Compress it
        Compress_Data SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopc).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        DoEvents
    Next loopc
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopc = 0 To FileHead.intNumFiles - 1
        InfoHead(loopc).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopc).lngFileSize
        'Once an InfoHead index is ready, we encrypt it
        Encrypt_Info_Header_B InfoHead(loopc), FileHead
    Next loopc
    
    'Encrypt the FileHeader
    Encrypt_File_Header_B FileHead
    
    '************ Write Data
    
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "No se puede crear el archivo de recursos. Motivo: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_File(ByVal file_type As resource_file_type, ByVal Resource_Path As String, ByVal file_name As String, ByVal OutputFilePath As String, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Graphics
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Graficos.cao"
            Else
                SourceFilePath = Resource_Path & "\Graficos.cao"
            End If
            
        Case Midi
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Sounds.cao"
            Else
                SourceFilePath = Resource_Path & "\Sounds.cao"
            End If
        
        Case MP3
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "MP3.cao"
            Else
                SourceFilePath = Resource_Path & "\MP3.cao"
            End If
        
        Case Wav
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Sounds.cao"
            Else
                SourceFilePath = Resource_Path & "\Sounds.cao"
            End If
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Init.cao"
            Else
                SourceFilePath = Resource_Path & "\Init.cao"
            End If
        
        Case Interface
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Interface.cao"
            Else
                SourceFilePath = Resource_Path & "\Interface.cao"
            End If
        
        Case Maps
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & "Mapas.cao"
            Else
                SourceFilePath = Resource_Path & "\Mapas.cao"
            End If
        
        Case Skins
            If LenB(strCurSkinName) = 0 Then Exit Function
            
            If UseOutputFolder Then
                SourceFilePath = Resource_Path & OUTPUT_PATH & strCurSkinName & ".ias"
            Else
                SourceFilePath = Resource_Path & "\" & strCurSkinName & ".ias"
            End If
        
        Case Else
            Exit Function
    End Select
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = vbNullString Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
        
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "Espacio insuficiente.", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get handle, InfoHead.lngFileStart, SourceData
        
    'Decompress all data
    Decompress_Data_B SourceData, InfoHead.lngFileSizeUncompressed
        
    'Close the binary file
    Close handle
    
    'Get a free handler
    handle = FreeFile
    
    Open OutputFilePath & InfoHead.strFileName For Binary As handle
    
    Put handle, 1, SourceData
     
    Close handle
    
    Erase SourceData
         
    Extract_File = True
    
Exit Function

ErrHandler:
    Close handle
    Erase SourceData

End Function

Public Function Extract_File_Ex(ByVal file_type As resource_file_type, ByVal Resource_Path As String, ByVal file_name As String, ByRef bytArr() As Byte) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Graphics
            SourceFilePath = Resource_Path & "\Graficos.cao"
            
        Case Wav
            SourceFilePath = Resource_Path & "\Sounds.cao"
        
        Case Interface
            SourceFilePath = Resource_Path & "\Interface.cao"
        
        Case Maps
            SourceFilePath = Resource_Path & "\Mapas.cao"
        
        Case Skins
            If LenB(strCurSkinName) = 0 Then Exit Function
            SourceFilePath = Resource_Path & "\" & strCurSkinName & ".ias"
        
        Case Else
            Exit Function
    End Select
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = vbNullString Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
        
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "�Espacio insuficiente!", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim bytArr(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get handle, InfoHead.lngFileStart, bytArr
    
    'Decompress all data
    Decompress_Data_B bytArr, InfoHead.lngFileSizeUncompressed
        
    'Close the binary file
    Close handle
            
    Extract_File_Ex = True
Exit Function

ErrHandler:
    Close handle
    Erase bytArr
    
End Function

Public Function Resource_File_Exists(ByVal SourceFilePath As String, ByVal file_name As String) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim loopc As Long
    Dim InfoHead As INFOHEADER
    
'Set up the error handler
On Local Error GoTo ErrHandler
        
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = vbNullString Or InfoHead.lngFileSize = 0 Then Exit Function
            
    Resource_File_Exists = True
Exit Function

ErrHandler:
    
End Function

Public Function Delete_File(ByVal file_path As String) As Boolean
'*****************************************************************
'Author: Juan Mart�n Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************
    
    On Error GoTo Error_Handler
        
    Kill file_path
    
    Delete_File = True
    
    Exit Function
    
Error_Handler:
    On Error Resume Next
    
    Dim handle As Integer
    Dim data() As Byte
    
    'We open the file to delete
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
    
    'We replace all the bytes in it with 0s
    ReDim data(LOF(handle) - 1)
    Put handle, 1, data
    
    'We close the file
    Close handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Delete_File = True
    
End Function

Private Function File_Find(ByVal resource_file_path As String, ByVal file_name As String) As INFOHEADER
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Looks for a compressed file in a resource file. Uses binary search ;)
'**************************************************************
On Error GoTo ErrHandler
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    Dim file_handler As Integer
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
    
    'Fill file name with spaces for compatibility
    If Len(file_name) < Len(info_head.strFileName) Then _
        file_name = file_name & Space$(Len(info_head.strFileName) - Len(file_name))
    
    'Open resource file
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
    
    'Get file head
    Get file_handler, 1, file_head
    Encrypt_File_Header_B file_head
    
    min = 1
    max = file_head.intNumFiles
    
    Do While min <= max
        mid = (min + max) / 2
        
        'Get the info header of the appropiate compressed file
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
        Encrypt_Info_Header_B info_head, file_head
                
        If file_name < info_head.strFileName Then
            max = mid - 1
        ElseIf file_name > info_head.strFileName Then
            min = mid + 1
        Else
            'Copy info head
            File_Find = info_head
            
            'Close file and exit
            Close file_handler
            Exit Function
        End If
    Loop
    
ErrHandler:
    'Close file
    Close file_handler
    File_Find.strFileName = vbNullString
    File_Find.lngFileSize = 0
End Function


Public Function Extract_BMP_Memory2(ByVal file_name As String, SourceData() As Byte) As Boolean
Dim loopc As Long
Dim InfoHead As INFOHEADER
Dim handle As Integer

'Set up the error handler
On Local Error GoTo ErrHandler

    'Find the Info Head of the desired file
    InfoHead = File_Find(Resource_Path & "Graficos.cao", file_name)

    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open Resource_Path & "Graficos.cao" For Binary Access Read Lock Write As handle
    
            'Make sure there is enough space in the HD
        If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
            Close handle
            MsgBox "�Espacio insuficiente!", , "Error"
            Exit Function
        End If
    
        'Resize the byte data array
        ReDim SourceData(InfoHead.lngFileSize - 1)
        
        'Get the data
        Get handle, InfoHead.lngFileStart, SourceData



    'Decompress all data
    Decompress_Data_B SourceData, InfoHead.lngFileSizeUncompressed
    

    'Close the binary file
    Close handle

    Extract_BMP_Memory2 = True
Exit Function

ErrHandler:
    Close handle
    Erase SourceData
End Function
