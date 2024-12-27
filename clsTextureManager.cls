VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextureManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit
 
Private mFreeMemoryBytes As Long

Private Const BYTES_PER_MB          As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE     As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE               As Long = 337

Private Type SURFACE_ENTRY_DYN

    FileIndex As Long
    LastAccess As Long
    Surface As Direct3DTexture8
    size As Long
    Width As Integer
    Height As Integer

End Type

Private Type HashNode

    SurfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN

End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode
Private DirectD3D                        As D3DX8
Private maxBytesToUse                    As Long
Private usedBytes                        As Long
Private ResourcePath                     As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE

End Sub

Private Sub Class_Terminate()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Clean up
    '**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1

        With surfaceList(i)

            For j = 1 To .SurfaceCount
                Set .SurfaceEntry(j).Surface = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry

        End With

    Next i

End Sub

Public Sub Initialize(ByRef DD As D3DX8, _
                      ByVal graphicPath As String, _
                      Optional ByVal maxMemoryUsageInMb As Long = -1)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 29/07/2012 - ^[GS]^
    'Initializes the manager
    '**************************************************************
    Set DirectD3D = DD
    
    ResourcePath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB

    End If

End Sub

Public Property Get Surface(ByVal FileIndex As Long, ByRef Width As Integer, ByRef Height As Integer) As Direct3DTexture8

    '**************************************************************
    'Retrieves the requested texture
    '**************************************************************
    Dim i As Long
    
    ' Search the index on the list
    With surfaceList(FileIndex Mod HASH_TABLE_SIZE)

        For i = 1 To .SurfaceCount

            If .SurfaceEntry(i).FileIndex = FileIndex Then
                .SurfaceEntry(i).LastAccess = GetTickCount
                Width = .SurfaceEntry(i).Width
                Height = .SurfaceEntry(i).Height
                Set Surface = .SurfaceEntry(i).Surface
                Exit Property

            End If
        Next i

    End With
    
    'Not in memory, load it!
    Set Surface = LoadSurface(FileIndex, Width, Height)

End Property
Private Function LoadSurface(ByVal FileIndex As Long, _
                             ByRef Width As Integer, _
                             ByRef Height As Integer) As Direct3DTexture8
    '**************************************************************
    'Loads the surface named fileIndex + ".bmp" and inserts it to the
    'surface list in the listIndex position
    '**************************************************************
    
    On Error GoTo ErrHandler

    Dim newSurface   As SURFACE_ENTRY_DYN
    Dim surface_desc As D3DSURFACE_DESC, bytArr() As Byte
    Dim DirPath      As String
    Dim bufferB() As Byte
    
    If FileIndex = 0 Then Exit Function

    DirPath = App.Path & "\Recursos\Graficos\" & CStr(FileIndex) & ".BMP"

    With newSurface
  
        .FileIndex = FileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .LastAccess = GetTickCount
        Dim file As String
        file = App.Path & "\" & FileIndex & ".bmp"
        
            If Extract_BMP_Memory2(FileIndex & ".bmp", bufferB) = False Then
                Debug.Print FileIndex & " no existe"
                Exit Function
            End If
            
        Extract_File_Ex Graphics, App.Path & "\Recursos\", LTrim$(CStr(FileIndex)) & ".bmp", bytArr
        
        'Set .Surface = DirectD3D.CreateTextureFromFileEx(DirectDevice, DirPath, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
            Set .Surface = DirectD3D.CreateTextureFromFileInMemoryEx(DirectDevice, bufferB(0), UBound(bufferB) + 1, _
                D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, surface_desc, ByVal 0)
                
        .Surface.GetLevelDesc 0, surface_desc
        
        Width = surface_desc.Width
        Height = surface_desc.Height
        .size = surface_desc.size
        
        .Width = Width
        .Height = Height
        
    End With
    
    'Insert surface to the list
    With surfaceList(FileIndex Mod HASH_TABLE_SIZE)
        .SurfaceCount = .SurfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .SurfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.SurfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface

    End With
    
    'Update used bytes
    usedBytes = usedBytes + surface_desc.size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse

        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do

        End If

    Loop
    
    Exit Function

ErrHandler:

End Function
Private Function RemoveLRU() As Boolean

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Removes the Least Recently Used surface to make some room for new ones
    '**************************************************************
    Dim LRUi    As Long

    Dim LRUj    As Long

    Dim LRUtime As Long

    Dim i       As Long

    Dim j       As Long

    Dim ddsd    As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1

        With surfaceList(i)

            For j = 1 To .SurfaceCount

                If LRUtime > .SurfaceEntry(j).LastAccess Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).LastAccess

                End If

            Next j

        End With

    Next i
    
    If LRUj Then
        RemoveLRU = True
        'Retrieve the surface desc
        Call surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetLevelDesc(0, ddsd)
        
        'Remove it
        Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
        surfaceList(LRUi).SurfaceEntry(LRUj).FileIndex = 0
        
        'Move back the list (if necessary)
        With surfaceList(LRUi)

            For j = LRUj To .SurfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .SurfaceCount = .SurfaceCount - 1

            If .SurfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .SurfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry

            End If

        End With
        
        'Update the used bytes
        usedBytes = usedBytes - ddsd.size

    End If

End Function



