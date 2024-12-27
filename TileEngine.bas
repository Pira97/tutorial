Attribute VB_Name = "Mod_TileEngine"
Option Explicit
                                   
'Const tamaños
Private NickModernoXX As Byte
Private NickModernoX As Byte
                         
'Constantes graficos
Private Estrella As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

Private LastInvRender        As Long
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
Private AlphaY As Integer 'Techos
Public MSRender As Integer
'Mermas, renderizamos nombres de mapas.
Private map_letter_grh As Grh
Private map_letter_grh_next As Long
Private map_letter_a As Single
Private map_letter_fadestatus As Byte
Private Type aura
    Grh As Integer          ' GrhIndex
    Rotation As Byte        ' Rotate or Not
    angle As Single         ' Angle
    speed As Single         ' Speed
    tickCount As Long       ' TickCount from Speed Controls
    color(0 To 3) As Long   ' Color
    OffSetX As Integer      ' PixelOffset X
    OffSetY As Integer      ' PixelOffset Y

End Type

Private auras()       As aura ' List of Aura's
 
Private Type Light

    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    color As Long
    range As Byte

End Type

Private Light_List()        As Light
Private Light_Count         As Long
Private Light_last          As Long

Public Const ScreenWidth    As Long = 544 'Keep this identical to the value on the server!
Public Const ScreenHeight   As Long = 416 'Keep this identical to the value on the server!

Public Const PI             As Single = 3.14159
Public Const DegreeToRadian As Single = PI / 180

'Major DX Objects
Public DirectX              As DirectX8
Public DirectD3D            As Direct3D8
Public DirectDevice         As Direct3DDevice8
Public DirectD3D8           As D3DX8
Public DispMode             As D3DDISPLAYMODE
Private D3DWindow           As D3DPRESENT_PARAMETERS
Private DirectD3Dcaps       As D3DCAPS8

Private Projection          As D3DMATRIX
Private View                As D3DMATRIX

Private MainViewRect        As D3DRECT

Private Const FVF           As Long = D3DFVF_XYZ Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
 
        
Private Type tStructureLng

    X As Long
    Y As Long

End Type
        
Public Type CharVA

    X As Integer
    Y As Integer
    w As Integer
    h As Integer
    
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single

End Type

Private Type VFH

    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA

End Type

Public Type CustomFont

    HeaderInfo As VFH            'Holds the header information
    Texture As Direct3DTexture8  'Holds the texture of the text
    RowPitch As Integer          'Number of characters per row
    RowFactor As Single          'Percentage of the texture width each character takes
    ColFactor As Single          'Percentage of the texture height each character takes
    CharHeight As Byte           'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As tStructureLng 'Size of the texture

End Type

Private Texture              As clsTextureManager
Public SpriteBatch          As clsBatch
Public cfonts                As CustomFont

 
Public White(0 To 3)         As Long
Public red(0 To 3)           As Long
Public Cyan(0 To 3)          As Long
Public Black(0 To 3)         As Long
Public FaintBlack(0 To 3)    As Long
Public Yellow(0 To 3)        As Long
Public Gray(0 To 3)          As Long
Public transparent(0 To 3)   As Long
Public green(0 To 3)         As Long
Public blue(0 To 3)          As Long

Public EngineRun             As Boolean

'Map sizes in tiles
Public Const XMaxMapSize     As Byte = 100
Public Const XMinMapSize     As Byte = 1
Public Const YMaxMapSize     As Byte = 100
Public Const YMinMapSize     As Byte = 1

Private Const GrhFogata      As Integer = 1521


'Posicion en un mapa

Public Type Position

    X As Long
    Y As Long

End Type

'Posicion en el Mundo

Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion

Public Type GrhData

    sX As Integer
    sY As Integer
       
    FileNum As Integer
       
    pixelWidth As Integer
    pixelHeight As Integer
       
    TileWidth As Single
    TileHeight As Single
       
    NumFrames As Integer
    Frames() As Integer
       
    speed As Single
    mini_map_color As Long

End Type

'apunta a una estructura grhdata y mantiene la animacion

Public Type Grh

    GrhIndex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos

Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas

Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas

Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de los escudos

Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Apariencia del personaje

Public Type char

    EsGM As Boolean
    
    'Particula
    Particula As Byte
    ParticulaTime As Long
    
    'Auras
    Arma_Aura As Byte
    Body_Aura As Byte
    Escudo_Aura As Byte
    Head_Aura As Byte
    Otra_Aura As Byte
    Anillo_Aura As Byte
    State As Byte
    active As Byte
    Heading As E_Heading
    Pos As Position
    donador As Byte
    iHead As Integer
    iBody As Integer
    body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fX As Grh
    FxIndex As Integer
    'Shermie80 Alpha-
    AlphaX As Double
    NPCHostil As Boolean
    EsNPC As Boolean
    EsUsuario As Boolean
    last_tick As Long
    
    '----------------
    
    Nombre As String
    dl As Boolean
    dialog() As String
    dialogColor(3) As Long
    dialogLife As Long
    dialogStart As Long
    dialogHeight As Single
    dialogIndex As Byte
    
    OffSetNombre As Integer
    Clan As String
    OffSetClan As Integer
    
    color(3) As Long
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    LastStep As Long
    pie As Boolean
    Muerto As Boolean
    Invisible As Boolean
    priv As Byte
    
    particle_count As Integer
    particle_group() As Long
    
    aura(1 To 6) As aura

End Type

'Info de un objeto

Public Type Obj

    OBJIndex As Integer
    Amount As Integer
    tipe As Byte
    EsFijo As Byte
End Type


Public Type Fantasma
    Nombre As String
    Activo As Boolean
    body As Grh
    Head As Grh
    Arma As Grh
    Casco As Grh
    Escudo As Grh
    
    AlphaB As Single
    OffX As Integer
    Offy As Integer
    Heading As Byte
    color(3) As Long
    donador As Byte
    
    'Auras
    Arma_Aura As Byte
    Body_Aura As Byte
    Escudo_Aura As Byte
    Head_Aura As Byte
    Otra_Aura As Byte
    Anillo_Aura As Byte
    
    EsUsuario As Byte
    Clan As String
    OffSetClan As Integer
    X As Integer
    Y As Integer
End Type

'Tipo de las celdas del mapa

Public Type MapBlock
    FXArea As Grh
    FxIndex As Integer
    fX As Grh
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    CharFantasma As Fantasma
    particle_group As Integer
    light_value(3) As Long
End Type

'Info de cada mapa

Public Type MapInfo

    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer

End Type

 ' $ Sistema de Cuentas $

Public Type PjCuenta
    Nombre      As String
    Head        As Integer
    body        As Integer
    Shield      As Byte
    Weapon      As Byte
    Nivel       As Byte
    Mapa        As Integer
    Clase       As Byte
    color       As Byte
    Helmet      As Byte
    GameMaster  As Boolean
End Type

Public cPJ() As PjCuenta

Public Type RenderSkin
    Head        As Integer
    body        As Integer
    Shield      As Byte
    Casco       As Byte
    Weapon      As Byte
    Helmet      As Byte
End Type

Public RSkin As RenderSkin

' $     Shermie80    $

'Bordes del mapa
Public MinXBorder             As Byte
Public MaxXBorder             As Byte
Public MinYBorder             As Byte
Public MaxYBorder             As Byte

'Status del user
Public UserIndex              As Integer
Public UserMoving             As Byte
Public UserBody               As Integer
Public UserHead               As Integer
Public UserPos                As Position 'Posicion

Public AddtoUserPos           As Position 'Si se mueve


Public fps                    As Long
Public FramesPerSecCounter    As Long
Private lFrameTimer            As Long
Public FrameTime          As Long


'Tamaño del la vista en Tiles
Private WindowTileWidth       As Integer
Private WindowTileHeight      As Integer

Public HalfWindowTileWidth   As Integer
Public HalfWindowTileHeight  As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize         As Byte

'Tamaño de los tiles en pixels
Public TilePixelHeight        As Integer
Public TilePixelWidth         As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public scroll_pixels_per_frame  As Single
Public Const scroll_pixels_per_frameBackUp   As Single = 8.5
Public Const velocidadMontando   As Single = 1.2

Public Const VelocidadMuerto  As Single = 1.2

Private OffSetCounterX        As Single
Private OffSetCounterY        As Single

Public timerElapsedTime           As Single
Public timerTicksPerFrame         As Single
Public engineBaseSpeed            As Single

Public NumBodies              As Integer
Public Numheads               As Integer
Public NumFxs                 As Integer
Public NumChars               As Integer
Public NumWeaponAnims         As Integer
Public NumEscudosAnims        As Integer
Public NumShieldAnims         As Integer
Public LastChar               As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData()              As GrhData
Public BodyData()             As BodyData
Public HeadData()             As HeadData
Public FxData()               As tIndiceFx
Public WeaponAnimData()       As WeaponAnimData
Public ShieldAnimData()       As ShieldAnimData
Public CascoAnimData()        As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
 

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()              As MapBlock ' Mapa
Public MapInfo                As MapInfo  ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public bRain        As Boolean 'está raineando?

Public bTecho                 As Boolean 'hay techo?

Public charlist(1 To 10000)   As char
 
'Timer con proc
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


'*********************************************************************************
'                                  APIS                                          '
'*********************************************************************************

'Shermie80; agrego esto para usar en el renderizado de los personajes de la cuenta
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, _
                                                       ByVal nXOriginDest As Long, _
                                                       ByVal nYOriginDest As Long, _
                                                       ByVal nWidthDest As Long, _
                                                       ByVal nHeightDest As Long, _
                                                       ByVal hdcsrc As Long, _
                                                       ByVal nXOriginSrc As Long, _
                                                       ByVal nYOriginSrc As Long, _
                                                       ByVal nWidthSrc As Long, _
                                                       ByVal nHeightSrc As Long, _
                                                       ByVal crTransparent As Long) As Long

' $

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (destination As Any, _
                                       source As Any, _
                                       ByVal Length As Long)

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

Public Sub SetCharacterAura(ByVal charindex As Integer, _
                            ByVal AuraIndex As Byte, _
                            ByVal Slot As Byte)
    '***************************************************
    'Author: Standelf
    'Last Modify Date: 27/05/2010
    '***************************************************
    
    If Slot <= 0 Or Slot >= 7 Then Exit Sub
    Call Set_Aura(charindex, Slot, AuraIndex)

End Sub
    
Private Function CargarAuras() As Boolean
    On Error Resume Next
    
    Dim i        As Long, AurasTotales As Integer
    Dim pathFile As String
    Dim ColorSet As Long, TempSet As String
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "auras.ini", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo errorhandler
    End If

    pathFile = Resource_Path & "auras.ini"
    AurasTotales = val(GetVar(pathFile, "Auras", "NumAuras"))
    
    ReDim auras(1 To AurasTotales) As aura
    
    For i = 1 To AurasTotales

        With auras(i)
            .Grh = val(GetVar(pathFile, i, "GrhIndex"))
            .Rotation = val(GetVar(pathFile, i, "Rotate"))
            .angle = 0
            .speed = val(GetVar(pathFile, i, "Speed"))
                
            .OffSetX = val(GetVar(pathFile, i, "OffsetX"))
            .OffSetY = val(GetVar(pathFile, i, "OffsetY"))
            
            For ColorSet = 0 To 3
                TempSet = GetVar(pathFile, val(i), "Color" & ColorSet)
                .color(ColorSet) = D3DColorXRGB(ReadField(1, TempSet, Asc(",")), ReadField(2, TempSet, Asc(",")), ReadField(3, TempSet, Asc(",")))
            Next ColorSet
                
            .tickCount = 0

        End With

    Next i
    
    Delete_File Resource_Path & "auras.ini"
    CargarAuras = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        If General_File_Exists(Resource_Path & "auras.ini", vbNormal) Then Delete_File Resource_Path & "auras.ini"
        MsgBox "Error while loading weapon index: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error!"
    End If
    
End Function

Private Sub Set_Aura(ByVal charindex As Integer, _
                     ByVal Slot As Byte, _
                     ByVal aura As Byte)

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Set Aura to Char
    '***************************************************
    If Slot <= 0 Or Slot >= 7 Then Exit Sub
    
    If aura = 0 Then Exit Sub
    
    With charlist(charindex).aura(Slot)
    
        .Grh = auras(aura).Grh
        .angle = auras(aura).angle
        .Rotation = auras(aura).Rotation
        .speed = auras(aura).speed
        
        .OffSetX = auras(aura).OffSetX
        .OffSetY = auras(aura).OffSetY
        
        .color(0) = auras(aura).color(0)
        .color(1) = auras(aura).color(1)
        .color(2) = auras(aura).color(2)
        .color(3) = auras(aura).color(3)
        
        .tickCount = GetTickCount

    End With

End Sub

Public Sub Delete_All_Auras(ByVal charindex As Integer)
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Kill all of aura´s from Char
    '***************************************************
    
    Delete_Aura charindex, 1
    Delete_Aura charindex, 2
    Delete_Aura charindex, 3
    Delete_Aura charindex, 4
    Delete_Aura charindex, 5
    Delete_Aura charindex, 6

End Sub
    
Public Sub Delete_Aura(ByVal charindex As Integer, _
                       ByVal Slot As Byte)

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 26/05/10
    'Kill Aura from Char
    '***************************************************
    If Slot <= 0 Or Slot >= 7 Then Exit Sub
    
    charlist(charindex).aura(Slot) = auras(1)    '1 = Fake Aura

End Sub

 
 
Private Function CargarCabezas() As Boolean
    On Error Resume Next
    
    Dim N            As Integer
    Dim i            As Long
    Dim Numheads     As Integer
    Dim Miscabezas() As tIndiceCabeza
    N = FreeFile()
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "cabezas.ind", Resource_Path, False) Then
        Err.Description = "No se ha logrado extraer el archivo de recurso."
        GoTo errorhandler
    End If

    Open Resource_Path & "cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If
    Next i
    
    Close #N
Delete_File Resource_Path & "cabezas.ind"
CargarCabezas = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        Close #N
        If General_File_Exists(Resource_Path & "cabezas.ind", vbNormal) Then Delete_File Resource_Path & "cabezas.ind"
        MsgBox "Error while loading head index: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error!"
    End If
    
End Function
Private Function CargarCascos() As Boolean

    On Error GoTo errorhandler
    
    Dim N            As Integer
    Dim i            As Long
    Dim NumCascos    As Integer
    Dim Miscabezas() As tIndiceCabeza

    N = FreeFile()
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "cascos.ind", Resource_Path, False) Then
        Err.Description = "No se ha logrado extraer el archivo de recurso."
        GoTo errorhandler
    End If

    Open Resource_Path & "cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If
    DoEvents
    Next i
    
    Close #N

Delete_File Resource_Path & "cascos.ind"
CargarCascos = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        Close #N
        If General_File_Exists(Resource_Path & "cascos.ind", vbNormal) Then Delete_File Resource_Path & "cascos.ind"
        MsgBox "Error while loading helmet index: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error!"
    End If
    
End Function

Private Function CargarCuerpos() As Boolean

    On Error GoTo errorhandler

    Dim N            As Integer
    Dim i            As Long
    Dim NumCuerpos   As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    N = FreeFile()
 
    If Not Extract_File(Scripts, App.Path & "\Recursos", "personajes.ind", Resource_Path, False) Then
        Err.Description = "No se ha logrado extraer el archivo de recurso."
        GoTo errorhandler
    End If

    Open Resource_Path & "personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

        End If

    Next i
    
    Close #N
    
Delete_File Resource_Path & "personajes.ind"
CargarCuerpos = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        If General_File_Exists(Resource_Path & "personajes.ind", vbNormal) Then Delete_File Resource_Path & "personajes.ind"
        MsgBox "Error while loading body index: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error!"
    End If
    
End Function
Private Function CargarFxs() As Boolean

    On Error Resume Next
    
    Dim N      As Integer

    Dim i      As Long

    Dim NumFxs As Integer
    
     N = FreeFile()
     
     If Not Extract_File(Scripts, App.Path & "\Recursos", "fxs.ind", Resource_Path, False) Then
     Err.Description = "No se ha logrado extraer el archivo de recurso."
     GoTo errorhandler
     End If

    Open Resource_Path & "fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    DoEvents
    Next i
    
    Close #N

    Delete_File Resource_Path & "fxs.ind"
    CargarFxs = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        Close #N
        If General_File_Exists(Resource_Path & "fxs.ind", vbNormal) Then Delete_File Resource_Path & "fxs.ind"
    End If
End Function

Private Function CargarAnimArmas() As Boolean
    On Error Resume Next


    Dim loopc As Long
    Dim Arch  As String
 
    If Not Extract_File(Scripts, App.Path & "\Recursos", "armas.dat", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo errorhandler
    End If

    Arch = Resource_Path & "armas.dat"
    
    NumWeaponAnims = val(GetVar(Arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
    
        If loopc <> 2 Then
            InitGrh WeaponAnimData(loopc).WeaponWalk(1), val(GetVar(Arch, "ARMA" & loopc, "Dir1")), 0
            InitGrh WeaponAnimData(loopc).WeaponWalk(2), val(GetVar(Arch, "ARMA" & loopc, "Dir2")), 0
            InitGrh WeaponAnimData(loopc).WeaponWalk(3), val(GetVar(Arch, "ARMA" & loopc, "Dir3")), 0
            InitGrh WeaponAnimData(loopc).WeaponWalk(4), val(GetVar(Arch, "ARMA" & loopc, "Dir4")), 0
        End If
    Next loopc
   
Delete_File Resource_Path & "armas.dat"
CargarAnimArmas = True
Exit Function

errorhandler:
    If Len(Err.Description) Then
        If General_File_Exists(Resource_Path & "armas.dat", vbNormal) Then Delete_File Resource_Path & "armas.dat"
        MsgBox "Error while loading weapon index: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error!"
    End If
    
End Function

Private Function CargarAnimEscudos() As Boolean

    On Error Resume Next
    Dim loopc As Long

    Dim Arch  As String
 

    If Not Extract_File(Scripts, App.Path & "\Recursos", "escudos.dat", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo errorhandler
    End If

    Arch = Resource_Path & "escudos.dat"
    
    NumEscudosAnims = val(GetVar(Arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        If loopc <> 2 Then
            InitGrh ShieldAnimData(loopc).ShieldWalk(1), val(GetVar(Arch, "ESC" & loopc, "Dir1")), 0
            InitGrh ShieldAnimData(loopc).ShieldWalk(2), val(GetVar(Arch, "ESC" & loopc, "Dir2")), 0
            InitGrh ShieldAnimData(loopc).ShieldWalk(3), val(GetVar(Arch, "ESC" & loopc, "Dir3")), 0
            InitGrh ShieldAnimData(loopc).ShieldWalk(4), val(GetVar(Arch, "ESC" & loopc, "Dir4")), 0
        End If
    Next loopc
    
Delete_File Resource_Path & "escudos.dat"
CargarAnimEscudos = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        If General_File_Exists(Resource_Path & "escudos.dat", vbNormal) Then Delete_File Resource_Path & "escudos.dat"
        MsgBox "Error while loading shield index: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error!"
    End If
    
End Function


Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef TX As Byte, _
                  ByRef TY As Byte)
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
   
    If viewPortX < 0 Or viewPortX > frmMain.MainViewPic.ScaleWidth Then Exit Sub
    If viewPortY < 0 Or viewPortY > frmMain.MainViewPic.ScaleHeight Then Exit Sub

    TX = UserPos.X + viewPortX \ 32 - frmMain.MainViewPic.ScaleWidth \ 64
    TY = UserPos.Y + viewPortY \ 32 - frmMain.MainViewPic.ScaleHeight \ 64

End Sub

Public Sub UpdateTagAndNameChar(ByVal charindex As Integer, ByVal Name As String)
    Dim Pos   As Integer
    Dim guild As String

    With charlist(charindex)
    
        Pos = getTagPosition(Name)
        
        .Nombre = (Left$(Name, Pos - 2))
        .OffSetNombre = (Text_GetWidth(.Nombre) \ 2) - cfonts.RowPitch
        
        'Clan
        guild = (mid$(Name, Pos))
               
        If (Len(guild) <> 0) Then
            .Clan = guild
            .OffSetClan = Text_Width(.Clan, 1) / 2
            
        ElseIf .priv = 7 Or .priv = 9 Or .priv = 8 Then
          
            .OffSetClan = Text_Width(.Clan, 1) / 2
            
        Else
            .Clan = vbNullString
            .OffSetClan = 0

        End If

    End With

    Exit Sub

End Sub

Sub MakeChar(ByVal charindex As Integer, _
             ByVal body As Integer, _
             ByVal Head As Integer, _
             ByVal Heading As Byte, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Arma As Integer, _
             ByVal Escudo As Integer, _
             ByVal Casco As Integer, ByVal ParticulaFx As Byte)

    On Error Resume Next

    'Apuntamos al ultimo Char

    If charindex > LastChar Then LastChar = charindex
    
    With charlist(charindex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count

        If .active = 0 Then NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = body
        .Head = HeadData(Head)
        .body = BodyData(body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1

        If .Particula = ParticulaFx Then
            ParticulaFx = 0

        End If
        
        If ParticulaFx <> 0 Then
            .Particula = ParticulaFx
            
            Call SetCharacterParticle(ParticulaFx, charindex, -1)
        End If
        
        .Muerto = (Head = CASPER_HEAD)
        
        Call ColorNombresPriv(charindex, .priv)
     End With

    'Plot on map
    MapData(X, Y).charindex = charindex
    
 

End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
        Grh.speed = 0.8
    Else
        Grh.Started = 0
    End If
    

    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    

End Sub

Sub MoveCharbyHead(ByVal charindex As Integer, ByVal nHeading As E_Heading)

    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************
    Dim addX As Integer
    Dim addY As Integer

    Dim X    As Integer
    Dim Y    As Integer

    Dim nX   As Integer
    Dim nY   As Integer
    
    With charlist(charindex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move

        Select Case nHeading

            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.south
                addY = 1
            
            Case E_Heading.WEST
                addX = -1

        End Select
        
        nX = X + addX
        nY = Y + addY
        
        If Not InMapBounds(nX, nY) Then Exit Sub
        
        MapData(nX, nY).charindex = charindex
        .Pos.X = nX
        .Pos.Y = nY
        
        MapData(X, Y).charindex = 0
        
        .MoveOffsetX = -1 * (32 * addX)
        .MoveOffsetY = -1 * (32 * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        If Not .Moving Then
        
            If .Muerto Then
                .Head = HeadData(CASPER_HEAD)
            End If
        
            'Start animations
            If .body.Walk(.Heading).Started = 0 Then
                .body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1

                .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS
            End If
            
            .Moving = True
        End If
        
    End With
    
    If CurrentUser.Muerto = True Then Call DoPasosFx(charindex)
    'areas viejos

    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then

        If charindex <> CurrentUser.UserCharIndex Then
            Call EraseChar(charindex)
        End If

    End If

End Sub

Public Sub DoFogataFx()

    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)

        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0

        End If

    Else
        bFogata = HayFogata(location)

    '    If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, _
                location.Y, LoopStyle.Enabled)

    End If

End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean

    With charlist(charindex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal charindex As Integer)
    If Not UserNavegando Then
        With charlist(charindex)
            If Not .Muerto And EstaPCarea(charindex) Then
                .pie = Not .pie
        'Si esta en una superficie de pasto?
                If MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 6000 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 6559 Then
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS3)
                    Else
                        Call Audio.PlayWave(SND_PASOS4)
                    End If
            'Si esta en una superficie de Arena?
            ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 7700 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 7719 Then
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS5)
                Else
                    Call Audio.PlayWave(SND_PASOS6)
                End If
            'Si esta en una superficie de Nieve?
            ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 7379 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 7507 Then
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS7)
                Else
                    Call Audio.PlayWave(SND_PASOS8)
                End If
            Else
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1)
                Else
                    Call Audio.PlayWave(SND_PASOS2)
                End If
            End If

    'Feo este Sistema****************************
    If UserNavegando Then
    'TODO : Actually we would have to check if the CharIndex char is in the water or not....
        'Call Audio.PlayWave(SND_NAVEGANDO)
    End If
    '********************************************
 
    End If
    End With
    End If
End Sub
Private Function Map_GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As Byte
  If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Then
  Map_GetTerrenoDePaso = 1
  Exit Function
  ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
  Map_GetTerrenoDePaso = 2
  Exit Function
  ElseIf (TerrainFileNum >= 5000 And TerrainFileNum <= 5004) Then
  Map_GetTerrenoDePaso = 3
  Exit Function
  ElseIf TerrainFileNum = 6021 Then
  Map_GetTerrenoDePaso = 4
  Exit Function
  Else
  Map_GetTerrenoDePaso = 5
  End If
End Function

Sub MoveCharbyPos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim X        As Integer
    Dim Y        As Integer

    Dim addX     As Integer
    Dim addY     As Integer

    Dim nHeading As E_Heading
    
    With charlist(charindex)
        X = .Pos.X
        Y = .Pos.Y
        
        If Not InMapBounds(X, Y) Then Exit Sub
        
        MapData(X, Y).charindex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addY) = 1 Then
            nHeading = E_Heading.south
        End If
        
        MapData(nX, nY).charindex = charindex
        
        If nHeading = 0 Then Exit Sub
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (32 * addX)
        .MoveOffsetY = -1 * (32 * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        .LastStep = FrameTime
        
        If Not .Moving Then
        
            If .Muerto Then
                .Head = HeadData(CASPER_HEAD)
            End If
        
            'Start animations
            If .body.Walk(.Heading).Started = 0 Then
                .body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1

                .Arma.WeaponWalk(.Heading).Loops = INFINITE_LOOPS
                .Escudo.ShieldWalk(.Heading).Loops = INFINITE_LOOPS
            End If
            
            .Moving = True
        End If
        
    End With
 
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)
    End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)

    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    Dim X  As Integer
    Dim Y  As Integer
    Dim TX As Integer
    Dim TY As Integer
    
    'Figure out which way to move

    Select Case nHeading

        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.south
            Y = 1
        
        Case E_Heading.WEST
            X = -1

    End Select
    
    'Fill temp pos
    TX = UserPos.X + X
    TY = UserPos.Y + Y
    
    'Check to see if its out of bounds

    'Check to see if its out of bounds
    If TX < XMinMapSize Or TX > XMaxMapSize Or TY < YMinMapSize Or TY > YMaxMapSize Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        UserPos.X = TX
        UserPos.Y = TY
        
        AddtoUserPos.X = Sgn(-X)
        AddtoUserPos.Y = Sgn(-Y)
 
        OffSetCounterX = -1 * (32 * -X)
        OffSetCounterY = -1 * (32 * -Y)
         
        UserMoving = 1
        
        bTecho = Char_Techo

    End If

End Sub

Private Function HayFogata(ByRef location As Position) As Boolean

    Dim J As Long

    Dim k As Long
    
    For J = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6

            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = J
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function

                End If

            End If

        Next k
    Next J

End Function
 
Private Function LoadGrhData() As Boolean
    '*****************************************************************
    'Loads Grh.dat
    '*****************************************************************

    On Error GoTo errorhandler
     
    Dim Grh     As Integer
    Dim Frame   As Integer
    Dim tempInt As Integer
    ReDim GrhData(0 To 40000) As GrhData
 
     If Not Extract_File(Scripts, App.Path & "\Recursos", "graficos.ind", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo errorhandler
    End If
    

    If Not Extract_File(Scripts, App.Path & "\Recursos", "minimap.dat", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo errorhandler
    End If


    
    Open Resource_Path & "graficos.ind" For Binary Access Read As #1
       
    Seek #1, 1
       
    Get #1, , tempInt
    Get #1, , tempInt
    Get #1, , tempInt
    Get #1, , tempInt
    Get #1, , tempInt
     
    'Get first Grh Number
    Get #1, , Grh
       
    Do Until Grh <= 0
        'Get number of frames
        Get #1, , GrhData(Grh).NumFrames
           
        If GrhData(Grh).NumFrames <= 0 Then
            GoTo errorhandler

        End If
           
        ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
           
        If GrhData(Grh).NumFrames > 1 Then
           
            'Read a animation GRH set

            For Frame = 1 To GrhData(Grh).NumFrames
                Get #1, , GrhData(Grh).Frames(Frame)

                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > 40000 Then GoTo errorhandler
                
            Next Frame
           
            Get #1, , tempInt
               
            If tempInt <= 0 Then GoTo errorhandler
            GrhData(Grh).speed = CSng(tempInt)
               
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight

            If GrhData(Grh).pixelHeight <= 0 Then GoTo errorhandler
               
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth

            If GrhData(Grh).pixelWidth <= 0 Then GoTo errorhandler
     
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth

            If GrhData(Grh).TileWidth <= 0 Then GoTo errorhandler
     
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight

            If GrhData(Grh).TileHeight <= 0 Then GoTo errorhandler
        Else
            'Read in normal GRH data
            Get #1, , GrhData(Grh).FileNum

            If GrhData(Grh).FileNum <= 0 Then GoTo errorhandler
     
            Get #1, , GrhData(Grh).sX

            If GrhData(Grh).sX < 0 Then GoTo errorhandler
               
            Get #1, , GrhData(Grh).sY

            If GrhData(Grh).sY < 0 Then GoTo errorhandler
     
            Get #1, , GrhData(Grh).pixelWidth

            If GrhData(Grh).pixelWidth <= 0 Then GoTo errorhandler
     
            Get #1, , GrhData(Grh).pixelHeight

            If GrhData(Grh).pixelHeight <= 0 Then GoTo errorhandler
     
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
               
            GrhData(Grh).Frames(1) = Grh

        End If

        'Get Next Grh Number
        Get #1, , Grh
    Loop
       
Close #1

Dim Count As Long
'Cargamos MiniMap
Open Resource_Path & "minimap.dat" For Binary As #1

    Seek #1, 1
    For Count = 1 To 40000
        If Grh_Check(Count) Then
            Get #1, , GrhData(Count).mini_map_color
        End If
    Next Count
    Close #1
    
    Delete_File Resource_Path & "Graficos.ind"
    Delete_File Resource_Path & "minimap.dat"


    LoadGrhData = True

Exit Function

errorhandler:
    If Len(Err.Description) Then
        Close #1
        If General_File_Exists(Resource_Path & "graficos.ind", vbNormal) Then Delete_File Resource_Path & "graficos.ind"
        If General_File_Exists(Resource_Path & "minimap.dat", vbNormal) Then Delete_File Resource_Path & "minimap.dat"
        MsgBox "Error while loading graphic index: " & Err.Description & " (" & Grh & ")", vbCritical, "Error!"
    Else
        Close #1
        If General_File_Exists(Resource_Path & "graficos.ind", vbNormal) Then Delete_File Resource_Path & "graficos.ind"
        If General_File_Exists(Resource_Path & "minimap.dat", vbNormal) Then Delete_File Resource_Path & "minimap.dat"
        MsgBox "Error while loading graphic index" & " (" & Grh & ")", vbCritical, "Error!"
    End If
    
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 01/08/2009
    'Checks to see if a tile position is legal, including if there is a casper in the tile
    '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
    '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
    '*****************************************************************
    Dim charindex As Integer
    
    'Limites del mapa
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    'Tile Bloqueado?

    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    charindex = MapData(X, Y).charindex

    '¿Hay un personaje?

    If charindex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(charindex)

            ' Si no es casper, no puede pasar

            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else

                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)

                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else

                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)

                    If HayAgua(X, Y) Then Exit Function

                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles

                If charlist(CurrentUser.UserCharIndex).priv > 0 And charlist(CurrentUser.UserCharIndex).priv < 6 Then

                    If charlist(CurrentUser.UserCharIndex).Invisible = True Then Exit Function

                End If

            End If

        End With

    End If
   
     If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    If CurrentUser.Montando = True Then
        If MapData(X, Y).Trigger = 1 Or MapData(X, Y).Trigger = 2 Or MapData(X, Y).Trigger = 4 Or MapData(X, Y).Trigger >= 20 Then
            Exit Function
        End If
    End If
    
    MoveToLegalPos = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************

    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function

    End If
    
    InMapBounds = True

End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Function LegalPos(ByVal X As Integer, ByVal Y As Integer, ByVal Heading As E_Heading) As Boolean
    
    On Error GoTo LegalPos_Err
    

    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************
    
    'Limites del mapa
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).charindex > 0 Then
        With charlist(MapData(X, Y).charindex)
            If Not (.Muerto Or (.Invisible And .priv > charlist(CurrentUser.UserCharIndex).priv)) Then
                Exit Function
            End If
        End With
    End If
    
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If

     If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    'If MapData(X, Y).OBJInfo.tipe = eObjType.otTeleport Then
 

    'End If
        
    
    LegalPos = True
    
    Exit Function

LegalPos_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.LegalPos", Erl)
    Resume Next
    
End Function

Public Sub DrawGrhtoSurface(ByRef Grh As Grh, _
                            ByVal X As Integer, _
                            ByVal Y As Integer, _
                            ByVal Center As Byte, _
                            ByVal Animate As Byte, _
                            ByRef color() As Long, _
                            Optional ByVal killAtEnd As Byte = 1, _
                            Optional ByVal angle As Single = 0, _
                            Optional ByVal AlphaB As Byte = 0)
                            
    '*****************************************************************
    'Draws a GRH transparently to a X and Y position
    '*****************************************************************
    Dim CurrentGrhIndex As Integer

    On Error GoTo hError
   
    If Grh.GrhIndex = 0 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1 + Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames
               
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
 
                    Else
                        Grh.Started = 0
                    End If

                End If

            End If

        End If

    End If
   
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
   
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos

        If Center Then
            If .TileWidth <> 1 Then
               X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

            End If
           
            If .TileHeight <> 1 Then
                Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

            End If

        End If

        Call Directx_Render_Texture(CLng(.FileNum), X, Y, .pixelHeight, .pixelWidth, .sX, .sY, color(), angle, AlphaB)

    End With

    Exit Sub
    
hError:

    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & _
                "Descripción del error: " & vbCrLf & Err.Description, "Graphics"
        'End

    End If

End Sub

Sub DrawGrhIndextoSurface(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Colour As Long = -1)
        
        If grh_index <= 0 Then Exit Sub
        
        Dim rgb_list(3) As Long
        
        rgb_list(0) = Colour
        rgb_list(1) = Colour
        rgb_list(2) = Colour
        rgb_list(3) = Colour
        
        'Draw
        Call Directx_Render_Texture(CLng(GrhData(grh_index).FileNum), X, Y, GrhData(grh_index).pixelHeight, GrhData(grh_index).pixelWidth, GrhData(grh_index).sX, GrhData(grh_index).sY, rgb_list())
 

End Sub

Sub DrawGrhtoHdc(desthDC As Long, ByVal grh_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean)
On Error Resume Next

Dim file_path As String
Dim src_x As Integer
Dim src_y As Integer
Dim src_width As Integer
Dim src_height As Integer
Dim hdcsrc As Long
Dim MaskDC As Long
Dim PrevObj As Long
Dim PrevObj2 As Long
Dim bRet As Boolean

If grh_index <= 0 Then Exit Sub
   'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(grh_index).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(grh_index).TileWidth * 16) + 16
        End If
    End If
    
    If v_centered Then
        If GrhData(grh_index).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32
        End If
    End If

'If it's animated switch grh_index to first frame
If GrhData(grh_index).NumFrames <> 1 Then
grh_index = GrhData(grh_index).Frames(1)
End If

'file_path = App.Path & "\RECURSOS\GRAFICOS\" & GrhData(grh_index).FileNum & ".bmp"
bRet = Extract_File(Graphics, App.Path & "\Recursos", GrhData(grh_index).FileNum & ".bmp", Resource_Path, False)


If bRet Then
file_path = Resource_Path & GrhData(grh_index).FileNum & ".bmp"
src_x = GrhData(grh_index).sX
src_y = GrhData(grh_index).sY
src_width = GrhData(grh_index).pixelWidth
src_height = GrhData(grh_index).pixelHeight
hdcsrc = CreateCompatibleDC(desthDC)

PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))

 

 If transparent = False Then
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    Else
        TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, &HFF000000
 End If
  
    DeleteDC hdcsrc
'Delete_File file_path
Delete_File (Resource_Path & GrhData(grh_index).FileNum & ".bmp")
End If
End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, _
                                   ByVal srchdc As Long, _
                                   ByRef SourceRect As Rect, _
                                   ByRef DestRect As Rect, _
                                   ByVal TransparentColor)

    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 12/22/2009
    'This method is SLOW... Don't use in a loop if you care about
    'speed!
    '*************************************************************
    Dim color As Long
    Dim X     As Long
    Dim Y     As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.bottom
            color = GetPixel(srchdc, X, Y)
            
            If color <> TransparentColor Then
                Call SetPixel(dsthdc, DestRect.Left + (X - SourceRect.Left), DestRect.Top + (Y - SourceRect.Top), color)

            End If

        Next Y
    Next X

End Sub
Sub RenderScreen()
    
    On Error GoTo Errormanejador
    
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '********
    Dim Y                As Long     'Keeps track of where on map we are
    Dim X                As Long     'Keeps track of where on map we are
    Dim Techos(0 To 3)   As Long     'Alpha techos
    Dim screenminY       As Integer  'Start Y pos on current screen
    Dim screenmaxY       As Integer  'End Y pos on current screen
    Dim screenminX       As Integer  'Start X pos on current screen
    Dim screenmaxX       As Integer  'End X pos on current screen
    Dim minY             As Integer  'Start Y pos on current map
    Dim maxY             As Integer  'End Y pos on current map
    Dim minX             As Integer  'Start X pos on current map
    Dim maxX             As Integer  'End X pos on current map
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim color As D3DCOLORVALUE
    Dim TempColor(3) As Long
    Dim tilex               As Integer
    Dim tiley               As Integer
    Dim PixelOffSetX        As Integer
    Dim PixelOffSetY        As Integer
    Dim addX As Integer, addY As Integer
        
    If Rendimiento = 1 Then Call ConvertCPtoTP(frmMain.MouseX, frmMain.MouseY, TX, TY)
    
    If UserMoving Then
        If AddtoUserPos.X <> 0 Then
            OffSetCounterX = OffSetCounterX + scroll_pixels_per_frame * Sgn(AddtoUserPos.X) * timerTicksPerFrame
            If (Sgn(AddtoUserPos.X) = 1 And OffSetCounterX >= 0) Or (Sgn(AddtoUserPos.X) = -1 And OffSetCounterX <= 0) Then
                OffSetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = False
            End If
        End If
        
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.Y <> 0 Then
            OffSetCounterY = OffSetCounterY + scroll_pixels_per_frame * Sgn(AddtoUserPos.Y) * timerTicksPerFrame
            If (Sgn(AddtoUserPos.Y) = 1 And OffSetCounterY >= 0) Or (Sgn(AddtoUserPos.Y) = -1 And OffSetCounterY <= 0) Then
                OffSetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = False
            End If
        End If
    End If
 

    tilex = UserPos.X
    tiley = UserPos.Y
    
    PixelOffSetX = OffSetCounterX '_offset_counter_x
    PixelOffSetY = OffSetCounterY '_offset_counter_y
    
    
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - (TileBufferSize + 4)
    maxY = screenmaxY + (TileBufferSize + 4)
    minX = screenminX - (TileBufferSize + 4)
    maxX = screenmaxX + (TileBufferSize + 4)
    
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
        addY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
        addX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    RoofAlphaCalculate Techos
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
 
            If InMapBounds(X, Y) Then
            
                'Layer 1 ****
                If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(1), (ScreenX - 1) * 32 + PixelOffSetX, (ScreenY - 1) * 32 + PixelOffSetY, 0, 1, MapData(X, Y).light_value, , X, Y, , 5)
                End If
            
            End If
            
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

    ScreenY = 0
    ScreenX = 0
     
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
                If InMapBounds(X, Y) Then
                    'Layer 2 ****
                    If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1 + addX) * 32 + PixelOffSetX, (ScreenY - 1 + addY) * 32 + PixelOffSetY, 1, 1, MapData(X, Y).light_value, , X, Y)
                    End If
                    '******
            End If
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

    'Draw Transparent Layers
    ScreenY = minYOffset - (TileBufferSize + 4)
    For Y = minY To maxY
        ScreenX = minXOffset - (TileBufferSize + 4)
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffSetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffSetY
            With MapData(X, Y)
            
                If InMapBounds(X, Y) Then
 
                
                    'Object Layer ****
                    If .ObjGrh.GrhIndex Then
                        If .ObjGrh.GrhIndex <> 0 Then
                            Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)
                        End If
                    End If
                    '*******
                    
                    'Layer 3 *****
                    If .Graphic(3).GrhIndex <> 0 And Not .Graphic(3).GrhIndex = .ObjGrh.GrhIndex Then
                        Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)
                    End If
                    '*******
                 
                        If Rendimiento = 1 Then
                        
                           If .CharFantasma.Activo Then
                               If .CharFantasma.AlphaB > 0 Then
                                   color = ambientLight
                                   color.a = .CharFantasma.AlphaB
                                   D3DColorToRgbList TempColor, color
                                   .CharFantasma.AlphaB = .CharFantasma.AlphaB - (timerTicksPerFrame * 6)
                                   'Redondeamos a 0 para prevenir errores
                                   If .CharFantasma.AlphaB < 0 Then .CharFantasma.AlphaB = 0
                                
                                    'Auras
                                   If (.CharFantasma.Body_Aura) <> 0 Then Call Renderizar_Aura(.CharFantasma.Body_Aura, PixelOffsetXTemp, PixelOffsetYTemp)
                                   If (.CharFantasma.Arma_Aura) <> 0 Then Call Renderizar_Aura(.CharFantasma.Arma_Aura, PixelOffsetXTemp, PixelOffsetYTemp)
                                   If (.CharFantasma.Otra_Aura) <> 0 Then Call Renderizar_Aura(.CharFantasma.Otra_Aura, PixelOffsetXTemp, PixelOffsetYTemp)
                                   If (.CharFantasma.Escudo_Aura) <> 0 Then Call Renderizar_Aura(.CharFantasma.Escudo_Aura, PixelOffsetXTemp, PixelOffsetYTemp)
                                   If (.CharFantasma.Anillo_Aura) <> 0 Then Call Renderizar_Aura(.CharFantasma.Anillo_Aura, PixelOffsetXTemp, PixelOffsetYTemp)
 
        
                                       If .CharFantasma.donador = 1 Then Call DrawGrhIndextoSurface(Estrella, PixelOffsetXTemp - CInt(Engine_Text_Width(.CharFantasma.Nombre, 1) / 2), PixelOffsetYTemp + 28, D3DColorXRGB(255, 255, 255))
                                       If .CharFantasma.EsUsuario Then
                                       
                                           If Len(.CharFantasma.Nombre) > 0 Then Engine_Text_Render .CharFantasma.Nombre, PixelOffsetXTemp + 15 + NickModernoXX - CInt(Engine_Text_Width(.CharFantasma.Nombre, True) / 2), PixelOffsetYTemp + 30 - Engine_Text_Height(.CharFantasma.Nombre, True), .CharFantasma.color, NombresModernos
                                           If .CharFantasma.OffSetClan > 0 Then Engine_Text_Render .CharFantasma.Clan, PixelOffsetXTemp - 13 + NickModernoX - CInt(Engine_Text_Width(.CharFantasma.Nombre, True) / 2), PixelOffsetYTemp + 46 - Engine_Text_Height(.CharFantasma.Nombre, True), .CharFantasma.color, NombresModernos
                                       
                                       End If
    
                                       'Seteamos el color
                                       If .CharFantasma.Heading = 1 Or .CharFantasma.Heading = 2 Then
                                       Call DrawGrhtoSurface(.CharFantasma.Escudo, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TempColor, 1, 0, 0)
                                       
                                       Call DrawGrhtoSurface(.CharFantasma.body, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Head, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 0, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Casco, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 0, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Arma, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TempColor, 1, 0, 0)
                                       Else
                                       Call DrawGrhtoSurface(.CharFantasma.body, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Head, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 0, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Escudo, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Casco, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 0, TempColor, 1, 0, 0)
                                       Call DrawGrhtoSurface(.CharFantasma.Arma, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, TempColor, 1, 0, 0)
                                       End If
            
                               Else
                                   .CharFantasma.Activo = False
                               End If
                           End If
                        End If
                        

                   If .charindex <> 0 Then
                       If charlist(.charindex).Pos.X <> X Or charlist(.charindex).Pos.Y <> Y Then
                           Call Char_Refresh(.charindex)
                           .charindex = 0
                       Else
                           Call CharRender(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
                       End If
                   End If
                
                '******
                If .FxIndex > 0 Then
                    Call DrawGrhtoSurface(.fX, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, .light_value, 1, 0, 1)
                If .fX.Started = 0 Then .FxIndex = 0
                End If
                '******
                
            End If
  
            End With

            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5

    ScreenY = minYOffset - (TileBufferSize + 4)
    For Y = minY To maxY
        ScreenX = minXOffset - (TileBufferSize + 4)
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffSetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffSetY
            
            'Particles*****************************************
            
            If InMapBounds(X, Y) Then
                If MapData(X, Y).particle_group > 0 Then
                    Particle_Group_Render MapData(X, Y).particle_group, PixelOffsetXTemp, PixelOffsetYTemp
                End If
            End If
            
             ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5

    If AlphaY <> 0 Then
    
    ScreenY = minYOffset - (TileBufferSize + 4)
    
    For Y = minY To maxY
        ScreenX = minXOffset - (TileBufferSize + 4)
        
        For X = minX To maxX
            
            If InMapBounds(X, Y) Then
                    'Layer 4 ****
                    If MapData(X, Y).Graphic(4).GrhIndex Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), ScreenX * 32 + PixelOffSetX, ScreenY * 32 + PixelOffSetY, 1, 1, Techos, , X, Y)
                    End If
            End If
        
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
    
    End If
    
    
    If map_letter_fadestatus > 0 Then
         
        If map_letter_fadestatus = 1 Then
            map_letter_a = map_letter_a + (timerTicksPerFrame * 3.5)
            If map_letter_a >= 255 Then
                map_letter_a = 255
                map_letter_fadestatus = 2
            End If
        Else
            map_letter_a = map_letter_a - (timerTicksPerFrame * 3.5)
            If map_letter_a <= 0 Then
                map_letter_fadestatus = 0
                map_letter_a = 0
                 
                If map_letter_grh_next > 0 Then
                    map_letter_grh.GrhIndex = map_letter_grh_next
                    map_letter_fadestatus = 1
                    map_letter_grh_next = 0
                End If
                
            End If
        End If

        Techos(0) = D3DColorARGB(CInt(map_letter_a), 255, 255, 255)
        Techos(1) = Techos(0)
        Techos(2) = Techos(0)
        Techos(3) = Techos(0)
        
        
        Grh_Render map_letter_grh, 250, 75, Techos
    
    End If
    
    If meteo_particle Then Particle_Group_Render meteo_particle, 250, 0
    
    If trueno And Not trueno = 1 Then
      '  Trueno_Render
        trueno = trueno - 1
    ElseIf trueno = 1 Then
        Call Audio.PlayWave("105.wav", RandomNumber(1, 100), RandomNumber(1, 100))
        'Trueno_Render
        trueno = 0
    End If
    

    If FPSFLAG = True Then
        Engine_Text_Render fps & " FPS", 494, 2, Default_RGB
        Engine_Text_Render MSRender & " MS", 494, 17, Default_RGB
    End If
    
    If charlist(CurrentUser.UserCharIndex).EsGM = True Then
        If CurrentUser.RenderGM = True Then
            Engine_Text_Render "Panel Game Master visible (Tecla P): " & Cuenta.UserName, 0, 0, Default_RGB
            Engine_Text_Render "Formulario de búsqueda (Tecla B)", 0, 20, Default_RGB
            Engine_Text_Render "Viajar (Tecla M)", 0, 40, Default_RGB
            Engine_Text_Render "Prueba Fonttypes (Tecla F)", 0, 60, Default_RGB
        End If
    End If
 
    Exit Sub

Errormanejador:
    Call RegistrarError(Err.number, Err.Description, "engine.Renderscreen", Erl)
    Resume Next
    
End Sub

Private Function GetElapsedTime() As Single
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time    As Currency
    Static end_time   As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq

    End If

    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Private Sub CharRender(ByVal charindex As Long, ByVal PixelOffSetX As Integer, ByVal PixelOffSetY As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim i                  As Long
    Dim rgb_list(0 To 3)   As Long

 
    With charlist(charindex)
    
        If .Heading = 0 Then Exit Sub

        If .Moving Then
        
            'If needed, move left and right

            If .scrollDirectionX <> 0 Then
            
                .MoveOffsetX = .MoveOffsetX + (IIf(CurrentUser.UserCharIndex = charindex, scroll_pixels_per_frame, scroll_pixels_per_frameBackUp) * Sgn(.scrollDirectionX) * timerTicksPerFrame)
 
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
                
            End If
           
            'If needed, move up and down

            If .scrollDirectionY <> 0 Then
            
                .MoveOffsetY = .MoveOffsetY + (IIf(CurrentUser.UserCharIndex = charindex, scroll_pixels_per_frame, scroll_pixels_per_frameBackUp) * Sgn(.scrollDirectionY) * timerTicksPerFrame)
 
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If

            End If
            
            If .scrollDirectionX = 0 And .scrollDirectionY = 0 Then
                .Moving = False
            End If
    
        Else
        
            .body.Walk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).Started = 0
            .body.Walk(.Heading).FrameCounter = 1
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
        End If
         
        PixelOffSetX = PixelOffSetX + .MoveOffsetX
        PixelOffSetY = PixelOffSetY + .MoveOffsetY
        
 
        If .body.Walk(.Heading).GrhIndex Then
            
            If Not .Invisible And Not .Muerto Then
                .AlphaX = 0
                rgb_list(0) = MapData(X, Y).light_value(0)
                rgb_list(1) = MapData(X, Y).light_value(1)
                rgb_list(2) = MapData(X, Y).light_value(2)
                rgb_list(3) = MapData(X, Y).light_value(3)
            
                If (.Body_Aura) <> 0 Then Call Renderizar_Aura(.Body_Aura, PixelOffSetX, PixelOffSetY)
                If (.Arma_Aura) <> 0 Then Call Renderizar_Aura(.Arma_Aura, PixelOffSetX, PixelOffSetY)
                If (.Otra_Aura) <> 0 Then Call Renderizar_Aura(.Otra_Aura, PixelOffSetX, PixelOffSetY)
                If (.Escudo_Aura) <> 0 Then Call Renderizar_Aura(.Escudo_Aura, PixelOffSetX, PixelOffSetY)
                If (.Anillo_Aura) <> 0 Then Call Renderizar_Aura(.Anillo_Aura, PixelOffSetX, PixelOffSetY)
                                
            Else
                Call RoofAlphaCalculateToAlpha(rgb_list(), charindex)
            End If
            

            If .Head.Head(.Heading).GrhIndex Then
            
                Select Case .Heading
                
                Case E_Heading.EAST
                
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 36, 1, 1, rgb_list(), , X, Y)
                    If .body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 1, rgb_list(), , X, Y)
                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                
                
                Case E_Heading.NORTH
                    
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                    If .body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 1, rgb_list(), , X, Y)
                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                                
                Case E_Heading.WEST
                
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                    If .body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 1, rgb_list(), , X, Y)
                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                                                
                Case E_Heading.south
                    If .body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 1, rgb_list(), , X, Y)
                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y + 37, 1, 1, rgb_list(), , X, Y)
                
                End Select
            
            Else
            
                Call Draw_Grh(.body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 1, rgb_list(), , X, Y)
            
            End If
            
        End If
        

        
        If Nombres Then
        
            If Len(.Nombre) > 0 Then
            
                Dim line As String
                Dim Pos As Integer, OffsetYname As Byte
                
                If Pos = 0 Then Pos = Len(.Nombre) + 2
        
                line = Left$(.Nombre, Pos - 2)
    
                If .EsUsuario Then
                    
                    If Not .Invisible Or .priv > 0 Then
                    
                    
                        If Rendimiento = 1 Then
                            If .EsUsuario Then
                                Engine_Text_Render line, PixelOffSetX + 15 + NickModernoXX - CInt(Engine_Text_Width(line, True) / 2), PixelOffSetY + 30 + OffsetYname - Engine_Text_Height(line, True), .color, NombresModernos
                                
                                If .OffSetClan > 0 Then
                                    Engine_Text_Render .Clan, PixelOffSetX + 15 + NickModernoX - CInt(Engine_Text_Width(.Clan, True) / 2), PixelOffSetY + 45 + OffsetYname - Engine_Text_Height(.Clan, True), .color, NombresModernos
                                End If
                            End If
                        Else
                            If .EsUsuario Then
                                Engine_Text_Render line, PixelOffSetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffSetY + 30 + OffsetYname - Engine_Text_Height(line, True), .color, 1
                            End If
                        End If
        
                        If .donador = 1 Then
                            line = Left$(.Nombre, Pos - 2)
                            Call DrawGrhIndextoSurface(Estrella, PixelOffSetX - CInt(Engine_Text_Width(line, 1) / 2), PixelOffSetY + 28 + OffsetYname, D3DColorXRGB(255, 255, 255))
                        End If
                        
                    End If
                    
                    
                ElseIf Not .NPCHostil And Rendimiento = 1 Then
                
                    If Abs(TX - .Pos.X) < 1 And TY - .Pos.Y < 1 And .Pos.Y - TY < 2 Then
                        If MapData(TX, TY).NPCIndex <> 0 Then
                            Engine_Text_Render line, PixelOffSetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffSetY + 30 + OffsetYname - Engine_Text_Height(line, True), .color, 1
                            
                            If General_Locale_NPCs((MapData(TX, TY).NPCIndex), 2) <> "" Then
                                line = "<" & General_Locale_NPCs((MapData(TX, TY).NPCIndex), 2) & ">"
                                Engine_Text_Render line, PixelOffSetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffSetY + 43 + OffsetYname - Engine_Text_Height(line, True), .color, 1
                            End If
                        End If
                    End If

                End If 'Fin EsUser o Npc
                
            End If 'Fin si tiene nick
        
        End If 'Fin si esta con el nombre visible
 
     
        If charlist(CurrentUser.UserCharIndex).EsGM Then
            If MapData(.Pos.X, .Pos.Y).charindex = UserFichado Then
                Call Engine_Text_Render("[Target]", PixelOffSetX - 10, PixelOffSetY, .color)
            End If
        End If
 
        'Draw FX
        If .FxIndex <> 0 Then
            Call DrawGrhtoSurface(.fX, PixelOffSetX + FxData(.FxIndex).OffSetX, PixelOffSetY + FxData(.FxIndex).OffSetY, 1, 1, MapData(X, Y).light_value, 1, 0, 1)
            'Check if animation is over
            If .fX.Started = 0 Then .FxIndex = 0
        End If
 
           
         If charlist(charindex).particle_count > 0 Then
            For i = 1 To UBound(charlist(charindex).particle_group)
                If charlist(charindex).particle_group(i) > 0 Then
                    Call Particle_Group_Render(.particle_group(i), PixelOffSetX, PixelOffSetY)
                End If
            Next i
         End If
        
         If .dl Then
            If FrameTime - .dialogStart >= .dialogLife Then
                Char_Dialog_Remove charindex
                
            Else
                If .dialogHeight > 0 Then .dialogHeight = .dialogHeight + (4 * timerTicksPerFrame * Sgn(-1))
                
                If Sgn(.dialogHeight) = -1 Then .dialogHeight = 0
 
                PixelOffSetY = PixelOffSetY - 13 * UBound(.dialog) + .dialogHeight
                For i = 0 To UBound(.dialog)
                    Engine_Text_Render LTrim(.dialog(i)), PixelOffSetX + .body.HeadOffset.X - Text_Width(.dialog(i), 1) / 2 + 16, PixelOffSetY + .body.HeadOffset.Y, .dialogColor, .dialogIndex
                    PixelOffSetY = PixelOffSetY + 8 + 5
                Next i
            End If
        End If
        
    End With

End Sub
Private Sub Renderizar_Aura(ByVal aura_index As String, ByVal X As Integer, ByVal Y As Integer)
    
    On Error GoTo Renderizar_Aura_Err


    With auras(aura_index)
    
         If .Grh <> 0 Then
             If .Rotation = 1 Then
                 If GetTickCount - .tickCount > fps Then
                     .angle = .angle + (timerTicksPerFrame * 5)
                     If .angle >= 360 Then .angle = 0
                 End If
             End If
            Call DrawGrhIndextoSurfaceAlpha(.Grh, X + .OffSetX, Y + .OffSetY, 1, .color, .angle, 1)
         End If
         
     End With

    Exit Sub

Renderizar_Aura_Err:
    Call RegistrarError(Err.number, Err.Description, "engine.Renderizar_Aura", Erl)
    Resume Next
    
End Sub
Public Sub SetCharacterFx(ByVal charindex As Integer, _
                          ByVal fX As Integer, _
                          ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************

    With charlist(charindex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops

        End If

    End With

End Sub

Private Sub InitColours()
        
    White(0) = D3DColorXRGB(255, 255, 255)
    White(1) = D3DColorXRGB(255, 255, 255)
    White(2) = D3DColorXRGB(255, 255, 255)
    White(3) = D3DColorXRGB(255, 255, 255)
    
    red(0) = D3DColorXRGB(255, 0, 0)
    red(1) = D3DColorXRGB(255, 0, 0)
    red(2) = D3DColorXRGB(255, 0, 0)
    red(3) = D3DColorXRGB(255, 0, 0)
    
    Cyan(0) = D3DColorXRGB(0, 255, 255)
    Cyan(1) = D3DColorXRGB(0, 255, 255)
    Cyan(2) = D3DColorXRGB(0, 255, 255)
    Cyan(3) = D3DColorXRGB(0, 255, 255)
    
    Black(0) = D3DColorARGB(255, 0, 0, 0)
    Black(1) = D3DColorARGB(255, 0, 0, 0)
    Black(2) = D3DColorARGB(255, 0, 0, 0)
    Black(3) = D3DColorARGB(255, 0, 0, 0)
    
    FaintBlack(0) = D3DColorXRGB(50, 50, 50)
    FaintBlack(1) = D3DColorXRGB(50, 50, 50)
    FaintBlack(2) = D3DColorXRGB(50, 50, 50)
    FaintBlack(3) = D3DColorXRGB(50, 50, 50)
    
    Yellow(0) = D3DColorXRGB(255, 255, 0)
    Yellow(1) = D3DColorXRGB(255, 255, 0)
    Yellow(2) = D3DColorXRGB(255, 255, 0)
    Yellow(3) = D3DColorXRGB(255, 255, 0)
    
    Gray(0) = D3DColorXRGB(150, 150, 150)
    Gray(1) = D3DColorXRGB(150, 150, 150)
    Gray(2) = D3DColorXRGB(150, 150, 150)
    Gray(3) = D3DColorXRGB(150, 150, 150)
    
    transparent(0) = D3DColorXRGB(255, 255, 255)
    transparent(1) = D3DColorXRGB(255, 255, 255)
    transparent(2) = D3DColorXRGB(255, 255, 255)
    transparent(3) = D3DColorXRGB(255, 255, 255)
    
    green(0) = D3DColorXRGB(0, 255, 0)
    green(1) = D3DColorXRGB(0, 255, 0)
    green(2) = D3DColorXRGB(0, 255, 0)
    green(3) = D3DColorXRGB(0, 255, 0)
    
    blue(0) = D3DColorXRGB(0, 0, 255)
    blue(1) = D3DColorXRGB(0, 0, 255)
    blue(2) = D3DColorXRGB(0, 0, 255)
    blue(3) = D3DColorXRGB(0, 0, 255)
    
End Sub


Public Sub Engine_DirectX8_Init()

    On Error GoTo ErrHandler:
    
        'Create the DirectX8 object
1     Set DirectX = New DirectX8
4     Set DirectD3D = DirectX.Direct3DCreate
      Set DirectD3D8 = New D3DX8
      
      DirectD3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DirectD3Dcaps
      
20    Select Case DeviceIndex

          Case 0
            If Not Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then GoTo ErrHandler
            
          Case 1
            If Not Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then GoTo ErrHandler
            
          Case 2
            If Not Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then GoTo ErrHandler
            
          Case Else
        
        'Detectamos el modo de renderizado mas compatible con tu PC.
        If Not Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    GoTo ErrHandler
                    End
                End If
            End If
        End If
43    End Select
    

46     Call D3DXMatrixOrthoOffCenterLH(Projection, 0, 800, 600, 0, -1#, 1#)
47     Call D3DXMatrixIdentity(View)
    
49     Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
50     Call DirectDevice.SetTransform(D3DTS_VIEW, View)
    
52     Call Directx_RenderStates
    
54     Set Texture = New clsTextureManager
55     Set SpriteBatch = New clsBatch

59     Call SpriteBatch.Initialise(2000)

61     Engine_DirectX8_Aditional_Init
    
63        Exit Sub

    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.Engine_DirectX8_Init", Erl)
    Call MsgBox(Locale_GUI_Frase(348) & " (" & Err.Description & " - " & Err.number & ")", vbCritical, Locale_GUI_Frase(331))
    Call CloseClient
End Sub

Public Sub InitTileEngine()
        On Error GoTo errorhandler:
    
        Dim setTilePixel As Byte
 
        setTilePixel = 32
 
166     TilePixelWidth = setTilePixel
168     TilePixelHeight = setTilePixel
 
174     TileBufferSize = 4
 
        Default_RGB(0) = -1
        Default_RGB(1) = -1
        Default_RGB(2) = -1
        Default_RGB(3) = -1
        
        'Resize mapdata array
194     ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
        'Set intial user position
196     UserPos.X = 50
198     UserPos.Y = 50
        
        AmbientColor = -1

176     HalfWindowTileHeight = 6
178     HalfWindowTileWidth = 8

        CurrentUser.UserMap = 1
 
 
 
204     If Not LoadGrhData Then Exit Sub
        Call frmCargando.progresoConDelay(45)
206     If Not CargarCuerpos Then Exit Sub
208     If Not CargarCabezas Then Exit Sub
210     If Not CargarCascos Then Exit Sub
212     If Not CargarFxs Then Exit Sub
        Call frmCargando.progresoConDelay(60)
        Call LoadGraphics

226     Call CargarParticulas
        Call frmCargando.progresoConDelay(80)
        If Not CargarAnimArmas Then Exit Sub
        If Not CargarAnimEscudos Then Exit Sub
        Call InitColours
        Call Meteo_Init_Time
        
        Estrella = 31981
 
 
     MinXBorder = XMinMapSize + (frmMain.MainViewPic.ScaleWidth / 64)
     MaxXBorder = XMaxMapSize - (frmMain.MainViewPic.ScaleWidth / 64)
     MinYBorder = YMinMapSize + (frmMain.MainViewPic.ScaleHeight / 64)
     MaxYBorder = YMaxMapSize - (frmMain.MainViewPic.ScaleHeight / 64)
     
    'Set scroll pixels per frame
    scroll_pixels_per_frame = 8.5
    
    
    Exit Sub
    
errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.InitTileEngine", Erl)
    MsgBox "¡No se ha logrado iniciar el engine gráfico! Reinstale los últimos controladores de DirectX desde www.Link-AO.com.ar (registrador de librerias) y actualize sus controladores de video. Si el problema persiste por favor consulte al soporte de Link-AO en sus redes sociales.", vbCritical, "Saliendo"
    Call CloseClient
End Sub
Public Sub LoadGraphics()
     Call Texture.Initialize(DirectD3D8, Resource_Path, 90)
End Sub
Public Sub Engine_ActFPS()
    
    On Error GoTo Engine_ActFPS_Err
    

    If FrameTime - lFrameTimer >= 1000 Then
        fps = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = FrameTime
    End If

    
    Exit Sub

Engine_ActFPS_Err:
    Call RegistrarError(Err.number, Err.Description, "engine.Engine_ActFPS", Erl)
    Resume Next
    
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
    On Error GoTo ErrHandler:
1     fps = 60
2     FramesPerSecCounter = 60

3     engineBaseSpeed = 0.018
4
5

6     With MainViewRect
         .x2 = frmMain.MainViewPic.ScaleWidth
         .y2 = frmMain.MainViewPic.ScaleHeight
      End With
      
      Engine_Init_FontTextures
      
  If Not prgRun Then
        Call Text_Font_Initialize
 
        Engine_Init_FontSettings
        
        If Not CargarAuras Then Exit Sub
    End If
      Exit Sub
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.Engine_DirectX8_Aditional_Init", Erl)
    MsgBox "¡No se ha logrado iniciar el engine gráfico! Reinstale los últimos controladores de DirectX desde www.Link-AO.com.ar (registrador de librerias) y actualize sus controladores de video. Si el problema persiste por favor consulte al soporte de Link-AO en sus redes sociales.", vbCritical, "Saliendo"
    Call CloseClient
End Sub
Private Function Init_DirectDevice(ByVal D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo errorhandler:
        
        
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
13         With D3DWindow
4          .Windowed = True

15          .SwapEffect = D3DSWAPEFFECT_DISCARD
18         .BackBufferFormat = DispMode.Format 'current display depth
16          .BackBufferWidth = 800
17         .BackBufferHeight = 600
         .hDeviceWindow = frmMain.MainViewPic.hwnd
19       End With
        
    If Not DirectDevice Is Nothing Then Set DirectDevice = Nothing
    
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, D3DCREATEFLAGS, D3DWindow)
 
    'Lo pongo xq es bueno saberlo...
    Select Case D3DCREATEFLAGS
    
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: SOFTWARE"
            
    End Select
    
    
    Init_DirectDevice = True
    
    
    Exit Function
    
errorhandler:
    
    Set DirectDevice = Nothing
    
    Init_DirectDevice = False

End Function
Public Sub Directx_DeInitialize()

    'Set no Textures to standard stage to avoid memory leak
    If Not DirectDevice Is Nothing Then
        DirectDevice.SetTexture 0, Nothing

    End If

    Set DirectX = Nothing
    Set DirectD3D = Nothing
    Set DirectD3D8 = Nothing
    Set DirectDevice = Nothing
    
    Set SpriteBatch = Nothing
    Set Texture = Nothing
    
    'Clear arrays
    Erase auras
    Erase GrhData
    Erase BodyData
    Erase HeadData
    Erase FxData
    Erase WeaponAnimData
    Erase ShieldAnimData
    Erase CascoAnimData
    Erase MapData
    Erase charlist

    Exit Sub

End Sub

Private Sub Directx_RenderStates()
        
    'Set the render states
    With DirectDevice
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, False
    
    End With
    
End Sub
        
Public Sub Directx_EndScene(ByRef Rect As D3DRECT, ByVal hwnd As Long)
    On Error GoTo Directx_EndScene_Err
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
    Call DirectDevice.Present(Rect, ByVal 0, hwnd, ByVal 0)
    Exit Sub

Directx_EndScene_Err:
 
    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
    
        Call Engine_DirectX8_Init
        
        Call LoadGraphics
    End If
End Sub

Public Function Text_GetWidth(ByVal Text As String) As Integer
    '***************************************************
    'Returns the width of text
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
    '***************************************************
    Dim i As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Text_GetWidth = Text_GetWidth + cfonts.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function
Public Sub Directx_Render_Texture(ByVal FileIndex As Long, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  ByVal Height As Integer, _
                                  ByVal Width As Integer, _
                                  ByVal sX As Integer, _
                                  ByVal sY As Integer, _
                                  ByRef color() As Long, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal AlphaB As Byte = 0)

        On Error GoTo Directx_Render_Texture_Err
        
        Dim TexSurface As Direct3DTexture8
        Dim texwidth   As Integer, texheight As Integer
        Static light_value(0 To 3) As Long
        
    
    
        If CurrentUser.Muerto = True Then
        
            Engine_Long_To_RGB_List light_value, D3DColorXRGB(67, 67, 67)
            
        Else
            light_value(0) = color(0)
            light_value(1) = color(1)
            light_value(2) = color(2)
            light_value(3) = color(3)
               
            If (light_value(0) = 0) Then light_value(0) = AmbientColor 'ambientColor
            If (light_value(1) = 0) Then light_value(1) = AmbientColor 'ambientColor
            If (light_value(2) = 0) Then light_value(2) = AmbientColor ' ambientColor
            If (light_value(3) = 0) Then light_value(3) = AmbientColor 'ambientColor
        End If

100     Set TexSurface = Texture.Surface(FileIndex, texwidth, texheight)
 
102     With SpriteBatch

104         Call .SetAlpha(AlphaB)
    
            '// Seteamos la textura
106         Call .SetTexture(TexSurface)


108         If texwidth <> 0 And texheight <> 0 Then
110             Call .Draw(X, Y, Width, Height, light_value, sX / texwidth, sY / texheight, (sX + Width + 0.1) / texwidth, (sY + Height + 0.1) / texheight, angle)
            Else
112             Call .Draw(X, Y, texwidth, texheight, light_value, , , , , angle)

            End If
  
        End With

        Exit Sub
        
        Exit Sub

Directx_Render_Texture_Err:
        LogError Err.Description & vbCrLf & "in Directx_Render_Texture " & "at line " & Erl
        
End Sub
Public Sub Directx_Render_Texture_Advance(ByVal FileIndex As Long, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  ByVal Height As Integer, _
                                  ByVal Width As Integer, _
                                  ByVal sX As Integer, _
                                  ByVal sY As Integer, _
                                  ByRef color() As Long, _
                                  ByVal dw As Integer, ByVal dH As Integer, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal AlphaB As Byte = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!, _
                                  Optional ByVal Z As Long = 1)
        On Error GoTo Directx_Render_Texture_Err

        
        If FileIndex = 0 Then Exit Sub
        
        Dim TexSurface As Direct3DTexture8
        Dim texwidth   As Integer, texheight As Integer
        Static light_value(0 To 3) As Long

    
        light_value(0) = color(0)
        light_value(1) = color(1)
        light_value(2) = color(2)
        light_value(3) = color(3)
    
        If (light_value(0) = 0) Then light_value(0) = AmbientColor 'ambientColor
        If (light_value(1) = 0) Then light_value(1) = AmbientColor 'ambientColor
        If (light_value(2) = 0) Then light_value(2) = AmbientColor ' ambientColor
        If (light_value(3) = 0) Then light_value(3) = AmbientColor 'ambientColor
        
100     Set TexSurface = Texture.Surface(FileIndex, texwidth, texheight)
 
102     With SpriteBatch

104         Call .SetAlpha(AlphaB)
    
106         Call .SetTexture(TexSurface)

108         If texwidth <> 0 And texheight <> 0 Then
110             Call .Draw(X, Y, dw * ScaleX, dH * ScaleY, color, sX / texwidth, sY / texheight, (sX + Width) / texwidth, (sY + Height) / texheight, angle)
            Else
112             Call .Draw(X, Y, texwidth * ScaleX, texheight * ScaleY, color, , , , , angle)

            End If
  
        End With

        Exit Sub

        '<EhFooter>
        Exit Sub

Directx_Render_Texture_Err:
    Call RegistrarError(Err.number, Err.Description, "Directx_Render_Texture_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Directx_Renderer()

    On Error GoTo error
    
    
    If CurrentUser.UserCharIndex = 0 Then Exit Sub

    If EngineRun Then
 
        Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0)

        SpriteBatch.Begin
 
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> vbMinimized Then
                    
            Call DirectDevice.BeginScene
            Call RenderScreen
            Call Directx_EndScene(MainViewRect, 0)

            'Call AutoUsar
            'Call AutoUsarU
            
            'Play ambient sounds
            'Call DoFogataFx
    
        End If
    
    SpriteBatch.Finish
        
    FrameTime = (timeGetTime() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
    Engine_ActFPS
    
    End If
    
    Exit Sub
    
error:
    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call Engine_DirectX8_Init
        
        Call LoadGraphics
    End If
 
End Sub
Public Sub StartClient()

     
    On Error GoTo Start_Err
    
    FrameTime = timeGetTime And &H7FFFFFFF
    
    Call AsignarHora
    
    Call FXTimer(True)
    Call MinutoTimer(True)
    
    DoEvents
    
    Do While prgRun
        
        Call FlushBuffer
        
      '  If IsAppActive Then
            If frmMain.Visible Then
                Call Meteo_Render
                
                Call Directx_Renderer
                
                Call CheckKeys
                Moviendose = False
                
                'Call RenderSounds
                 Inventory_Render
            End If
      '  Else
      '      If frmMain.Visible Then RenderInv = True
      '      Sleep 60&
      '  End If
 
      DoEvents
        
    Loop
    

    EngineRun = False
    
    Call CloseClient(True)

    Exit Sub

Start_Err:
    Call RegistrarError(Err.number, Err.Description, "engine.Start", Erl)
    Resume Next
    
End Sub
Public Sub Draw_Box(ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal w As Integer, _
                    ByVal h As Integer, _
                    ByRef color As Long)
        Static rgb_list(3) As Long
 
    
        rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color

    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(X, Y, w, h, rgb_list)
     
End Sub

 
Public Sub Char_Clean()

    Dim X As Long
    Dim Y As Long

    For X = 1 To 100
        For Y = 1 To 100

            If MapData(X, Y).charindex Then
                EraseChar MapData(X, Y).charindex

            End If

            If MapData(X, Y).ObjGrh.GrhIndex Then
                Map_Obj_Delete X, Y, True

            End If

        Next Y
    Next X

End Sub

Public Function Light_Remove(ByVal light_index As Long) As Boolean

    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True

    End If
    
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, _
                                      ByRef color_value As Long) As Boolean

    If Light_Check(light_index) Then
        color_value = Light_List(light_index).color
        Light_Color_Value_Get = True

    End If
    
End Function

Public Function Light_Create(ByVal map_x As Integer, _
                             ByVal map_y As Integer, _
                             Optional ByVal color_value As Long = &HFFFFFF, _
                             Optional ByVal range As Byte = 1, _
                             Optional ByVal id As Long) As Long

    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
       ' If Light_Map_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
       '     Exit Function
       ' End If
        Light_Create = Light_Next_Open
        
        Dim r As Integer, g As Integer, b As Integer
        General_Long_Color_to_RGB color_value, r, g, b
        color_value = D3DColorXRGB(r, g, b)
        
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
 
End Function

Public Function Light_Move(ByVal light_index As Long, _
                           ByVal map_x As Integer, _
                           ByVal map_y As Integer) As Boolean

    If Light_Check(light_index) Then
        If InMapBounds(map_x, map_y) Then
        
            Light_Erase light_index
            Light_List(light_index).map_x = map_x
            Light_List(light_index).map_y = map_y
    
            Light_Move = True
            
        End If

    End If
    
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, _
                                   ByVal Heading As Byte) As Boolean

    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX    As Integer
    Dim nY    As Integer
    
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function

    End If

    If Light_Check(light_index) Then
    
        map_x = Light_List(light_index).map_x
        map_y = Light_List(light_index).map_y
        
        nX = map_x
        nY = map_y
        
        Convert_Heading_to_Direction Heading, nX, nY
        
        If InMapBounds(nX, nY) Then
        
            Light_Erase light_index

            Light_List(light_index).map_x = nX
            Light_List(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If

    End If
    
End Function

Private Sub Convert_Heading_to_Direction(ByVal Heading As Byte, _
                                         ByRef direction_x As Integer, _
                                         ByRef direction_y As Integer)

    Dim addY As Integer
    Dim addX As Integer
    
    Select Case Heading

        Case 1
            addY = -1
        
        Case 2
            addY = -1
            addX = 1
        
        Case 3
            addX = 1
        
        Case 4
            addX = 1
            addY = 1
        
        Case 5
            addY = 1
        
        Case 6
            addX = -1
            addY = 1
        
        Case 7
            addX = -1
        
        Case 8
            addX = -1
            addY = -1

    End Select
    
    direction_x = direction_x + addX
    direction_y = direction_y + addY

End Sub

Private Sub Light_Make(ByVal light_index As Long, _
                       ByVal map_x As Integer, _
                       ByVal map_y As Integer, _
                       ByVal rgb_value As Long, _
                       ByVal range As Long, _
                       Optional ByVal id As Long)

    If light_index > Light_last Then
        Light_last = light_index
        ReDim Preserve Light_List(1 To Light_last)

    End If

    Light_Count = Light_Count + 1
    
    Light_List(light_index).active = True
    Light_List(light_index).map_x = map_x
    Light_List(light_index).map_y = map_y
    Light_List(light_index).color = rgb_value
    Light_List(light_index).range = range
    Light_List(light_index).id = id
    
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean

    If light_index > 0 And light_index <= Light_last Then
        If Light_List(light_index).active Then
            Light_Check = True

        End If

    End If
    
End Function

Public Sub Light_Render_All()

    Dim loop_counter As Long
            
    For loop_counter = 1 To Light_Count
        
        If Light_List(loop_counter).active Then
            Light_Render loop_counter

        End If
    
    Next loop_counter
    
End Sub
Private Sub Light_Render(ByVal light_index As Long)
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X     As Long
    Dim Y     As Long
    Dim color As Long
   
    'Set up light borders
    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
    
    'Set color
    color = Light_List(light_index).color
 
    MapData(Light_List(light_index).map_x, Light_List(light_index).map_y).light_value(0) = color
    MapData(Light_List(light_index).map_x, Light_List(light_index).map_y).light_value(1) = color
    MapData(Light_List(light_index).map_x, Light_List(light_index).map_y).light_value(2) = color
    MapData(Light_List(light_index).map_x, Light_List(light_index).map_y).light_value(3) = color
               
    'Arrange corners
   'NE
   If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(1) = color
   End If
'    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(0) = color
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = color
    End If
    
   
    'Arrange borders
    'Upper border
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            
                MapData(X, min_y).light_value(1) = color
                MapData(X, min_y).light_value(2) = color

        End If
    Next X
    'Lower
    For X = min_x + 1 To max_x - 1
    If InMapBounds(X, max_y) Then
    
            MapData(X, max_y).light_value(0) = color
            MapData(X, max_y).light_value(3) = color
        
    End If
    Next X
   
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
                MapData(min_x, Y).light_value(2) = color
                MapData(min_x, Y).light_value(3) = color
        End If
    Next Y
   
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
                MapData(max_x, Y).light_value(0) = color
                MapData(max_x, Y).light_value(1) = color
        End If
        
    Next Y
   
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                    MapData(X, Y).light_value(0) = color
                    MapData(X, Y).light_value(1) = color
                    MapData(X, Y).light_value(2) = color
                    MapData(X, Y).light_value(3) = color
            End If
        Next Y
    Next X
End Sub

Private Function Light_Next_Open() As Long

    On Error GoTo errorhandler:

    Dim loopc As Long
    
    loopc = 1

    Do Until Light_List(loopc).active = False

        If loopc = Light_last Then
            Light_Next_Open = Light_last + 1
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
    Exit Function
errorhandler:
    Light_Next_Open = 1
    
End Function

Public Function Light_Find(ByVal id As Long) As Long

    On Error GoTo errorhandler:

    Dim loopc As Long
    
    loopc = 1

    Do Until Light_List(loopc).id = id

        If loopc = Light_last Then
            Light_Find = 0
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
    Exit Function
errorhandler:
    Light_Find = 0
    
End Function

Public Function Light_Remove_All() As Boolean

    Dim Index As Long
    
    For Index = 1 To Light_last

        If Light_Check(Index) Then
            Light_Destroy Index

        End If

    Next Index
    
    Light_Remove_All = True
    
End Function

Private Sub Light_Destroy(ByVal light_index As Long)

    Dim temp As Light
    
    Light_Erase light_index
    
    Light_List(light_index) = temp
    
    If light_index = Light_last Then

        Do Until Light_List(Light_last).active
            Light_last = Light_last - 1

            If Light_last = 0 Then
                Light_Count = 0
                Exit Sub

            End If

        Loop
        ReDim Preserve Light_List(1 To Light_last)

    End If

    Light_Count = Light_Count - 1
    
End Sub

Private Sub Light_Erase(ByVal light_index As Long)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    
    Dim X     As Long
    Dim Y     As Long


    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
    
    If InMapBounds(min_x, min_y) Then
    
        MapData(min_x, min_y).light_value(2) = AmbientColor

    End If
    
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(1) = AmbientColor

    End If

    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(0) = AmbientColor

    End If

    If InMapBounds(min_x, max_y) Then
       MapData(min_x, max_y).light_value(3) = AmbientColor

    End If
    
    For X = min_x + 1 To max_x - 1

        If InMapBounds(X, min_y) Then
            
                MapData(X, min_y).light_value(1) = AmbientColor
                MapData(X, min_y).light_value(2) = AmbientColor
        End If
        
    Next X
    
    For X = min_x + 1 To max_x - 1

        If InMapBounds(X, max_y) Then
            
                MapData(X, max_y).light_value(0) = AmbientColor
                MapData(X, max_y).light_value(3) = AmbientColor
        End If

    Next X
    
    For Y = min_y + 1 To max_y - 1

        If InMapBounds(min_x, Y) Then
 
                MapData(min_x, Y).light_value(2) = AmbientColor
                MapData(min_x, Y).light_value(3) = AmbientColor
        End If

    Next Y

    For Y = min_y + 1 To max_y - 1

        If InMapBounds(max_x, Y) Then
  
                MapData(max_x, Y).light_value(0) = AmbientColor
                MapData(max_x, Y).light_value(1) = AmbientColor
            
        End If

    Next Y
    
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1

            If InMapBounds(X, Y) Then
                    MapData(X, Y).light_value(0) = AmbientColor
                    MapData(X, Y).light_value(1) = AmbientColor
                    MapData(X, Y).light_value(2) = AmbientColor
                    MapData(X, Y).light_value(3) = AmbientColor
            End If

        Next Y
    Next X
    
End Sub

Public Function Map_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long

    On Error GoTo errorhandler:

    Dim loopc As Long
    
    loopc = Light_last

    Do Until Light_List(loopc).map_x = map_x And Light_List(loopc).map_y = map_y

        If loopc = 0 Then
            Map_Light_Get = 0
            Exit Function

        End If

        loopc = loopc - 1
    Loop
    
    Map_Light_Get = loopc
    Exit Function
errorhandler:
    Map_Light_Get = 0
    
End Function

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    If grh_index > 0 And grh_index <= 40000 Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function


Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
    Dim c As Long
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    ARGB = c
End Function
Public Function Light_Remove_From_Pos(ByVal PosX As Byte, ByVal PosY As Byte, Optional vColor As Long = 0) As Boolean
    Dim i As Long
    
    For i = 1 To Light_last
        If Light_List(i).map_x = PosX Then
            If Light_List(i).map_y = PosY Then
                If vColor <> 0 Then
                    If Light_List(i).color = vColor Then
                        If Light_Check(i) Then
                            Light_Destroy i
                        End If
                    End If
                Else
                    If Light_Check(i) Then
                        Light_Destroy i
                    End If
                End If
            End If
        End If
    Next i
    
    Light_Render_All
End Function

Private Sub RoofAlphaCalculate(roofrgb_list() As Long)

    Dim color As D3DCOLORVALUE
    Static last_tick As Long
    If UserPos.X = 0 Or UserPos.Y = 0 Then Exit Sub
    
    
    If IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
            MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False) Then
            
            
        If CurrentUser.Muerto = True And AlphaY <> 0 Then
            AlphaY = 0
            Exit Sub
        End If
    
        If AlphaY > 0 Then 'Entra techo
        
            If GetTickCount - last_tick >= 18 Then
                AlphaY = AlphaY - 5
                last_tick = GetTickCount
            End If
                
            color = ambientLight
            color.a = AlphaY
            D3DColorToRgbList roofrgb_list(), color
            
            
        End If
    
    Else
    
        If AlphaY < 255 Then 'Sale techo

                
        If GetTickCount - last_tick >= 18 Then
            AlphaY = AlphaY + 5
            last_tick = GetTickCount
        End If

                
        color = ambientLight
        color.a = AlphaY
        D3DColorToRgbList roofrgb_list(), color
        
        
        End If
        
    End If

End Sub

Private Sub RoofAlphaCalculateToAlpha(roofrgb_list() As Long, ByVal charindex As Integer)
    Dim color As D3DCOLORVALUE
    Dim suma As Double
    suma = (timerTicksPerFrame * 6.7)
    
    
        
    If charlist(charindex).AlphaX < 30 Then
        If suma > 0 Then
            suma = suma / 10
        End If
    End If
    
    
   With charlist(charindex)
        If .State Then
            If .AlphaX > 0 Then
                If .AlphaX - suma <= 0 Then
                    .AlphaX = 0
                Else
                    .AlphaX = .AlphaX - suma
                End If
                .last_tick = GetTickCount
            Else
                .State = 0
                .AlphaX = 0
            End If
        Else
            If .AlphaX < 255 Then
                If .AlphaX + suma > 255 Then
                    .AlphaX = 255
                Else
                    .AlphaX = .AlphaX + suma
                End If
                .last_tick = GetTickCount
            Else
                .State = 1
                .AlphaX = 255
            End If
        End If
    End With
       
    color = ambientLight
    color.a = charlist(charindex).AlphaX
    D3DColorToRgbList roofrgb_list(), color
End Sub

Public Sub D3DColorToRgbList(rgb_list() As Long, color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(color.a, color.r, color.g, color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Function isMontura(ByVal body As Integer) As Boolean
    isMontura = (body <> 416 And body <> 415 And _
                 body <> 412 And body <> 381 And _
                 body <> 383 And body <> 384 And _
                 body <> 382 And body <> 413 And _
                 body <> 292 And body <> 291 And _
                 body <> 272 And body <> 291 And _
                 body <> 317)
End Function
 

Public Sub Inventory_Render()
    
    If Not RenderInv Then Exit Sub 'Dibujamos cuando es necesario
    
    Static re As Rect
    re.Left = 0
    re.Top = 0
    re.bottom = 160
    re.Right = 160
  
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    
    DirectDevice.BeginScene
    
    Inventario.DrawInventory
    
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
    Call DirectDevice.Present(re, ByVal 0, frmMain.picInv.hwnd, ByVal 0)
    
 End Sub
Public Function Map_Letter_Fade_Set(ByVal grh_index As Long, Optional ByVal after_grh As Long = -1) As Boolean
    If grh_index <= 0 Or grh_index = map_letter_grh.GrhIndex Then Exit Function
        
    If after_grh = -1 Then
        map_letter_grh.GrhIndex = grh_index
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = 0
    Else
        map_letter_grh.GrhIndex = after_grh
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = grh_index
    End If
    
    Map_Letter_Fade_Set = True
End Function

Public Function Map_Letter_UnSet() As Boolean
    map_letter_grh.GrhIndex = 0
    map_letter_fadestatus = 0
    map_letter_a = 0
    map_letter_grh_next = 0
    Map_Letter_UnSet = True
End Function
Public Sub Draw_Grh_Hdc(ByVal desthDC As Long, ByVal Grh As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal h_centered As Boolean, ByVal v_centered As Boolean)

    On Error GoTo Err
    
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    Dim grh_index As Integer
    Dim bRet As Boolean
    
    grh_index = Grh

    If grh_index <= 0 Then Exit Sub
    
    If GrhData(grh_index).NumFrames = 0 Then Exit Sub

    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If

    src_x = GrhData(grh_index).sX
    src_y = GrhData(grh_index).sY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
            
    If h_centered Then
        If GrhData(grh_index).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(grh_index).TileWidth * 16) + 16
        End If
    End If
    
    If v_centered Then
        If GrhData(grh_index).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32
        End If
    End If
    
    bRet = Extract_File(Graphics, App.Path & "\Recursos", GrhData(grh_index).FileNum & ".bmp", Resource_Path, False)
    
    If bRet Then
    
        hdcsrc = CreateCompatibleDC(desthDC)
        file_path = Resource_Path & GrhData(grh_index).FileNum & ".bmp"
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
        'BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, RGB(0, 0, 0)
        
        Call DeleteObject(SelectObject(hdcsrc, PrevObj))
        DeleteDC hdcsrc
        
    End If
    
    Delete_File (Resource_Path & GrhData(grh_index).FileNum & ".bmp")
    
    Exit Sub
    
Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Draw_Grh_Hdc", Erl)
    Resume Next
End Sub

Public Sub DrawPJ(ByVal Index As Integer)
    
    On Error GoTo ErrorHndle
    
1     Dim cColor As Long
    
3     'Index = Index
    
2     If cPJ(Index).Nombre <> "" Then
    
4         frmCharList.lblAccData(Index).Caption = cPJ(Index).Nombre
        
    
5         Select Case cPJ(Index).color
              Case 1 'Renegado
7                 cColor = RGB(175, 175, 175)
                
8             Case 2, 5 'Gris
9                 cColor = RGB(39, 131, 243)
                
10            Case 3, 6 'Repu Mili
11                cColor = RGB(243, 147, 1)
                
12            Case 4 'Rojo
13                cColor = RGB(217, 0, 5)

14            Case Else
            
15        End Select
16
17        If cPJ(Index).GameMaster Then cColor = RGB(0, 145, 72)
        
18        frmCharList.lblAccData(Index).ForeColor = cColor
        
19    End If
    
20    If cPJ(Index).Nombre = "" Then Exit Sub

21    Dim init_x As Integer
22    Dim init_y As Integer
    
23    Dim grhtemp As Grh

27    init_x = 25
28    init_y = 40

30    With cPJ(Index)
        
31        If .body <> 0 Then
32            grhtemp.GrhIndex = BodyData(.body).Walk(3).GrhIndex
33            Call Draw_Grh_Hdc(frmCharList.picChar(Index - 1).hDC, grhtemp.GrhIndex, init_x, init_y, True, True)
34        End If
            
          If .body = 84 Or .body = 85 Or .body = 86 Or .body = 87 Then Exit Sub
          
36        If .Head <> 0 Then
35            grhtemp.GrhIndex = HeadData(cPJ(Index).Head).Head(3).GrhIndex
37            Call Draw_Grh_Hdc(frmCharList.picChar(Index - 1).hDC, grhtemp.GrhIndex, init_x + BodyData(cPJ(Index).body).HeadOffset.X, init_y + BodyData(cPJ(Index).body).HeadOffset.Y, True, True)
38        End If
        
39        If .Helmet <> 0 Then
40            grhtemp.GrhIndex = CascoAnimData(cPJ(Index).Helmet).Head(3).GrhIndex
41            Call Draw_Grh_Hdc(frmCharList.picChar(Index - 1).hDC, grhtemp.GrhIndex, init_x + BodyData(cPJ(Index).body).HeadOffset.X, init_y + BodyData(cPJ(Index).body).HeadOffset.Y, True, True)
42        End If
        
43        If .Weapon <> 0 Then
44            grhtemp.GrhIndex = WeaponAnimData(cPJ(Index).Weapon).WeaponWalk(3).GrhIndex
45            Call Draw_Grh_Hdc(frmCharList.picChar(Index - 1).hDC, grhtemp.GrhIndex, init_x, init_y, True, True)
46        End If
        
47        If .Shield <> 0 Then
48            grhtemp.GrhIndex = ShieldAnimData(cPJ(Index).Shield).ShieldWalk(3).GrhIndex
50            Call Draw_Grh_Hdc(frmCharList.picChar(Index - 1).hDC, grhtemp.GrhIndex, init_x, init_y, True, True)
49        End If

51    End With

    Exit Sub

ErrorHndle:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.DrawPJ", Erl)
    Resume Next
End Sub
 
Public Sub DrawSkinPJ()
 
    Dim i As Integer
    
    Dim init_x As Integer
    Dim init_y As Integer
    Dim tempito(3) As Long
    Dim grhtemp As Grh
    Static re As Rect
    
    re.Left = 0
    re.Top = 0
    re.bottom = 80
    re.Right = 76
    
 
    
    If RSkin.body = 8 Then
        RSkin.Head = 500
    End If
    
    init_x = 25
    init_y = 40
        
    frmSkins.picChar.Cls
 
    With RSkin
    
 

        If RSkin.body <> 0 Then
        grhtemp.GrhIndex = BodyData(.body).Walk(3).GrhIndex
        Call DrawGrhtoHdc(frmSkins.picChar.hDC, grhtemp.GrhIndex, init_x, init_y, True, True, True)
        End If
        
If RSkin.body <> 84 And RSkin.body <> 85 And RSkin.body <> 86 And RSkin.body <> 87 Then
        
        
        If RSkin.Head <> 0 Then
            If Not (.body = 84 Or .body = 85 Or .body = 86 Or .body = 87) Then
            grhtemp.GrhIndex = HeadData(.Head).Head(3).GrhIndex
            Call DrawGrhtoHdc(frmSkins.picChar.hDC, grhtemp.GrhIndex, init_x + BodyData(.body).HeadOffset.X, init_y + BodyData(.body).HeadOffset.Y, True, True, True)
            End If
        End If
        
        If RSkin.Casco <> 0 Then
          grhtemp.GrhIndex = CascoAnimData(.Casco).Head(3).GrhIndex
          Call DrawGrhtoHdc(frmSkins.picChar.hDC, grhtemp.GrhIndex, init_x + BodyData(.body).HeadOffset.X, init_y + BodyData(.body).HeadOffset.Y, True, True, True)
        End If
        
        
        If RSkin.Weapon <> 0 Then
            grhtemp.GrhIndex = WeaponAnimData(.Weapon).WeaponWalk(3).GrhIndex
             Call DrawGrhtoHdc(frmSkins.picChar.hDC, grhtemp.GrhIndex, init_x, init_y, True, True, True)
        End If
        
        If RSkin.Shield <> 0 Then
            grhtemp.GrhIndex = ShieldAnimData(.Shield).ShieldWalk(3).GrhIndex
            Call DrawGrhtoHdc(frmSkins.picChar.hDC, grhtemp.GrhIndex, init_x, init_y, True, True, True)
        End If
End If

  
End With
       frmSkins.picChar.Refresh
End Sub
 Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bRain Then
        If bTecho Then
            If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                If RainBufferIndex Then _
                    Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                frmMain.IsPlaying = PlayLoop.plLluviain
            End If
        Else
            If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                If RainBufferIndex Then _
                    Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                frmMain.IsPlaying = PlayLoop.plLluviaout
            End If
        End If
    End If
    
    DoFogataFx
End Function
 
Public Sub Trueno_Render()

'    Static src_rect As Rect
'    Static dest_rect As Rect
'    Static temp_verts(3) As TLVERTEX
' '   Static light_value(0 To 3) As Long
    '
'    light_value(0) = 1179010815
'    light_value(1) = 2017871615
'    light_value(2) = 1179010815
'    light_value(3) = 910575359'''

    'Set up the source rectangle
'    With src_rect
'        .bottom = 0 + frmMain.MainViewPic.ScaleHeight
'        .Left = 0
'        .Right = 0 + frmMain.MainViewPic.ScaleWidth
'        .Top = 0
'    End With
'
'    'Set up the destination rectangle
'    With dest_rect
'        .bottom = 0 + frmMain.MainViewPic.ScaleHeight
'        .Left = 0
'        .Right = 0 + frmMain.MainViewPic.ScaleWidth
'        .Top = 0
'    End With
'
'    'Set up the TempVerts(3) vertices
'    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), 0, 0, 0
'
'    'Set Textures
 '   D3DDevice.SetTexture 0, Nothing
    
    'If alpha_blend Then
    '   'Set Rendering for alphablending
    '    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    '    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    'End If
    
    'Draw the triangles that make up our square Textures
   ' D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    'If alpha_blend Then
    '    'Set Rendering for colokeying
    '    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    '    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    'End If

End Sub

Public Sub Char_Refresh(ByVal charindex As Integer)
    If charlist(charindex).Pos.X = 0 Or charlist(charindex).Pos.Y = 0 Then Exit Sub
    MapData(charlist(charindex).Pos.X, charlist(charindex).Pos.Y).charindex = charindex
End Sub
Public Sub CrearFantasma(ByVal charindex As Integer)
    
    On Error GoTo CrearFantasma_Err
 
    
1      If charlist(charindex).Pos.Y = 0 Or charlist(charindex).Pos.X = 0 Then Exit Sub
3      If charlist(charindex).body.Walk(charlist(charindex).Heading).GrhIndex <= 0 Then Exit Sub
      
      With MapData(charlist(charindex).Pos.X, charlist(charindex).Pos.Y)
 
24    .CharFantasma.Heading = charlist(charindex).Heading
13    .CharFantasma.body.GrhIndex = charlist(charindex).body.Walk(charlist(charindex).Heading).GrhIndex
14    .CharFantasma.Head.GrhIndex = charlist(charindex).Head.Head(charlist(charindex).Heading).GrhIndex
15    .CharFantasma.Arma.GrhIndex = charlist(charindex).Arma.WeaponWalk(charlist(charindex).Heading).GrhIndex
16    .CharFantasma.Casco.GrhIndex = charlist(charindex).Casco.Head(charlist(charindex).Heading).GrhIndex
17    .CharFantasma.Escudo.GrhIndex = charlist(charindex).Escudo.ShieldWalk(charlist(charindex).Heading).GrhIndex

18    .CharFantasma.donador = charlist(charindex).donador
    
20    .CharFantasma.AlphaB = 255

21    .CharFantasma.Activo = True

22    .CharFantasma.OffX = charlist(charindex).body.HeadOffset.X
23    .CharFantasma.Offy = charlist(charindex).body.HeadOffset.Y
 
25    .CharFantasma.Nombre = charlist(charindex).Nombre
26    .CharFantasma.EsUsuario = charlist(charindex).EsUsuario

28    .CharFantasma.OffSetClan = charlist(charindex).OffSetClan

29    .CharFantasma.Clan = charlist(charindex).Clan

30    .CharFantasma.Arma_Aura = charlist(charindex).Arma_Aura
31    .CharFantasma.Body_Aura = charlist(charindex).Body_Aura
32    .CharFantasma.Escudo_Aura = charlist(charindex).Escudo_Aura
33    .CharFantasma.Head_Aura = charlist(charindex).Head_Aura
34    .CharFantasma.Otra_Aura = charlist(charindex).Otra_Aura
35    .CharFantasma.Anillo_Aura = charlist(charindex).Anillo_Aura
      
36    .CharFantasma.color(0) = charlist(charindex).color(0)
37    .CharFantasma.color(1) = charlist(charindex).color(1)
38    .CharFantasma.color(2) = charlist(charindex).color(2)
39    .CharFantasma.color(3) = charlist(charindex).color(3)

40    End With
    
    Exit Sub

CrearFantasma_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CrearFantasma", Erl)
    Resume Next
    
End Sub
Sub Map_Obj_Create(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer, ByVal OBJIndex As Integer, ByVal tipe As Byte, ByVal Amount As Integer, Optional ByVal fijo As Byte = 0)
'**************************************************************
'Author: Leandro Mendoza (Mannakia) & Mermas (Lo adapte que sea carga dinamica :p y que no haya que hardcodear cada item de luz
'Last Modify Date: 13/11/2010
'Set ObjGrh in array mapdata
'**************************************************************

If X = 0 Or Y = 0 Then Exit Sub

If fijo = 0 And tipe = 19 Then
    MapData(X, Y).particle_group = 0
    Call SetMapParticle(34, X, Y)
    Light_Create X, Y, &HFFFFFF, 2
    Light_Render_All
End If

MapData(X, Y).ObjGrh.GrhIndex = GrhIndex
MapData(X, Y).OBJInfo.OBJIndex = OBJIndex
MapData(X, Y).OBJInfo.EsFijo = fijo
MapData(X, Y).OBJInfo.tipe = tipe
MapData(X, Y).OBJInfo.Amount = Amount


InitGrh MapData(X, Y).ObjGrh, GrhIndex

If Len(General_Locale_Obj(OBJIndex, 9)) > 0 Then
    
    Light_Create X, Y, CLng(General_Locale_Obj(OBJIndex, 9)), (General_Locale_Obj(OBJIndex, 10))
    Light_Render_All
    
End If

End Sub
Sub Map_Obj_Delete(ByVal X As Integer, ByVal Y As Integer, Optional ByVal fijo As Boolean = False)
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 14/10/2010
'Set ObjGrh in array mapdata the nothing value
'**************************************************************

If X = 0 Or Y = 0 Then Exit Sub
If fijo = False Then
    If MapData(X, Y).OBJInfo.EsFijo = 1 Then
        Exit Sub
    End If
End If



If MapData(X, Y).particle_group > 0 Then
   If Particle_Get_Type(MapData(X, Y).particle_group) = 34 Or General_Locale_Obj(MapData(X, Y).OBJInfo.OBJIndex, 2) = 19 Then
       Call Particle_Group_Remove(MapData(X, Y).particle_group)
       Light_Remove_From_Pos X, Y, -1
   End If
End If

MapData(X, Y).ObjGrh.GrhIndex = 0


If Len(General_Locale_Obj(MapData(X, Y).OBJInfo.OBJIndex, 9)) > 0 Then

Dim color As Long
Dim r As Integer, g As Integer, b As Integer

color = CLng(General_Locale_Obj(MapData(X, Y).OBJInfo.OBJIndex, 9))

General_Long_Color_to_RGB color, r, g, b

color = D3DColorXRGB(r, g, b)
    
Light_Remove_From_Pos X, Y, color

End If


MapData(X, Y).OBJInfo.OBJIndex = 0

End Sub

Sub DrawGrhIndextoSurfaceAlpha(ByVal GrhIndex As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal Center As Byte, _
                          ByRef color() As Long, _
                          Optional ByVal angle As Single = 0, _
                          Optional ByVal AlphaB As Byte = 0)
                          
    With GrhData(GrhIndex)

        'Center Grh over X,Y pos

        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If

        'Draw
        Call Directx_Render_Texture(CLng(.FileNum), X, Y, .pixelHeight, .pixelWidth, .sX, .sY, color(), angle, AlphaB)

    End With

End Sub
Public Function ColorNombresPriv(ByVal charindex As Integer, ByVal priv As Byte)
            
    On Error GoTo hError
    
    With charlist(charindex)
    
        Dim long_color As Long
 
        Select Case charlist(charindex).priv
        
            Case 1: long_color = -9276564 'Renegado
            Case 2, 5: long_color = -14659077 'Ciudadano, Armi
            Case 3, 6: long_color = -3380480 'Repu, Mili
            Case 4: long_color = -3932155 'Caos
            Case 7, 8, 9: long_color = D3DColorXRGB(FontTypes(14).red, FontTypes(14).green, FontTypes(14).blue) 'Gms
            Case Else: long_color = -9276564
    
        End Select
    
        Engine_Long_To_RGB_List .color, long_color
        
    If .priv >= 7 And .priv <= 9 Then
        .EsGM = True
    Else
        .EsGM = False
    End If
    
    End With
                
    Exit Function
    
hError:

 Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.ColorNombresPriv", Erl)
Resume Next

End Function
Public Function CargarMedidasNombresModernos()

If NickModerno Then

NombresModernos = 5

NickModernoXX = 25
NickModernoX = 50

Else

NombresModernos = 1

NickModernoXX = 0
NickModernoX = 0

End If

End Function
Public Sub UserExpPerc()

On Error GoTo errorhandler

    If CurrentUser.UserExp > 0 And CurrentUser.UserPasarNivel > 0 Then
        CurrentUser.UserPercExp = CLng(CurrentUser.UserExp / (CurrentUser.UserPasarNivel / 100))
        If CurrentUser.UserPercExp = 100 Then CurrentUser.UserPercExp = 99
    Else
        CurrentUser.UserPercExp = 0
    End If

Exit Sub

errorhandler:

End Sub



Public Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
    
    '*********************************************
    'Draws a GRH transparently to a X and Y position
    '*****************************************************************
    On Error GoTo hError
    
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If
    
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    Call Directx_Render_Texture_Advance(CLng(GrhData(Grh.GrhIndex).FileNum), screen_x, screen_y, GrhData(Grh.GrhIndex).pixelHeight, GrhData(Grh.GrhIndex).pixelWidth, GrhData(Grh.GrhIndex).sX, GrhData(Grh.GrhIndex).sY, rgb_list, Width, Height, Grh.angle, alpha_blend)

    Exit Sub
    
hError:

 Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.Grh_Render_Advance", Erl)
    Resume Next

End Sub



Public Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)

    On Error GoTo hError
   
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
        
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If


    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
    
    If v_centered Then
        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * 32) + 32
        End If
    End If
    
    Call Directx_Render_Texture(CLng(GrhData(Grh.GrhIndex).FileNum), screen_x, screen_y, GrhData(Grh.GrhIndex).pixelHeight, GrhData(Grh.GrhIndex).pixelWidth, GrhData(Grh.GrhIndex).sX, GrhData(Grh.GrhIndex).sY, rgb_list(), Grh.angle, alpha_blend)
    

    Exit Sub
    
hError:

 Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.Grh_Render", Erl)
    Resume Next

End Sub
Public Function Engine_Scroll_Pixels(ByVal Valor As Single)
    scroll_pixels_per_frame = Valor
End Function


Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef color() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Shadow As Byte = 0)
    
    On Error GoTo hError
    
    Dim CurrentGrhIndex As Integer
    
    If Grh.GrhIndex = 0 Then Exit Sub
    If GrhData(Grh.GrhIndex).NumFrames = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * 0.5)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1 + Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If
    
    If CurrentUser.Muerto = True Then
          Static light_value(0 To 3) As Long
       ' light_value(0) = D3DColorXRGB(67, 67, 67)
       ' light_value(1) = D3DColorXRGB(67, 67, 67)
       ' light_value(2) = D3DColorXRGB(67, 67, 67)
       ' light_value(3) = D3DColorXRGB(67, 67, 67)
    End If

    Call Directx_Render_Texture(CLng(GrhData(CurrentGrhIndex).FileNum), X, Y, GrhData(CurrentGrhIndex).pixelHeight, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, color, angle, Alpha)
    
    Exit Sub
hError:

 Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.Draw_Grh", Erl)
    Resume Next

End Sub


