Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public Enum Atributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum
 
Public Default_RGB(0 To 3) As Long

'Hora
Public thFPSAndHour As Long

'Ver Map
Public VerLugar As Byte

'Nombre caption forms
Public Form_Caption As String

Public Const CentroInventario As Byte = 1
Public Const CentroHechizos As Byte = 1
Public Const CentroMenu As Byte = 1
Public Const Solapas As Byte = 4

Public ListaIgnorados As String


Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera
 


'
Public ClientTCP As clsClientTCP

'Web
Public Declare Function GetDesktopWindow Lib "user32" () As Integer
Public Const SW_Normal = 1

'Cargar IP/PUERTO
Public Const CurServerIP   As String = "192.168.1.34"
Public Const CurServerPort As Integer = 7666

'Daño
Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Crafteo
Public ArmasHerrero(1 To 16)       As Integer
Public ArmadurasHerrero(1 To 30)   As Integer
Public CascosHerrero(1 To 6)       As Integer
Public EscudosHerrero(1 To 9)      As Integer

Public ObjCarpintero(1 To 17)      As Integer
Public ObjAlquimia(1 To 9)         As Integer
Public ObjSastre(1 To 9)           As Integer

Public Const MAX_BANCOINVENTORY_SLOTS                    As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    south = 3
    WEST = 4
End Enum

'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1

'Objetos
Public Const MAX_INVENTORY_OBJS         As Integer = 10000
Public Const MAX_INVENTORY_SLOTS        As Byte = 25
Public Const MAX_NPC_INVENTORY_SLOTS    As Byte = 50
Public Const INV_OFFER_SLOTS            As Byte = 20
Public Const INV_GOLD_SLOTS             As Byte = 1
Public Const MAXHECHI                   As Byte = 35

Public Const NUMSKILLS                            As Byte = 27
Public Const NUMATRIBUTOS                         As Byte = 5
Public Const NUMCLASES                            As Byte = 18
Public Const NUMRAZAS                             As Byte = 6


Public Const FLAGORO                    As Integer = MAX_INVENTORY_SLOTS + 1
 
Public Const GOLD_OFFER_SLOT            As Integer = INV_OFFER_SLOTS + 1
Public Const FOgata                     As Integer = 1521
Public Const NUMCIUDADES                As Byte = 2

Public Const FundirMetal                As Integer = 88

'Inventario
Type Inventory
    OBJIndex  As Integer
    Name      As String
    GrhIndex  As Integer
    Amount    As Long
    Equipped  As Byte
    Valor     As Single
    ObjType   As Integer
    MaxDef    As Integer
    MinDef    As Integer
    MaxHit    As Integer
    MinHit    As Integer
    PuedeUsar As Byte
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Single
    ObjType As Integer
    MinDef As Integer
    MaxDef As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    UsaItemNPC As Byte
End Type

 Type tHeadRange
    mStart As Integer
    mEnd As Integer
    fStart As Integer
    fEnd As Integer
End Type

Public Ciudades(1 To NUMCIUDADES)                 As String
 
Public ListaRazas() As String
 
Public ListaClases() As String
Public Head_Range() As tHeadRange

Public UserHechizos(1 To MAXHECHI)                As Integer
Global OtroInventario(1 To MAX_INVENTORY_SLOTS)   As Inventory

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV


Public SkillPoints                                As Integer
Public Alocados                                   As Integer
Public Flags()                                    As Integer

 
 


'Mermas Carga Recursos

'LLuvia
Public meteo_particle As Integer
Public trueno As Byte
Public Queclima As Byte


'Mouse location
Public TX                  As Byte
Public TY                  As Byte
 
 

'seguridad
Public Security As New clsSecurity

 
'Click usuario tag [TARGET]
Public UserFichado As Integer
 

'Renderizar inventario
Public RenderInv As Boolean
 
'Velocidad Montura
Public Velocidades As Byte
 

'Cursores graficos
Public FormParser As clsCursor
 
 
Public NUMFONTS As Integer

Public FontTypes() As tFontType

Public Type tFontType
    bold As Boolean
    italic As Boolean
    red As Integer
    green As Integer
    blue As Integer
End Type

Public Const FONTTYPE_SERVER As Integer = 8
Public Const FONTTYPE_TALK As Byte = 1
Public Const FONTTYPE_GUILDMSG As Byte = 17
Public Const FONTTYPE_PIEL As Byte = 11
Public Const FONTTYPE_PIEL2 As Byte = 12

Public Type tCurrentUser
Creditos As Long

Montando As Boolean
UserMaxAGU As Byte
UserMinAGU As Byte
UserMaxHAM As Byte
UserMinHAM As Byte

UserLvl As Integer
TiempoSalida As Boolean
SendingType As Byte
sndPrivateTo As String
Logged As Boolean
RenderGM As Boolean
AutoNavigation As Boolean
Ping As Long
PingRequested As Boolean
Muerto As Boolean
UserDescansar As Boolean
UserAtributos(1 To NUMATRIBUTOS) As Integer
UserSkills(1 To NUMSKILLS) As Byte
LastItem As Integer
UserGLD As Long
UsingSkill As Integer
UserCharIndex As Integer
UserExp As Long
UserPasarNivel As Long
UserPercExp As Long
UserMap As Integer
UserMinHP As Integer
UserMaxHP As Integer
UserMaxSTA As Integer
UserMinSTA As Integer
UserMaxMAN As Integer
UserMinMAN As Integer
LogeoAlgunaVez As Boolean
Nivel As Integer
CurMapBattle As Byte
End Type

Public CurrentUser As tCurrentUser

'Condicionales para el frmPregunta
Public RetirarFacciones As Boolean
Public EstasMuerto      As Boolean
Public CantidadGlobal As Long
Public LinkConsola As Boolean

 
 
'MiniMap Lectura de Pixel
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Old fashion BitBlt function
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long


'Objetos públicos
Public Audio                                      As New clsAudio
Public Inventario                                 As New clsGrapchicalInventory
Public Inventory(1 To MAX_INVENTORY_SLOTS) As Inventory

'Inventarios de herreria
Public Const MAX_LIST_ITEMS                       As Byte = 4

Public SurfaceDB                                  As New clsTextureManager
Public CustomKeys                                 As New clsCustomKeys
Public incomingData                               As New clsByteQueue
Public outgoingData                               As New clsByteQueue

''
'The main timer of the game.
Public MainTimer                                  As New clsTimer


'Sonidos
Public Const SND_CLICK               As String = "190.Wav"
Public Const SND_INFO               As String = "385.Wav"

Public Const SND_PASOS1              As String = "23.Wav"
Public Const SND_PASOS2              As String = "24.Wav"
Public Const SND_PASOS3              As String = "201.Wav"
Public Const SND_PASOS4              As String = "202.Wav"
Public Const SND_PASOS5              As String = "197.Wav"
Public Const SND_PASOS6              As String = "198.Wav"
Public Const SND_PASOS7              As String = "199.Wav"
Public Const SND_PASOS8              As String = "200.Wav"


' Constantes de intervalo
Public Const INT_MACRO_HECHIS        As Integer = 1400
Public Const INT_MACRO_TRABAJO       As Integer = 1000
Public Const INT_ATTACK              As Integer = 1200
Public Const INT_ARROWS              As Integer = 1050
Public Const INT_CAST_SPELL          As Integer = 500
Public Const INT_CAST_ATTACK         As Integer = 500
Public Const INT_WORK                As Integer = 700
Public Const INT_USEITEMU            As Integer = 300
Public Const INT_USEITEMDCK          As Integer = 1000
Public Const INT_SENTRPU             As Integer = 2000


'Constantes de graficos
Public Const CASPER_HEAD             As Integer = 500
Public Const FRAGATA_FANTASMAL       As Integer = 87
 

Public CreandoClan        As Boolean
Public ClanName           As String
Public Site               As String
Public UserCiego          As Boolean
Public UserEstupido       As Boolean
 
Public RainBufferIndex    As Long
Public FogataBufferIndex  As Long
 
Public UsaMacro                                          As Boolean
Public TradingUserName                                   As String
Public NumberOfCharacters As Byte
Public PremiosInv(1 To 20)              As PremiosList
 
Public Enum eClass
    Clerigo = 1
    Mago = 2
    Guerrero = 3
    Asesino = 4
    Ladron = 5
    Bardo = 6
    Druida = 7
    Gladiador = 8
    Paladin = 9
    Cazador = 10
    Pescador = 11
    Herrero = 12
    Leñador = 13
    Minero = 14
    Carpintero = 15
    Sastre = 16
    Mercenario = 17
    Nigromante = 18
End Enum

Public Enum eCiudad
    cNix = 1
    cilliandor
    cUllathorpe
    cBanderbill
    cRinkel
    cDungeonNewbie
    cLindos
    cArghal
    cTiama
    cOrac
    cSuramei
    cNueva
    cPrision
    cLibertad
    cIntermundia
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ELFOOSCURO
    gnomo
    enano
    Orco
End Enum

Public Type tFont
Font_size As Integer
Ascii_code(0 To 255) As Integer
End Type

Public font_types(1 To 3) As tFont

Public Enum eSkill
    Suerte = 13 'Es MUSICA EN EL SKILLS DE IAO, PERO AC ES SUERTE
    magia = 8
    robar = 14
    Tacticas = 1
    armas = 2
    Meditar = 10
    Apuñalar = 4
    Ocultarse = 11
    Supervivencia = 16
    talar = 20
    Comerciar = 15
    Defensa = 7
    pesca = 18
    mineria = 19
    Carpinteria = 23
    Herreria = 22
    Liderazgo = 17
    domar = 12
    proyectiles = 6
    Wrestling = 3
    Navegacion = 26
    ResistenciaMagica = 9
    armasarrojadizas = 5
    Alquimia = 24
    Botanica = 21
    Sastreria = 25
    Equitacion = 27
End Enum

Enum eGenero
    Hombre = 1
    Mujer = 2
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otpociones = 11
    otLibros = 12
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otAnillo = 18
    otTeleport = 19
    otMuebles = 20
    otItemsMagicos = 21
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otLingotes = 29
    otPieles = 30
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otArbolElfico = 36
    otMapa = 37
    otRuna = 38
    otBolsas = 39 ' Blosas de Oro  (contienen más de 10k de oro)
    otPozos = 40 'Pozos Mágicos
    otEsposas = 41
    otRaíces = 42
    otCadáveres = 43
    otMonturas = 44
    otPuestos = 45 ' Puestos de Entrenamiento
    otNudillos = 46
    otAnillos = 47
    otcorreo = 48
    otAnilloEspec = 49
    otInvi = 50
    otRegalos = 53
    otCualquiera = 1000

End Enum


Public Enum eGMCommands

    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    GoNearby                '/IRCERCA
    Comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpChar                '/TELEP
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    Invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN /UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    nickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    MakeDumbNoMore          '/NOESTUPIDO
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    RemoveCharFromGuild     '/RAJARCLAN
    AlterPassword           '/APASS
    ToggleCentinelActivated '/CENTINELAACTIVADO
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    DARPUN                  '/DARPUN NOMBRE@CANTIDAD
    ResponderGM
    donador
    EventoOro
    EventoExperiencia
    CuentaRegresiva
End Enum
 

Type PremiosList
    Name As String
    Puntos As Integer
End Type

 

Type tCorreo
    Mensaje   As String
    De        As String
    GrhIndex  As Integer
    Cantidad  As Integer
    Nombre    As String
    OBJIndex  As Integer
    Leido     As Byte
End Type
    
Public Correos(1 To 10) As tCorreo
 

Type tEstadisticasUsu
    RepublicanosMatados     As Long
    ArmadasRealesMatados    As Long
    MiliciasMatados         As Long
    CaosMatados             As Long
    CiudadanosMatados       As Long
    CriminalesMatados       As Long
    UsuariosMatados         As Long
    NpcsMatados             As Long
    Clase                   As String
    Raza                    As Byte
    Genero                  As Byte
    MuertesUsuario          As Long
    status                  As Byte
    
End Type
 
Public Nombres                                    As Boolean

'Types de Usuario

Public Type tCuenta
    
    UserName As String
    UserAccount As String
    UserPassword As String
    UserCode As String
    EsChange As Byte
    
End Type

Public Cuenta As tCuenta

Public UserMeditar                                As Boolean
  
 Public UserPort                                   As Integer
Public UserServerIP                               As String
 

Public UserEstadisticas                           As tEstadisticasUsu
 
Public pausa                                      As Boolean
Public IScombate As Boolean
Public isSeguro As Boolean
Public FPSFLAG As Boolean
Public UserParalizado                             As Boolean
Public UserNavegando                              As Boolean
Public UserHogar                                  As eCiudad
 
Public UserWeaponEqpSlot                          As Byte
Public UserArmourEqpSlot                          As Byte
Public UserHelmEqpSlot                            As Byte
Public UserShieldEqpSlot                          As Byte

'<-------------------------NUEVO-------------------------->
Public Comerciando                                As Boolean
Public MirandoAsignarSkills                       As Boolean
Public MirandoEstadisticas                        As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase                                  As eClass
Public UserSexo                                   As eGenero
Public UserRaza                                   As eRaza


'Shermie80; Esto es un dolor de cabeza breo
 


Public PorcentajeSkills(1 To NUMSKILLS)           As Byte
Public SkillsNames() As String
Public SkillsOrig(1 To NUMSKILLS)                 As Byte

 
Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearNuevaCuenta = 4
    RecuperarCuenta = 5
    ConectarPersonaje = 6
    CambiarContraseña = 7
    BorrarPersonaje = 8
End Enum

Public EstadoLogin As E_MODO

Public Enum eClanType

    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
    ct_Milicia

End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Sex
    eo_Raza
    eo_addGold

End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger

    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6

End Enum

'Server stuff
Public stxtbuffer      As String 'Holds temp raw data from server
Public stxtbuffercmsg  As String 'Holds temp raw data from server
Public Connected       As Boolean 'True when connected to server


'Control
Public prgRun          As Boolean 'When true the program ends
Public FinPres As Boolean

'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long
 
''[END]''
'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Integer

End Type

Public Type tIndiceCuerpo

    body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Integer
    OffSetX As Integer
    OffSetY As Integer

End Type
 
Public GuildNames()      As String
Public GuildMembers()    As String
 
Public Const ORO_INDEX         As Integer = 12
Public Const ORO_GRH           As Integer = 511
 
