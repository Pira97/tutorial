Attribute VB_Name = "Declaraciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public tHora As Byte
Public tMinuto As Byte
Public tSeg As Byte

    Public Expc As Integer
    Public Oroc As Integer
    
    Public Administradores                              As clsIniManager
    
'seguridad
 Public Security As New clsSecurity
'Grabado de pj

Public NUMCIUDADES As Integer

Public INTERVALO_AUTO_GP As Byte
Public Queclima As Byte
'mermas mulitiplicadores
Public multiplicadorOro As Byte
Public OroModificada As Boolean
Public multiplicadorExp As Byte
Public ExpModificada As Boolean
Public CuentaRegresivaTimer As Byte
Public EventosOroandExp As Integer

''
' Modulo de declaraciones. Aca hay de todo.
'

Public aDos As New clsAntiDos

 Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 60
Public Const INFINITE_LOOPS As Integer = -1
Public Const FXSANGRE = 14
Public Const fxvenenopota = 2
Public Const MAXAMIGOS As Byte = 5
''
' The color of chats over head of dead characters.
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0

''
' The color of yells made by any kind of game administrator.
Public Const CHAT_COLOR_GM_YELL   As Long = &HF82FF
''
' Coordinates for normal sounds (not 3D, like rain)
Public Const NO_3D_SOUND          As Byte = 0

' Cantidad maximo de correos
 
'Grh en un Enum
Public Enum iGraficos
iBarca = 84
iGalera = 85
iGaleon = 86
iFragataFantasmal = 87
End Enum


Public Enum iMinerales

    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388

End Enum

Public Enum PlayerType

    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20

End Enum

'  $ CLASES $
' $ Shermie80 $

Public Enum eClass
    Clerigo = 1
    Mago = 2
    Guerrero = 3
    Asesino = 4
    ladron = 5
    Bardo = 6
    Druida = 7
    Gladiador = 8 'Cazarecompensas
    Paladin = 9
    Cazador = 10
    PescadoR = 11
    Herrero = 12
    Leñador = 13
    Minero = 14
    Carpintero = 15
    Sastre = 16
    Mercenario = 17 'Drakkar
    Nigromante = 18
End Enum

' $ Shermie80 $
'  $ CLASES $


Public Enum eCiudad

    cNix = 1
    cIlliandor = 2
    cUllathorpe = 3
    cBanderbill = 4
    cRinkel = 5
    cDungeonNewbie = 6
    cLindos = 7
    cARGHAL = 8
    cTIAMA = 9
    cORAC = 10
    cSURAMEI = 11
    cNueva = 12
    cPrision = 13
    cLibertad = 14
    cIntermundia = 15
    
End Enum

Public Enum eRaza

    Humano = 1
    Elfo
    Drow
    gnomo
    enano
    Orco
    
End Enum

Enum eGenero

    Hombre = 1
    Mujer

End Enum

Public Enum eClanType

    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
    ct_Milicia

End Enum

Public Const LimiteNewbie As Byte = 13


Public Type tCabecera 'Cabecera de los con

    desc As String * 255
    crc As Long
    MagicWord As Long

End Type

Public MiCabecera                    As tCabecera

'Barrin 3/10/03
'Cambiado a 2 segundos el 30/11/07
Public Const TIEMPO_INICIOMEDITAR    As Integer = 2000

Public Const NingunEscudo            As Integer = 2
Public Const NingunCasco             As Integer = 2
Public Const NingunArma              As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 402
Public Const RYKAN As Integer = 1601
Public Const SLOTS_POR_FILA          As Byte = 5


Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs

    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARXXGRANDE = 34

End Enum

Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

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

    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6

End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6

    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3

End Enum

'TODO : Reemplazar por un enum
Public Const Bosque   As String = "BOSQUE"
Public Const Nieve    As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"
Public Const Ciudad   As String = "CIUDAD"
Public Const Campo    As String = "CAMPO"
Public Const Dungeon  As String = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType

    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4

End Enum

' <<<<<< Acciona sobre >>>>>>

Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uInvocacion = 4
    uCreateTelep = 5
    uFamiliar = 6
    uMaterializa = 7
    uPropEsta = 8
    uCalmacion = 9
    uCreateMagic = 10
    uEquipamiento = 11
    uDetectarInvis = 12
End Enum



Public Const MAXUSERHECHIZOS      As Byte = 35

' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoBotanicaGeneral As Byte = 4
Public Const EsfuerzoBotanicaDruida As Byte = 2

Public Const EsfuerzoPescarPescador        As Byte = 1
Public Const EsfuerzoPescarGeneral         As Byte = 3

Public Const EsfuerzoExcavarMinero         As Byte = 2
Public Const EsfuerzoExcavarGeneral        As Byte = 5

Public Const FX_TELEPORT_INDEX             As Integer = 1

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo

    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6

End Enum

Public Const Guardias                       As Integer = 6
Public Const MAX_ORO_EDIT                   As Long = 5000000

Public Const STANDARD_BOUNTY_HUNTER_MESSAGE As String = _
        "Se te ha otorgado un premio por ayudar al proyecto reportando bugs, el mismo está disponible en tu bóveda."

Public Const MAXREP                         As Long = 6000000
Public Const MAXORO                         As Long = 90000000
Public Const MAXEXP                         As Long = 99999999

Public Const MAXUSERMATADOS                 As Long = 65000

Public Const MAXATRIBUTOS                   As Byte = 35
Public Const MINATRIBUTOS                   As Byte = 35

Public Const LingoteHierro                  As Integer = 386
Public Const LingotePlata                   As Integer = 387
Public Const LingoteOro                     As Integer = 388
Public Const Leña                           As Integer = 58
Public Const Raiz                           As Integer = 888

Public Const MAXNPCS  As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const HACHA_LEÑADOR As Integer = 127
Public Const PIQUETE_MINERO      As Integer = 187

Public Const DAGA                As Integer = 15
Public Const FOGATA_APAG         As Integer = 136
Public Const FOGATA              As Integer = 63
Public Const ORO_MINA            As Integer = 194
Public Const PLATA_MINA          As Integer = 193
Public Const HIERRO_MINA         As Integer = 192
Public Const MARTILLO_HERRERO    As Integer = 389
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const OLLA                As Integer = 887
Public Const TIJERAS             As Integer = 885
Public Const COSTURERO           As Integer = 886
Public Const ObjArboles          As Integer = 4
Public Const RED_PESCA           As Integer = 138
Public Const CAÑA_PESCA          As Integer = 881

Public Const PielLobo            As Integer = 414
Public Const PielOso             As Integer = 415
Public Const PielOsoPolar        As Integer = 1145

Public Enum eNPCType '

    Comun = 0
    Revividor = 1
    GuardiasCity = 2
    Entrenador = 3
    Banquero = 4
    facciones = 5
    BlancosCombate = 6
    transportadores = 7
    Veterinarias = 11
    Timbero = 12
    Subastadores = 16
    Convertidores = 18
    Shop = 19
    Dragon = 20
End Enum

Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS      As Byte = 27

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS   As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES      As Byte = 18

''
' Cantidad de Razas
Public Const NUMRAZAS       As Byte = 6

''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS   As Byte = 3
   
'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500

Public Const iORO          As Byte = 12
Public Const Pescado       As Byte = 139

Public Enum PECES_POSIBLES

    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546

End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
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
    comerciar = 15
    Defensa = 7
    pesca = 18
    mineria = 19
    Carpinteria = 23
    Herreria = 22
    Liderazgo = 17
    domar = 12
    Proyectiles = 6
    Wrestling = 3
    Navegacion = 26
    Resistencia = 9
    ArmasArrojadizas = 5
    alquimia = 24
    botanica = 21
    Sastreria = 25
    Equitacion = 27
End Enum
'%%%%%%%%%%%%%%%% Shermie80 %%%%%%%%%%%%%%%%%%%%%

Public Const FundirMetal = 88

Public Enum eAtributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5

End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador  As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef        As Byte = 15
Public Const AumentoStBandido    As Byte = AumentoSTDef + 23
Public Const AumentoSTLadron     As Byte = AumentoSTDef + 3
Public Const AumentoSTMago       As Byte = AumentoSTDef - 1
Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize         As Byte = 100
Public Const XMinMapSize         As Byte = 1
Public Const YMaxMapSize         As Byte = 100
Public Const YMinMapSize         As Byte = 1

'Tamaño del tileset
Public Const TileSizeX           As Byte = 32
Public Const TileSizeY           As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow             As Byte = 17
Public Const YWindow             As Byte = 13


'//Mermas, meto las constantes acá, mas orden

'Sonidos
Public Enum Sonidos
SND_SWING = 2
SND_WARP = 3
SND_CLICK = 190
SND_PUERTA = 5
SND_NIVEL = 6
SND_IMPACTO3 = 10
SND_USERMUERTE = 11
SND_LEÑADOR = 13
SND_TALAR = 13
SND_PESCAR = 14
SND_MINERO = 15
SND_SACARARMA = 25
SND_ESCUDO = 37
MARTILLOHERRERO = 41
LABUROCARPINTERO = 42
SND_SANAR = 55
SND_RESUCITAR = 84
SND_IMPACTO = 86
SND_DROP = 132
SND_BEBER = 135
SND_FALLASFLECHA = 145
snd_casamiento = 161
SND_ORO2 = 172
SND_RESUCITADO = 204
SND_VENENO = 239
End Enum
 
 

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS         As Integer = 10000

 
''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 25

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO                    As Integer = MAX_INVENTORY_SLOTS + 1

' CATEGORIAS PRINCIPALES
Public Enum eOBJType

    otUseOnce = 1           '  1 Comidas
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otPociones = 11
    otLibros = 12
    otbebidas = 13
    otLeña = 14
    otFuego = 15
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
    otManchas = 35   'No se usa
    otPasajes = 36
    otRuna = 38
    otMapa = 37
    otBolsas = 39 ' Blosas de Oro  (contienen más de 10k de oro)
    otPozos = 40 'Pozos Mágicos
    otEsposas = 41
    otRaíces = 42
    otCadáveres = 43
    otMonturas = 44
    otPuestos = 45 ' Puestos de Entrenamiento
    otNudillos = 46
    otAnillos = 47
    otCorreo = 48
    otAnilloEspec = 49
    otInvi = 50
    otRegalos = 53
    otCualquiera = 1000
    
End Enum

'Texto
Public Const FONTTYPE_TALK            As String = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT           As String = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING         As String = "~32~51~223~1~1"
Public Const FONTTYPE_INFO            As String = "~204~193~115~0~1"
Public Const FONTTYPE_INFOBOLD        As String = "~65~190~156~1~0"
Public Const FONTTYPE_EJECUCION       As String = "~130~130~130~1~0"
Public Const FONTTYPE_VENENO          As String = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD           As String = "~255~255~255~1~0"
Public Const FONTTYPE_SERVER          As String = "~0~185~0~0~0"
Public Const FONTTYPE_GUILDMSG        As String = "~228~199~27~0~0"
Public Const FONTTYPE_CONSEJO         As String = "~130~130~255~1~0"
Public Const FONTTYPE_CONSEJOCAOS     As String = "~255~60~00~1~0"
Public Const FONTTYPE_CONSEJOVesA     As String = "~0~200~255~1~0"
Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
Public Const FONTTYPE_CENTINELA       As String = "~0~255~0~1~0"
Public Const FONTTYPE_INFOBOLD2       As String = "~65~190~156~0~1"
Public Const FONTTYPE_INFOBOLD3       As String = "~182~226~29~0~1"
Public Const FONTTYPE_INFOBOLD4       As String = "~220~124~4~0~1"
Public Const FONTTYPE_LETRADIOS       As String = "~2~162~38~1~0"
Public Const FONTTYPE_LETRASEMIDIOS   As String = "~193~159~69~1~0"
Public Const FONTTYPE_IMPERIAL        As String = "~32~81~251~1~0"
Public Const FONTTYPE_RENEGADO        As String = "~114~115~108~1~0"
Public Const FONTTYPE_REPUBLICANO     As String = "~204~107~0~1~0"
Public Const FONTTYPE_MILICIANO       As String = "~204~107~0~1~0"
Public Const FONTTYPE_FUERZASCAOS     As String = "~196~0~15~1~0"
Public Const FONTTYPE_LETRACONSEJERO  As String = "~0~170~228~1~0"
Public Const FONTTYPE_GRITAR     As String = "~200~25~25~0~0"

Public CantPremios                    As Integer
'Estadisticas
Public Const STAT_MAXELV              As Byte = 50
Public Const STAT_MAXHP               As Integer = 999
Public Const STAT_MAXSTA              As Integer = 999
Public Const STAT_MAXMAN              As Integer = 9999
Public Const STAT_MAXHIT_UNDER36      As Byte = 99
Public Const STAT_MAXHIT_OVER36       As Integer = 999
Public Const STAT_MAXDEF              As Byte = 99

Public Const ELU_SKILL_INICIAL        As Byte = 200
Public Const EXP_ACIERTO_SKILL        As Byte = 50
Public Const EXP_FALLO_SKILL          As Byte = 20

' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type Position

    X As Integer
    Y As Integer

End Type

Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

    Dead_Map As Integer
    Dead_X As Integer
    Dead_Y As Integer

End Type

Public Type tHechizo
    Nombre As String
    desc As String
    PalabrasMagicas As String
    TimeParticula As Integer
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    HechizoDeArea As Byte
    AreaEfecto As Byte
    Afecta As Byte
    
    '    Resis As Byte
    
    Tipo As TipoHechizo
    
    MaterializaX As Integer
    MaterializaObj As Integer
    MaterializaCant As Integer
    WAV As Integer
    FXgrh As Integer
    Particle As Integer
    Loops As Byte
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer

    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Incinera As Byte
    Estupidez As Byte
    Ceguera As Byte
    AutoLanzar As Byte
    Revivir As Byte
    ResucitaFamiliar As Byte
    RemueveInvisibilidadParcial As Byte
    AfectaArea As Byte
    AreaX As Byte
    AreaY As Byte
    extrahit As Byte
    Metamorfosis As Byte
    body As Integer
    Desencantar As Byte
    Warp As Byte
    Invoca As Byte
    NumNpc As Integer
    cant As Integer
    Sanacion As Byte
    '    Materializa As Byte
    '    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    'Barrin 29/9/03
    StaRequerido As Integer

    Target As TargetType
    Anillo As Byte
    
End Type

Public Type LevelSkill

    LevelValue As Integer

End Type

Public Type UserObj

    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
    ProbTirar As Byte

End Type

'Objetos nacimiento Dinamico, manejado desde dat :p //Mermas
Public Type TNacimiento
    ObjIndex() As Integer
    Amount() As Integer
    Equipped() As Integer
    
End Type

Public Type TOBJNacimiento
    Clase(1 To NUMCLASES) As TNacimiento
    
End Type

Public OBJNacimiento As TOBJNacimiento
'End obj dinamico


Public Type Inventario

    Object(1 To MAX_INVENTORY_SLOTS) As UserObj
    NudiEqpObjIndex As Integer
    NudiEqpSlot As Byte
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    MagicIndex As Integer
    MagicSlot As Integer
    
End Type


Public Type FXdata

    Nombre As String
    GrhIndex As Integer
    Delay As Integer

End Type

Public Enum eMagicType
    ResistenciaMagica = 1
    ModificaAtributo = 2
    ModificaSkill = 3
    AceleraVida = 4
    AceleraMana = 5
    AumentaGolpe = 6
    DisminuyeGolpe = 7
    Nada = 8
    MagicasNoAtacan = 9
    Incinera = 10
    Paraliza = 11
    CarroMinerales = 12
    CaminaOculto = 13
    DañoMagico = 14
    Sacrificio = 15
    Silencio = 16
    NadieDetecta = 17
    Experto = 18
    Envenena = 19
End Enum


'Datos de user o npc
Public Type Char

    CharIndex As Integer
    Head As Integer
    body As Integer
    Donador As Byte
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    Particles(1 To 15) As Integer
    FX As Integer
    Loops As Integer
    
    heading As eHeading
    ParticulaFx As Integer
    
    WeaponAnimSkin As Byte
    ShieldAnumSkin As Byte
    CascoAnimSkin As Byte
    BodySkin As Byte
    ObjetoSkin As Byte
    
    'Auras
    Arma_Aura As Byte
    Body_Aura As Byte
    Escudo_Aura As Byte
    Head_Aura As Byte
    Otra_Aura As Byte
    Anillo_Aura As Byte
End Type

'Tipos de objetos
Public Type ObjData

    QueAtributo As Byte
    EfectoMagico As eMagicType
    CuantoAumento As Byte
    QueSkill As Byte
    
    levelItem          As Byte

    Velocidades        As Byte 'Velocidad de Monturas
    Aura               As Byte
    SndEspecial        As Integer
    Name               As String 'Nombre del obj
   
    OBJType            As eOBJType 'Tipo enum que determina cuales son las caract del obj
   
    GrhIndex           As Integer ' Indice del grafico que representa el obj
    GrhSecundario      As Integer
   
    'Solo contenedores
    MaxItems           As Integer
    Apuñala            As Byte
   
    HechizoIndex       As Integer
   
    MinHP              As Integer ' Minimo puntos de vida
    MaxHP              As Integer ' Maximo puntos de vida
   
    SubTipo            As Integer
    Numero             As Integer
    LanzaHechizo       As Integer
    MineralIndex       As Integer
    LingoteInex        As Integer
   
    proyectil          As Integer
    Municion           As Integer
    TieneSkin()        As Integer
    CantidadSkin       As Integer

    Crucial            As Byte
    Newbie             As Integer
    Destruir           As Byte
 
    DesdeMap           As Long
    HastaMap           As Long
    HastaY             As Byte
    HastaX             As Byte
    NecesitaSkill      As Byte
    CantidadSkill      As Byte
    
    MaxModificador     As Integer
    MinModificador     As Integer
    DuracionEfecto     As Long
    MinSkill           As Integer
    LingoteIndex       As Integer
   
    Shop               As Byte 'Items de Shop
   
    MinHIT             As Integer 'Minimo golpe
    MaxHIT             As Integer 'Maximo golpe
   
    MinHam             As Integer
    MinSed             As Integer
   
    MinDef             As Integer ' Armaduras
    MaxDef             As Integer ' Armaduras
   
    Ropaje             As Integer 'Indice del grafico del ropaje
    
    WeaponAnim         As Integer ' Apunta a una anim de armas
    ShieldAnim         As Integer ' Apunta a una anim de escudo
    CascoAnim          As Integer ' Apunta a una anim de casco
   
    Valor              As Long     ' Precio
   
    Cerrada            As Integer
    Llave              As Byte
    clave              As Long 'si clave=llave la puerta se abre o cierra
   
    IndexAbierta       As Integer
    IndexCerrada       As Integer
    IndexCerradaLlave  As Integer
    MinELV             As Byte

    Mujer              As Byte
    Hombre             As Byte
   
    Envenena           As Byte
    Paraliza           As Byte
   
    Agarrable          As Byte
    StaRequerido       As Byte
    LingH              As Integer
    LingO              As Integer
    LingP              As Integer
    
    Madera             As Integer
    Raies              As Integer
    
    PielLobo           As Integer
    PielOsoPardo       As Integer
    PielOsoPolar       As Integer
   
    SkPociones         As Integer
    SkSastreria        As Integer
    SkHerreria         As Integer
    SkCarpinteria      As Integer
   
    texto              As String
   
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    
    RazaProhibida(1 To NUMRAZAS) As eRaza
   
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    'Items Faccionarios
    Real              As Integer
    Caos              As Integer
    Milicia           As Integer
    '
    
    ResistenciaMagica As Integer
    NoSeCae As Integer

    puntos As Integer
   
    Radio As Integer
    
    DosManos As Byte
End Type

Public Type Obj

    ObjIndex As Integer
    Amount As Integer

End Type

'[Pablo ToxicWaste]
Public Type ModClase

    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double
    AtaqueArpon As Double
    DañoArpon As Double
End Type

Public Type ModRaza

    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single

End Type

'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario

    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserObj
    NroItems As Integer

End Type

'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
 
Public Type tPremios
    puntos As Long
    ObjIndex As Long
End Type

'Estadisticas de los usuarios
Public Type UserStats

    GLD As Long 'Dinero
    Banco As Long
    
    MaxHP As Integer
    MinHP As Integer
    
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    MaxAGU As Integer
    MinAGU As Integer
        
    def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer

    ExpSkills(1 To NUMSKILLS) As Long
    EluSkills(1 To NUMSKILLS) As Long
    
End Type

'Donador
Public Type TDonador
    activo As Byte
    CreditoDonador As Long
End Type


'Flags
Public Type UserFlags
    CantidadCorreos As Byte
    CasteandoPortal As Boolean
    TipoPocion As Byte
    DondeTiroMap As Integer
    DondeTiroX As Integer
    DondeTiroY As Integer
    
    CheckAmigos As Byte
    CantidadAmigos As Byte
    Dueleando As Boolean

    RecibioCorreo As Byte
    
    MuertesUsuario As Long
    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    Meditando As Boolean
    ModoCombate As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Resucitando As Boolean
    Envenenado As Byte
    Incinerado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean

    NoPuedeSerAtacado As Boolean
    Vuela As Byte
 
    
    
    Trabajando As Boolean
    Lingoteando As Byte
    
    Navegando As Byte
    Montando  As Byte
    SeguroResu As Boolean
   
    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    OwnedNpc As Integer ' Npc que le pertenece (no puede ser atacado)
    NpcInv As Integer
   
    Ban As Byte
    AdministrativeBan As Byte
   
    TargetUser As Integer ' Usuario señalado
   
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
   
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
   
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
   
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
   
    NPCAtacado As Integer
    Ignorado As Boolean
   
    StatsChanged As Byte
    Privilegios As PlayerType
   
    ValCoDe As Integer
   
    LastCrimMatado As String
    LastCiudMatado As String
   
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean

   
    '[el oso]
    MD5Reportado As String
    '[/el oso]
   
    '[Barrin 30-11-03]
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    '[/Barrin 30-11-03]
   
    '[CDT 17-02-04]
    UltimoMensaje As Byte
    '[/CDT]
   
    Silenciado As Byte
   
   
    CentinelaOK As Boolean 'Centinela
    ParalizedBy As String
    
    ParalizedByIndex As Integer
    ParalizedByNpcIndex As Integer
End Type

Public Const Max_Correos As Byte = 10

Public Type tCorreos

      CantidadMensajes As Byte
      
      Carta As String
      Emisor As String
      Leida As Byte
      
      ObjetoIndex As Integer
      ObjetoCantidad As Integer
      
End Type

Public Type UserCounters
    CreoTeleport As Boolean
    TimeTeleport As Integer
    PacketsTick As Byte

    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    Lava As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    Incinerado As Integer
    Invisibilidad As Integer
    TiempoOculto As Integer
    Mimetismo As Integer
    PiqueteC As Long
    ContadorPiquete As Long
    Pena As Long
    SendMapCounter As WorldPos
     Pasos As Integer
    '[Gonzalo]
    Saliendo As Boolean
    Salir As Integer
    '[/Gonzalo]
    
    TiempoDeMapeo As Byte
    
    'Barrin 3/10/03
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    'Barrin
    
    IntervaloRevive As Long
    'Mermas
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long
    TimerEstadoAtacable As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
    
    failedUsageAttempts As Long
    
End Type
 
' $ Nuevo sistema de facciones $
Public Type tFacciones
    Status As Integer
    CiudadanosMatados    As Integer
    RenegadosMatados     As Integer
    RepublicanosMatados  As Integer
    
    MilicianosMatados    As Integer
    ArmadaMatados        As Integer
    CaosMatados          As Integer
    
    Rango                As Integer
End Type

' $ Shermie80 $

Public Type tCrafting

    Cantidad As Long
    PorCiclo As Integer

End Type
 
Public Type Amigos
  Nombre As String
  index As Integer
End Type

Public Type TCasamiento
  Candidato As Integer
  Casado As Byte
  Pareja As String
End Type

'Tipo de los Usuarios
Public Type User
 
    Donador As TDonador
    
    Casamiento As TCasamiento
    
    Correos(1 To Max_Correos) As tCorreos

    Redundance As Byte

    
    Name            As String
    id              As Long
    Account         As String 'Shermie80; Sistema de cuentas
    showName        As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    'Amigos
    Amigos(1 To MAXAMIGOS) As Amigos
    QuienAmigo As String
    Char            As Char 'Define la apariencia
    OrigChar        As Char
    
    desc            As String ' Descripcion
    
    Clase           As eClass
    raza            As eRaza
    Genero          As eGenero
    Hogar           As eCiudad
        
    Invent          As Inventario
    
    Pos             As WorldPos
    
    ConnIDValida    As Boolean
    ConnID          As Long
    
    BancoInvent     As BancoInventario
    
    Counters        As UserCounters
    Construir       As tCrafting
    
    MascotasIndex(1 To MAXMASCOTAS)  As Integer
    MascotasType(1 To MAXMASCOTAS)   As Integer
    NroMascotas                      As Integer
    
    Stats          As UserStats
    flags          As UserFlags
    Faccion        As tFacciones
   elpedidor As Integer

         LogOnTime  As Date
        UpTime     As Long
 
    ip             As String
    
    GuildIndex     As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan    As Integer
       
    KeyCrypt       As Integer
    
    AreasInfo      As AreaInfo
    
    'Outgoing and incoming messages
    outgoingData   As clsByteQueue
    incomingData   As clsByteQueue
    'Outgoing and incoming messages
    
End Type

'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

' New type for holding the pathfinding info
Public Type NpcPathFindingInfo

    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type

' New type for holding the pathfinding info

Public Type tDrops

    ObjIndex As Integer
    Amount As Long

End Type

Public Const MAX_NPC_DROPS As Byte = 5

Public Type NPCStats
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
End Type
 
Public Type NpcCounters

    Paralisis As Integer
    TiempoExistencia As Long
    Ataque As Long
End Type
 
Public Type NPCFlags

    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Status As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
   
    ExpCount As Long
   
    OldMovement As TipoAI
    OldHostil As Byte
   
    AguaValida As Byte
    TierraInvalida As Byte
   
    Sound As Integer
    AttackedBy As String
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
   
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
   
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer

End Type
 
Public Type tCriaturasEntrenador
    npcindex As Integer
End Type
 
Public Type npc

    Name As String
    Char As Char 'Define como se vera
    desc As String
 
    NPCtype As eNPCType
    Numero As Integer
 
    InvReSpawn As Byte
 
    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer
 
    Veneno As Byte
   
    Pos As WorldPos 'Posicion
    oldPos As WorldPos
    Orig As WorldPos
    StartPos As WorldPos
 
    SkillDomar As Integer

    Movement As TipoAI
    Leveles As Integer
    Attackable As Byte
    Hostile As Byte
    
    PoderAtaque As Long
    PoderEvasion As Long
 
    GiveEXP As Long
    GiveGLD As Long
 
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
   
    Invent As Inventario
    CanAttack As Byte

    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
   
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
   
    ' New!! Needed for pathfindig
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
   
    Owner As Integer
    Drop(1 To MAX_NPC_DROPS) As tDrops
    Ciudad As Integer

End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock

    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    npcindex As Integer
    ObjInfo As Obj
    ObjEsFijo As Byte
    TileExit As WorldPos
    Trigger As eTrigger

End Type

'Info del mapa
Type MapInfo

    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Seguro As Byte
    Pk As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
   
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    RoboNpcsPermitido As Byte
    
    battle_mode As Byte
    
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE                       As Boolean
Public ULTIMAVERSION                      As String
Public BackUp                             As Boolean ' TODO: Se usa esta variable ?

Public ListaRazas(1 To NUMRAZAS)          As String
Public SkillsNames(1 To NUMSKILLS)        As String
Public ListaClases(1 To NUMCLASES)        As String
Public ListaAtributos(1 To NUMATRIBUTOS)  As String

Public recordusuarios                     As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath                            As String

''
'Ruta base para guardar los chars
Public CharPath                           As String

''
'Ruta base para los archivos de mapas
Public MapPath                            As String

''
'Ruta base para los DATs
Public DatPath                            As String
Public DocPath As String
Public DocConsultas As String
''
'Bordes del mapa
Public MinXBorder                         As Byte
Public MaxXBorder                         As Byte
Public MinYBorder                         As Byte
Public MaxYBorder                         As Byte

''
'Numero de usuarios actual
Public NumUsers                           As Integer
Public LastUser                           As Integer
Public LastChar                           As Integer
Public NumChars                           As Integer
Public LastNPC                            As Integer
Public NumNPCs                            As Integer
Public NumFX                              As Integer
Public NumMaps                            As Integer
Public NumObjDatas                        As Integer
Public NumeroHechizos                     As Integer
Public AllowMultiLogins                   As Byte
Public IdleLimit                          As Integer
Public MaxUsers                           As Integer
Public HideMe                             As Byte
Public LastBackup                         As String
Public Minutos                            As String
Public PuedeCrearPersonajes               As Integer
Public ServerSoloGMs                      As Integer

''
'Esta activada la verificacion MD5 ?
Public MD5ClientesActivado                As Byte

Public EnPausa                            As Boolean
Public EnTesting                          As Boolean

'*****************ARRAYS PUBLICOS*************************
Public UserList()                         As User 'USUARIOS
Public Npclist(1 To MAXNPCS)              As npc 'NPCS
Public MapData()                          As MapBlock
Public MapInfo()                          As MapInfo
Public Hechizos()                         As tHechizo
Public CharList(1 To MAXCHARS)            As Integer
Public ObjData()                          As ObjData
Public FX()                               As FXdata
Public SpawnList()                        As tCriaturasEntrenador
Public LevelSkill(1 To 50)                As LevelSkill
Public ForbidenNames()                    As String
Public Correos()                          As Integer

' $ Shermie80 $
Public ArmasHerrero()                     As Integer
Public CascosHerrero()                    As Integer
Public EscudosHerrero()                   As Integer
Public ArmadurasHerrero()                 As Integer

Public ObjCarpintero()                    As Integer
Public ObjDruida()                        As Integer
Public ObjSastre()                        As Integer
' $ Shermie80 $

Public MD5s()                             As String
Public BanIps                             As New Collection
Public PremiosInfo()                      As tPremios
Public ModClase(1 To NUMCLASES)           As ModClase
Public ModRaza(1 To NUMRAZAS)             As ModRaza
Public ModVida(1 To NUMCLASES)            As Double
Public DistribucionEnteraVida(1 To 5)     As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
 
'*********************************************************
 


'Alls variables Ciudad //Mermas
Public Ciudades() As WorldPos

'Ciudades as WorldPos
Public Type tCiudades

Nix As WorldPos
Illiandor As WorldPos
Ullathorpe As WorldPos
Rinkel As WorldPos
Banderbill As WorldPos
DungeonNewbie As WorldPos
Lindos As WorldPos
Arghal As WorldPos
Tiama As WorldPos
Orac As WorldPos
Suramei As WorldPos
Nueva As WorldPos
Prision As WorldPos
Libertad As WorldPos
Intermundia As WorldPos

End Type

Public tCiudades As tCiudades

'Fin ciudades



Public Ayuda           As New cCola
Public SonidosMapas    As New SoundMapInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nsize As Long, _
                                                 ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (ByRef destination As Any, _
                                      ByVal length As Long)

Public Enum e_ObjetosCriticos

    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467

End Enum

Public Enum eGMCommands

    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    GoNearby                '/IRCERCA
    comment                 '/REM
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
    BanChar                 '/BAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
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
    DARPUN                  '/DARPUN USUARIO@CANTIDAD
    ResponderGM
    Donador
    EventoOro
    EventoExperiencia
    CuentaRegresiva
End Enum

Public levelELU(1 To STAT_MAXELV) As Long
 
