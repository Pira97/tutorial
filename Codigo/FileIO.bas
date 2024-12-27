Attribute VB_Name = "ES"
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

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type: Private HistorialError As UltimoError
Private Type tMapHeader

    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long

End Type
 
Private Type tDatosBloqueados

    X As Integer
    Y As Integer

End Type
 
Private Type tDatosGrh

    X As Integer
    Y As Integer
    GrhIndex As Long

End Type
 
Private Type tDatosTrigger

    X As Integer
    Y As Integer
    Trigger As Integer

End Type
 
Private Type tDatosLuces

    X As Integer
    Y As Integer
    color As Long
    Rango As Byte

End Type
 
Private Type tDatosParticulas

    X As Integer
    Y As Integer
    Particula As Long

End Type
 
Private Type tDatosNPC

    X As Integer
    Y As Integer
    npcindex As Integer

End Type
 
Private Type tDatosObjs

    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer

End Type
 
Private Type tDatosTE

    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer

End Type
 
Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type
 
Private Type tMapDat
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    ambient As String * 16
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32

End Type

Private MapSize As tMapSize
Private MapDat As tMapDat
 
Function EsAdmin(ByVal Name As String) As Boolean

EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)


End Function

Function EsDios(ByVal Name As String) As Boolean
EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

End Function

Function EsConsejero(ByVal Name As String) As Boolean
   EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
  EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)
End Function
Public Function EsGmAccount(ByRef Name As String) As Boolean

    EsGmAccount = (val(GetVar(AccountPath & Name & ".cnt", Name, "CuentaGM")) = 1)
    
End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    'Returns true if char is administrative user.
    '***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)

    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)

    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)

    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function


Public Function TxtDimension(ByVal Name As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, cad As String, Tam As Long
    n = FreeFile(1)
    Open Name For Input As #n
    Tam = 0

    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    TxtDimension = Tam

End Function

Public Sub CargarForbidenWords()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo CargarForbidenWords_Err
     If frmMain.Visible Then frmMain.AgregarConsola "Cargando Nombres prohibidos (NombresInvalidos.txt)."
100    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))

       Dim n As Integer, i As Integer
       
102    n = FreeFile(1)
104    Open DatPath & "NombresInvalidos.txt" For Input As #n
    
106    For i = 1 To UBound(ForbidenNames)
108        Line Input #n, ForbidenNames(i)
110    Next i
    
112    Close n
    
        If frmMain.Visible Then frmMain.AgregarConsola "NombresInvalidos.txt han cargado con exito."
    Exit Sub

CargarForbidenWords_Err:
114     Call RegistrarError(Err.Number, Err.description, "ES.CargarForbidenWords", Erl)
116     Resume Next
        
End Sub

Public Sub CargarHechizos()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
    '
    'El que ose desafiar esta LEY, se las tendrá que ver
    'con migo. Para leer Hechizos.dat se deberá usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.AgregarConsola "Cargando Hechizos."
    
    Dim Hechizo As Integer
    Dim Leer    As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
   ' frmCargando.cargar.min = 0
    'frmCargando.cargar.max = NumeroHechizos
   ' frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .Loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            .Particle = val(Leer.GetValue("Hechizo" & Hechizo, "Particle"))
            .TimeParticula = val(Leer.GetValue("Hechizo" & Hechizo, "TimeParticula"))
            '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))

            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))

            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))

            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Incinera = val(Leer.GetValue("Hechizo" & Hechizo, "Incinera"))

            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .AfectaArea = val(Leer.GetValue("Hechizo" & Hechizo, "Area"))
            .AreaX = val(Leer.GetValue("Hechizo" & Hechizo, "AreaX"))
            .AreaY = val(Leer.GetValue("Hechizo" & Hechizo, "AreaY"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            
            'Nuevos Mermas
            .HechizoDeArea = val(Leer.GetValue("Hechizo" & Hechizo, "HechizoDeArea"))
            .AreaEfecto = val(Leer.GetValue("Hechizo" & Hechizo, "AreaEfecto"))
            .Afecta = val(Leer.GetValue("Hechizo" & Hechizo, "Afecta"))
            .AutoLanzar = val(Leer.GetValue("Hechizo" & Hechizo, "autolanzar"))
            
            
            Hechizos(Hechizo).ResucitaFamiliar = val(Leer.GetValue("Hechizo" & Hechizo, "ResucitaFamiliar"))
    
            Hechizos(Hechizo).extrahit = val(Leer.GetValue("Hechizo" & Hechizo, "extrahit"))
            Hechizos(Hechizo).Metamorfosis = val(Leer.GetValue("Hechizo" & Hechizo, "metamorfosis"))
            Hechizos(Hechizo).body = val(Leer.GetValue("Hechizo" & Hechizo, "body"))
            Hechizos(Hechizo).Desencantar = val(Leer.GetValue("Hechizo" & Hechizo, "desencantar"))
            .Sanacion = val(Leer.GetValue("Hechizo" & Hechizo, "Sanacion"))
            '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
            '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            
            .Anillo = val(Leer.GetValue("Hechizo" & Hechizo, "Anillo"))
            
           ' frmCargando.cargar.Value = frmCargando.cargar.Value + 1


            Hechizos(Hechizo).MaterializaX = val(Leer.GetValue("Hechizo" & Hechizo, "Materializax"))
            Hechizos(Hechizo).MaterializaObj = val(Leer.GetValue("Hechizo" & Hechizo, "MaterializaObj"))
            Hechizos(Hechizo).MaterializaCant = val(Leer.GetValue("Hechizo" & Hechizo, "MaterializaCant"))

        End With

    Next Hechizo
    
    Set Leer = Nothing
    If frmMain.Visible Then frmMain.AgregarConsola "Los hechizos se han cargado con exito."
    Exit Sub

ErrHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub
 
Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

On Error GoTo ErrorHandler

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Integer
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE

Dim i As Integer
Dim j As Integer
Dim tmpLng As Long

ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)

For j = MapSize.YMin To MapSize.YMax
    For i = MapSize.XMin To MapSize.XMax
        With MapData(i, j)
            If .Blocked Then
                MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                Blqs(MH.NumeroBloqueados).X = i
                Blqs(MH.NumeroBloqueados).Y = j
            End If
       
            L1(i, j) = .Graphic(1)
       
            If .Graphic(2) > 0 Then
                MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                ReDim Preserve L2(1 To MH.NumeroLayers(2))
                L2(MH.NumeroLayers(2)).X = i
                L2(MH.NumeroLayers(2)).Y = j
                L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2)
            End If
       
            If .Graphic(3) > 0 Then
                MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                ReDim Preserve L3(1 To MH.NumeroLayers(3))
                L3(MH.NumeroLayers(3)).X = i
                L3(MH.NumeroLayers(3)).Y = j
                L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3)
                End If
       
            If .Graphic(4) > 0 Then
                MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                ReDim Preserve L4(1 To MH.NumeroLayers(4))
                L4(MH.NumeroLayers(4)).X = i
                L4(MH.NumeroLayers(4)).Y = j
                L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4)
            End If
       
            If .Trigger > 0 Then
                MH.NumeroTriggers = MH.NumeroTriggers + 1
                ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                Triggers(MH.NumeroTriggers).X = i
                Triggers(MH.NumeroTriggers).Y = j
                Triggers(MH.NumeroTriggers).Trigger = .Trigger
            End If
       
             'tmpLng = Map_Light_Get(i, j)
       
            If tmpLng > 0 Then
                MH.NumeroLuces = MH.NumeroLuces + 1
                ReDim Preserve Luces(1 To MH.NumeroLuces)
                Luces(MH.NumeroLuces).X = i
                Luces(MH.NumeroLuces).Y = j
               ' Luces(MH.NumeroLuces).color = light_list(tmpLng).color
               ' Luces(MH.NumeroLuces).Rango = light_list(tmpLng).range
            End If
       
            If .ObjInfo.ObjIndex > 0 Then
                MH.NumeroOBJs = MH.NumeroOBJs + 1
                ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                Objetos(MH.NumeroOBJs).ObjIndex = .ObjInfo.ObjIndex
                Objetos(MH.NumeroOBJs).ObjAmmount = .ObjInfo.Amount
                Objetos(MH.NumeroOBJs).X = i
                Objetos(MH.NumeroOBJs).Y = j
            End If
       
            If .npcindex > 0 Then
                MH.NumeroNPCs = MH.NumeroNPCs + 1
                ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                NPCs(MH.NumeroNPCs).npcindex = .npcindex
                NPCs(MH.NumeroNPCs).X = i
                NPCs(MH.NumeroNPCs).Y = j
            End If
       
            If .TileExit.Map > 0 Then
                MH.NumeroTE = MH.NumeroTE + 1
                ReDim Preserve TEs(1 To MH.NumeroTE)
                TEs(MH.NumeroTE).DestM = .TileExit.Map
                TEs(MH.NumeroTE).DestX = .TileExit.X
                TEs(MH.NumeroTE).DestY = .TileExit.Y
                TEs(MH.NumeroTE).X = i
                TEs(MH.NumeroTE).Y = j
            End If
        End With
    Next i
Next j
     
fh = FreeFile
Open MAPFILE For Binary As fh

    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat
    Put #fh, , L1

    With MH
        If .NumeroBloqueados > 0 Then _
            Put #fh, , Blqs
        If .NumeroLayers(2) > 0 Then _
            Put #fh, , L2
        If .NumeroLayers(3) > 0 Then _
            Put #fh, , L3
        If .NumeroLayers(4) > 0 Then _
            Put #fh, , L4
        If .NumeroTriggers > 0 Then _
            Put #fh, , Triggers
        If .NumeroParticulas > 0 Then _
            Put #fh, , Particulas
        If .NumeroLuces > 0 Then _
            Put #fh, , Luces
        If .NumeroOBJs > 0 Then _
            Put #fh, , Objetos
        If .NumeroNPCs > 0 Then _
            Put #fh, , NPCs
        If .NumeroTE > 0 Then _
            Put #fh, , TEs
    End With

Close fh

'Save_Map_Data = True

Exit Sub

ErrorHandler:
    If fh <> 0 Then Close fh
End Sub

Sub LoadCascosHerreria()
    '***************************************************
    'Author: Shermie80
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando cascos crafteables por Herreria."
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "CascosHerrero.dat", "INIT", "NumCascos"))
    
    ReDim Preserve CascosHerrero(1 To n) As Integer
    
    For lc = 1 To n
        CascosHerrero(lc) = val(GetVar(DatPath & "CascosHerrero.dat", "Casco" & lc, "Index"))
    Next lc
    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo las cascos crafteables por Herreria. Operacion Realizada con exito."
End Sub

Sub LoadEscudosHerreria()
    '***************************************************
    'Author: Shermie80
    'Last Modification: -
    '
    '***************************************************
     If frmMain.Visible Then frmMain.AgregarConsola "Cargando escudos crafteables por Herreria."
         
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "EscudosHerrero.dat", "INIT", "NumEscudos"))
    
    ReDim Preserve EscudosHerrero(1 To n) As Integer
    
    For lc = 1 To n
        EscudosHerrero(lc) = val(GetVar(DatPath & "EscudosHerrero.dat", "Escudo" & lc, "Index"))
    Next lc
   If frmMain.Visible Then frmMain.AgregarConsola "Se cargo los escudos crafteables por Herreria. Operacion Realizada con exito."
    
    
End Sub

Sub LoadArmasHerreria()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando armas crafteables por Herreria."
    
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc
    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo las armas crafteables por Herreria. Operacion Realizada con exito."
End Sub

Sub LoadArmadurasHerreria()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando armaduras crafteables por Herreria."

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc
    
    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo las armaduras crafteables por Herreria. Operacion Realizada con exito."
    
End Sub

Sub LoadBalance()
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/04/2010
    '15/04/2010: ZaMa - Agrego recompensas faccionarias.
    '***************************************************
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando el archivo Balance.dat"
    
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES

        With ModClase(i)
            .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
            .DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
            .DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
            .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
.AtaqueArpon = val(GetVar(DatPath & "Balance.dat", "MODAtaqueArpon", ListaClases(i)))
.DañoArpon = val(GetVar(DatPath & "Balance.dat", "MODDañoArpon", ListaClases(i)))
 
        End With

    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS

        With ModRaza(i)
            .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))

        End With

    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i

    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))
    
        
    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo con exito el archivo Balance.dat"

End Sub

Sub LoadObjDruida()
If frmMain.Visible Then frmMain.AgregarConsola "Cargando druida.dat."
Dim n As Integer, lc As Integer

    n = val(GetVar(DatPath & "objdruida.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjDruida(1 To n) As Integer
    
    For lc = 1 To n
        ObjDruida(lc) = val(GetVar(DatPath & "objdruida.dat", "Obj" & lc, "Index"))
    Next lc
 If frmMain.Visible Then frmMain.AgregarConsola "Se cargo druida.dat. Operacion Realizada con exito."
End Sub

Sub LoadObjSastre()
If frmMain.Visible Then frmMain.AgregarConsola "Cargando sastre.dat."
Dim n As Integer, lc As Integer

    n = val(GetVar(DatPath & "objsastre.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjSastre(1 To n) As Integer
    
    For lc = 1 To n
        ObjSastre(lc) = val(GetVar(DatPath & "objsastre.dat", "Obj" & lc, "Index"))
    Next lc
 If frmMain.Visible Then frmMain.AgregarConsola "Se cargo sastre.dat. Operacion Realizada con exito."
End Sub

Sub LoadObjCarpintero()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando los objetos crafteables via Carpinteria"
    
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To n) As Integer
    
    For lc = 1 To n
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc
    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo con exito los objetos crafteables via Carpinteria."

End Sub
Sub LoadOBJData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
 
    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
    '
    'El que ose desafiar esta LEY, se las tendrá que ver
    'con migo. Para leer desde el OBJ.DAT se deberá usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################
 
    'Call LogTarea("Sub LoadOBJData")
 
    On Error GoTo ErrHandler
 
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando base de datos de los objetos."
   
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    
    Dim Leer   As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Obj.dat")
   
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
 
   
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
 For Object = 1 To NumObjDatas
        Debug.Print "Cargando Objs... " & Round(Object / NumObjDatas * 100, 2) & "%"
        frmCargando.Label3.Caption = "Cargando dats... " & Object & "/" & NumObjDatas
        
        ObjData(Object).Numero = Object
        
        ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
        
        ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

        If ObjData(Object).GrhIndex = 0 Then
            ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
        End If
   
        ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
   
        ObjData(Object).SubTipo = val(Leer.GetValue("OBJ" & Object, "SubTipo"))
        ObjData(Object).LanzaHechizo = val(Leer.GetValue("OBJ" & Object, "LanzaHechizo"))
        
        ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
        ObjData(Object).Destruir = val(Leer.GetValue("OBJ" & Object, "Destruir"))
        'ObjTypes efectos magicos
        ObjData(Object).EfectoMagico = val(Leer.GetValue("OBJ" & Object, "EfectoMagico"))
         
        
        
         
 
        ObjData(Object).Shop = val(Leer.GetValue("OBJ" & Object, "Shop"))
        
        ObjData(Object).QueSkill = val(Leer.GetValue("OBJ" & Object, "QueSkill"))
        
        ObjData(Object).QueAtributo = val(Leer.GetValue("OBJ" & Object, "QueAtributo"))
        
        ObjData(Object).CuantoAumento = val(Leer.GetValue("OBJ" & Object, "CuantoAumento"))
        
        With ObjData(Object)

            Select Case ObjData(Object).OBJType
                
                Case eOBJType.otArmadura
           
                    If .SubTipo = 1 Then
                        ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    ElseIf .SubTipo = 2 Then
                        ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                        ObjData(Object).DosManos = val(Leer.GetValue("OBJ" & Object, "DosManos"))
                    End If
    
                    ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    'Items Faccionarios
                    ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    ObjData(Object).Milicia = val(Leer.GetValue("OBJ" & Object, "Milicia"))

                       Case eOBJType.otNudillos
                    ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    
                Case eOBJType.otWeapon
                    ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                    ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    ObjData(Object).Milicia = val(Leer.GetValue("OBJ" & Object, "Milicia"))
                    ObjData(Object).DosManos = val(Leer.GetValue("OBJ" & Object, "DosManos"))
                    ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    ObjData(Object).SndEspecial = val(Leer.GetValue("OBJ" & Object, "SndEspecial"))
                                        
                Case eOBJType.otInstrumentos
                    ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
 
                    ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    ObjData(Object).Milicia = val(Leer.GetValue("OBJ" & Object, "Milicia"))
               
                Case eOBJType.otMinerales
                    ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    
                Case eOBJType.otPozos
                    ObjData(Object).SubTipo = val(Leer.GetValue("OBJ" & Object, "SubTipo"))
           
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
           
                Case otPociones
                    ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
           
                Case eOBJType.otBarcos
                    ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
           
                Case eOBJType.otMonturas
                    .Velocidades = val(Leer.GetValue("OBJ" & Object, "Velocidad"))
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
           
                Case eOBJType.otFlechas
                    ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "Incinera")) 'Usamos el slot snd3, está al pedo
                    ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "FX1"))
                    
                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                   ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                Case eOBJType.otPasajes
                    ObjData(Object).DesdeMap = val(Leer.GetValue("OBJ" & Object, "Desde"))
                    ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
                    ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
                    ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))

            End Select

        End With
 
        ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
        ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
   
        ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
   
        ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
   
        ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
        ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
   
        ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
        ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
   
        ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
        ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
        ObjData(Object).levelItem = val(Leer.GetValue("OBJ" & Object, "MinELV"))
    
        ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
        ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
        
        ObjData(Object).MinELV = val(Leer.GetValue("OBJ" & Object, "MinELV"))
        ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
        
        ObjData(Object).puntos = val(Leer.GetValue("OBJ" & Object, "Puntos"))
        
        ObjData(Object).Aura = val(Leer.GetValue("OBJ" & Object, "Aura"))
        
        ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
   
        ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
         
         
        ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
            ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

        End If
   
        'Puertas y llaves
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
   
        ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
        ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
   
        ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))

        ObjData(Object).StaRequerido = val(Leer.GetValue("OBJ" & Object, "StaRequerido"))
        
        If ObjData(Object).StaRequerido < 1 Then ObjData(Object).StaRequerido = 10
        
        'Conversiones de razas, clases de la 1.4.5 a nuestros indices /Mermas
        Dim i As Integer
        Dim strString As String, ClasesProhibidas() As String
        
        If Len(Leer.GetValue("OBJ" & Object, "ClasesProhibidas")) > 0 Then 'Si hay ClaseProhibida
            
            strString = Leer.GetValue("OBJ" & Object, "ClasesProhibidas")
            
            ClasesProhibidas() = Split(strString, ",") 'Dividimos las palabras
            
            For i = 0 To UBound(ClasesProhibidas)
                ObjData(Object).ClaseProhibida(i + 1) = ClasesProhibidas(i)
            Next i
            
        End If
        
        i = 0
        strString = vbNullString
        Dim RazasProhibidas() As String
        
        If Len(Leer.GetValue("OBJ" & Object, "RazasProhibidas")) > 0 Then 'Si hay razas
            
            strString = Leer.GetValue("OBJ" & Object, "RazasProhibidas")
            
            RazasProhibidas() = Split(strString, ",") 'Dividimos las palabras
            
            For i = 0 To UBound(RazasProhibidas)
                ObjData(Object).RazaProhibida(i + 1) = QueRazaEs(RazasProhibidas(i))
            Next i
            
        End If
        '//Fin conversion
            
        'Load Skins
        
        i = 0
        strString = vbNullString
        Dim TieneSkin() As String
        
        If Len(Leer.GetValue("OBJ" & Object, "TieneSkin")) > 0 Then 'Si hay Skins
            
            strString = Leer.GetValue("OBJ" & Object, "TieneSkin")
            
            TieneSkin() = Split(strString, ",") 'Dividimos las palabras
            ObjData(Object).CantidadSkin = UBound(TieneSkin)
            
            ReDim ObjData(Object).TieneSkin(0 To ObjData(Object).CantidadSkin)
            
            For i = 0 To UBound(TieneSkin)
                ObjData(Object).TieneSkin(i) = TieneSkin(i)
            Next i
            
        End If
        
        
        ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
   
        ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
        
        If ObjData(Object).SkCarpinteria > 0 Then
            ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))

        End If
        
        ObjData(Object).SkPociones = val(Leer.GetValue("OBJ" & Object, "SkPociones"))
        
        ObjData(Object).SkSastreria = val(Leer.GetValue("OBJ" & Object, "SkSastreria"))
    
    If ObjData(Object).SkSastreria > 0 Then
        ObjData(Object).PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
        ObjData(Object).PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPolar"))
        ObjData(Object).PielOsoPolar = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
    End If
        
        If ObjData(Object).SkPociones > 0 Then
         ObjData(Object).Raies = val(Leer.GetValue("OBJ" & Object, "Raices"))
        End If
    Next Object
    DoEvents
    Set Leer = Nothing
     If frmMain.Visible Then frmMain.AgregarConsola "Se cargo base de datos de los objetos. Operacion Realizada con exito."
    
    Exit Sub
 
ErrHandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description
    
End Sub

Function GetVar(ByVal File As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.AgregarConsola "Cargando backup."
    
    Dim Map       As Integer
    Dim TempInt   As Integer
    Dim tFileName As String
    Dim npcfile   As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
        
    'frmCargando.cargar.min = 0
   ' frmCargando.cargar.max = NumMaps
   ' frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
        
    For Map = 1 To NumMaps

        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
                
            If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                tFileName = App.Path & MapPath & "Mapa" & Map

            End If

        Else
            tFileName = App.Path & MapPath & "Mapa" & Map

        End If
            
        Call CargarMapa(Map, tFileName)
            
       ' frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.AgregarConsola "Cargando mapas..."
   
    Dim Map       As Integer
    Dim TempInt   As Integer
    Dim tFileName As String
    Dim npcfile   As String
    
    On Error GoTo man
        
    Call InitAreas
        
    'frmCargando.cargar.min = 0
   ' frmCargando.cargar.max = NumMaps
   ' frmCargando.cargar.Value = 0
        
 
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
          
    For Map = 1 To NumMaps
        Debug.Print "Cargando Mapas... " & Round(Map / NumMaps * 100, 2) & "%"
        frmCargando.Label3.Caption = "Cargando mapas... " & Map & "/" & NumMaps
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
            
      '  frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
        If frmMain.Visible Then frmMain.AgregarConsola "Se cargaron todos los mapas. Operacion Realizada con exito."

    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByVal MAPFl As String)

    On Error GoTo errh
 
    Dim fh           As Integer
    Dim MH           As tMapHeader
    Dim Blqs()       As tDatosBloqueados
    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh
    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE
    Dim MapSize      As tMapSize
    Dim MapDat       As tMapDat
 
    Dim i            As Long
    Dim j            As Long
 
100    If Not FileExist(MAPFl & ".csm", vbNormal) Then
102        MsgBox "El arhivo " & App.Path & "\Maps\Mapa" & Map & ".csm" & " no existe."
104        Exit Sub
106    End If
   
108    fh = FreeFile
    
110    Dim fTxt As Integer
112    fTxt = FreeFile
    
114    Open MAPFl & ".csm" For Binary Access Read As fh
116    Get #fh, , MH
118    Get #fh, , MapSize
120    Get #fh, , MapDat

122    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
       
124    Get #fh, , L1
       
126    With MH

128        If .NumeroBloqueados > 0 Then
130            ReDim Blqs(1 To .NumeroBloqueados)
132            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Map, Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i
            
134        End If
           
136        If .NumeroLayers(2) > 0 Then
138            ReDim L2(1 To .NumeroLayers(2))
140            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                MapData(Map, L2(i).X, L2(i).Y).Graphic(2) = L2(i).GrhIndex
            Next i
            
142        End If
           
144        If .NumeroLayers(3) > 0 Then
146            ReDim L3(1 To .NumeroLayers(3))
148            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                MapData(Map, L3(i).X, L3(i).Y).Graphic(3) = L3(i).GrhIndex
            Next i
            
150        End If
           
152        If .NumeroLayers(4) > 0 Then
154            ReDim L4(1 To .NumeroLayers(4))
156            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                MapData(Map, L4(i).X, L4(i).Y).Graphic(4) = L4(i).GrhIndex
            Next i
            
158        End If
           
160        If .NumeroTriggers > 0 Then
162            ReDim Triggers(1 To .NumeroTriggers)
164            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                MapData(Map, Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i
            
166        End If
           
168        If .NumeroParticulas > 0 Then
170            ReDim Particulas(1 To .NumeroParticulas)
172            Get #fh, , Particulas
174        End If
           
176        If .NumeroLuces > 0 Then
178            ReDim Luces(1 To .NumeroLuces)
180            Get #fh, , Luces
182        End If
           
184        If .NumeroOBJs > 0 Then
186            ReDim Objetos(1 To .NumeroOBJs)
188            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
                MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount
 
                 If ObjData(Objetos(i).ObjIndex).OBJType <> eOBJType.otPuertas Then
                    MapData(Map, Objetos(i).X, Objetos(i).Y).ObjEsFijo = 1
                End If
                
            Next i
            
190        End If
               
192        If .NumeroNPCs > 0 Then
194            ReDim NPCs(1 To .NumeroNPCs)
196            Get #fh, , NPCs

            For i = 1 To .NumeroNPCs
                MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex = NPCs(i).npcindex
                If MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex > 0 Then
                    Dim npcfile As String
                    
                    npcfile = DatPath & "NPCs.dat"
    
                    If val(GetVar(npcfile, "NPC" & MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex, "PosOrig")) = 1 Then
                        MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex)
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).Orig.Map = Map
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).Orig.X = NPCs(i).X
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).Orig.Y = NPCs(i).Y
                    Else
                        MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex)
                    End If
                    
                    If Not MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex = 0 Then
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).Pos.Map = Map
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).Pos.X = NPCs(i).X
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).Pos.Y = NPCs(i).Y
                        
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).StartPos.Map = Map
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).StartPos.X = NPCs(i).X
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex).StartPos.Y = NPCs(i).Y
                        
                        Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).X, NPCs(i).Y).npcindex, Map, NPCs(i).X, NPCs(i).Y)
                    End If
                End If
            Next i
            
198        End If
               
200        If .NumeroTE > 0 Then
202            ReDim TEs(1 To .NumeroTE)
204            Get #fh, , TEs

            For i = 1 To .NumeroTE
                MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
            Next i
            
206        End If
           
208    End With
   
210    Close fh
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                MapData(Map, i, j).Graphic(1) = L1(i, j)
            End If
        Next i
    Next j


   
    MapDat.map_name = Trim$(MapDat.map_name)
    
    MapInfo(Map).Name = MapDat.map_name
    MapInfo(Map).Music = MapDat.music_number
    MapInfo(Map).Seguro = MapDat.extra1
    
    
    'If Not (Left$(MapDat.zone, 6) = "CIUDAD") Then
    '    MapInfo(Map).Pk = True
    'Else
    '    MapInfo(Map).Pk = False
    'End If
    
    MapInfo(Map).battle_mode = MapDat.battle_mode
    MapInfo(Map).Terreno = MapDat.terrain
    MapInfo(Map).Zona = Trim$(MapDat.zone)
    MapInfo(Map).Restringir = MapDat.restrict_mode
    
    MapInfo(Map).BackUp = MapDat.backup_mode
    
    'Arreglos manual de mapas :p, // Mermas, aca cambiamos datos al cargar que en la 1.4.5 no se cargaban, hardcodeado pero bueno
    If (Left$(MapDat.zone, 6) = "CIUDAD") Then MapInfo(Map).Pk = False
    
    If MapInfo(Map).battle_mode = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    If Map = 457 Then MapInfo(Map).Pk = False
    Exit Sub
 
errh:
    Call RegistrarError(Err.Number, Err.description, "ES.CargarMapa", Erl)
    Call LogError("Error cargando mapa: " & Map & " ." & Err.description)

End Sub

Sub LoadSini()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo LoadSini_Err
    
100    Dim Temporal As Long
    
102    Dim Lector As clsIniManager
104    Set Lector = New clsIniManager
    
106    If frmMain.Visible Then
108        frmMain.AgregarConsola "Cargando info de inicio del server."
    End If
    
110    Call Lector.Initialize(IniPath & "Server.ini")
 
112    BootDelBackUp = CBool(val(Lector.GetValue("INIT", "IniciarDesdeBackUp")))
    
    'Misc
114    Puerto = val(Lector.GetValue("INIT", "StartPort"))
116    LastSockListen = val(Lector.GetValue("INIT", "LastSockListen"))
118    HideMe = CBool(Lector.GetValue("INIT", "Hide"))
120    AllowMultiLogins = CBool(val(Lector.GetValue("INIT", "AllowMultiLogins")))
122    IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
    
    'Lee la version correcta del cliente
124    ULTIMAVERSION = Lector.GetValue("INIT", "Version")
    
    
    'Variables de Experiencia, oro, porcentaje de subida de Skills
126    Expc = val(Lector.GetValue("INIT", "Exp"))
128    Oroc = val(Lector.GetValue("INIT", "Oro"))
130    PorcentajeSkill = val(Lector.GetValue("INIT", "PorcentajeSkills"))

132    PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
134    ServerSoloGMs = val(Lector.GetValue("INIT", "ServerSoloGMs"))

    'Intervalos
136    SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
138    StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
140    SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
142    StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
148    IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
150    IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
152    IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
154    IntervaloIncinerado = val(Lector.GetValue("INTERVALOS", "IntervaloIncinerado"))
156    IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
158    IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
160    IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
162    IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
164    IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
166    IntervaloParaConexion = val(Lector.GetValue("INTERVALOS", "IntervaloParaConexion"))
168    IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
170    IntervaloUserPuedeTrabajar = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajo"))
172    IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
    
    'TODO : Agregar estos intervalos al form!!!
174    IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
176    IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
178    IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))

177    IntervaloMensajeAutomatico1 = val(Lector.GetValue("INTERVALOS", "IntervaloMensajeAutomatico1"))
179    IntervaloMensajeAutomatico2 = val(Lector.GetValue("INTERVALOS", "IntervaloMensajeAutomatico2"))

    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
180    IntervaloPuedeSerAtacado = val(Lector.GetValue("TIMERS", "IntervaloPuedeSerAtacado"))
182    IntervaloAtacable = val(Lector.GetValue("TIMERS", "IntervaloAtacable"))
184    IntervaloOwnedNpc = val(Lector.GetValue("TIMERS", "IntervaloOwnedNpc"))
    
    
190    MinutosGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
192    IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
194    IntervaloUserPuedeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsar"))
196    IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
    
198    IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))
    
200    recordusuarios = val(Lector.GetValue("INIT", "Record"))
201    NumMaps = val(Lector.GetValue("INIT", "NumMaps"))
202    MapPath = Lector.GetValue("INIT", "MapPath")
                 
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
 
       'Max users
203    Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

204    If MaxUsers = 0 Then
206        MaxUsers = Temporal
208        ReDim UserList(1 To MaxUsers) As User
210    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agregó en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    
    
    ' $ Shermie80 / Cargamos los items de Shop /Mermas 2021, lo optimizamos xxd
214    CantPremios = val(Lector.GetValue("SHOP", "CantObjetos"))
    
216    If CantPremios > 0 Then
218        ReDim PremiosInfo(1 To CantPremios) As tPremios
220         Dim i As Byte
222         For i = 1 To CantPremios
224            PremiosInfo(i).ObjIndex = val(Lector.GetValue("SHOP", "ObjIndex" & i))
226            PremiosInfo(i).puntos = val(Lector.GetValue("SHOP", "Creditos" & i))
228         Next i
230    End If
    ''//
    
232    Set Lector = Nothing
    
234    Call MD5sCarga

235    Call LoadMotd
236    Call LoadUpdate 'Load actualizaciones
    
    ' Admins
238    Call loadAdministrativeUsers


    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo la info de inicio del server (Sinfo.ini)"
    
    
    Exit Sub

LoadSini_Err:
240     Call RegistrarError(Err.Number, Err.description, "ES.LoadSini", Erl)
242     Resume Next
        
End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Escribe VAR en un archivo
    '***************************************************

    writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal LogOut As Boolean = False)
    
    On Error GoTo ErrHandler
    
    Dim UserFile    As String
    Dim OldUserHead As Long

    With UserList(UserIndex)
 
        UserFile = CharPath & UCase$(.Name) & ".chr"
    
        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
        If .Clase = 0 Or .Stats.ELV = 0 Then
            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
            Exit Sub
        End If

        If FileExist(UserFile, vbNormal) Then
            If .flags.Muerto = 1 Then
                OldUserHead = .Char.Head
                .Char.Head = GetVar(UserFile, "INIT", "Head")
            End If
        End If
    
        Dim loopc As Integer
        
        If FileExist(UserFile, vbNormal) Then Kill UserFile

        Dim File As String: File = UserFile
        Dim n As Integer: n = FreeFile
        
        Open File For Output Access Write As n
        
        
        '[FLAGS]
        With .flags
122         Print #n, "[FLAGS]"
            Print #n, "Muerto=" & CStr(.Muerto)
            Print #n, "Escondido=" & CStr(.Escondido)
            Print #n, "Hambre=" & CStr(.Hambre)
            Print #n, "Sed=" & CStr(.Sed)
            Print #n, "Desnudo=" & CStr(.Desnudo)
            Print #n, "Ban=" & CStr(.Ban)
            Print #n, "Navegando=" & CStr(.Navegando)
            Print #n, "Montando=" & CStr(.Montando)
            Print #n, "Envenenado=" & CStr(.Envenenado)
            Print #n, "Paralizado=" & CStr(.Paralizado)
            Print #n, "Incinerado=" & CStr(.Incinerado)
            Print #n, "Correos=" & CStr(.CantidadCorreos)
            Print #n, "Murio=" & CStr(.MuertesUsuario)
            Print #n, "RecibioCorreo=" & CStr(.RecibioCorreo)
            Print #n, "CantidadAmigos=" & CStr(.CantidadAmigos)
        
        End With

        '[COUNTERS]
        With .Counters
123         Print #n, "[COUNTERS]"
            Print #n, "Pena=" & CStr(.Pena)
        
        End With


        '[FACCIONES]
        With .Faccion
125         Print #n, "[FACCIONES]"
            Print #n, "Rango=" & CStr(.Rango)
            Print #n, "Status=" & CStr(.Status)
            Print #n, "CiudMatados=" & CStr(.CiudadanosMatados)
            Print #n, "ReneMatados=" & CStr(.RenegadosMatados)
            Print #n, "RepuMatados=" & CStr(.RepublicanosMatados)
            Print #n, "CaosMatados=" & CStr(.CaosMatados)
            Print #n, "ArmiMatados=" & CStr(.ArmadaMatados)
            Print #n, "MiliMatados=" & CStr(.MilicianosMatados)
        
        End With

        '[DONADOR]
        With .Donador
126         Print #n, "[DONADOR]"
            Print #n, "Donador=" & CStr(.activo)
            Print #n, "Puntos=" & CStr(.CreditoDonador)
        
        End With

        '[ATRIBUTOS]
        Print #n, "[ATRIBUTOS]"
        
        '¿Fueron modificados los atributos del usuario?
        If Not .flags.TomoPocion Then

            For loopc = 1 To UBound(.Stats.UserAtributos)
                Print #n, "AT" & loopc & "=" & CStr(.Stats.UserAtributos(loopc))
            Next

        Else

            For loopc = 1 To UBound(.Stats.UserAtributos)
                Print #n, "AT" & loopc & "=" & CStr(.Stats.UserAtributosBackUP(loopc))
            Next

        End If

        '[SKILLS]
        Print #n, "[SKILLS]"
     
        For loopc = 1 To UBound(.Stats.UserSkills)
            Print #n, "SK" & loopc & "=" & CStr(.Stats.UserSkills(loopc))
            Print #n, "ELUSK" & loopc & "=" & CStr(.Stats.EluSkills(loopc))
            Print #n, "EXPSK" & loopc & "=" & CStr(.Stats.ExpSkills(loopc))
        Next

        '[CASAMIENTO]
        With .Casamiento
127         Print #n, "[CASAMIENTO]"
            Print #n, "Pareja=" & CStr(.Pareja)
            Print #n, "Casado=" & CStr(.Casado)
        
        End With


        '[INIT]
128     Print #n, "[INIT]"
        Print #n, "Genero=" & CStr(.Genero)
        Print #n, "Raza=" & CStr(.raza)
        Print #n, "Hogar=" & CStr(.Hogar)
        Print #n, "Clase=" & CStr(.Clase)
        Print #n, "Desc=" & CStr(.desc)
        Print #n, "Heading=" & CStr(.Char.heading)
        
        If .Char.Head = 0 Then
            Print #n, "Head=" & CStr(.OrigChar.Head)
        Else
            Print #n, "Head=" & CStr(.Char.Head)
        End If
        
        Print #n, "Body=" & CStr(.Char.body)
        Print #n, "Arma=" & CStr(.Char.WeaponAnim)
        Print #n, "Escudo=" & CStr(.Char.ShieldAnim)
        Print #n, "Casco=" & CStr(.Char.CascoAnim)



        Dim TempDate As Date
        TempDate = Now - .LogOnTime
        .LogOnTime = Now
        .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        .UpTime = .UpTime
        
        Print #n, "UpTime=" & .UpTime

        If LogOut Then
            Print #n, "Logged=0"
            Call WriteVar(AccountPath & UCase$(UserList(UserIndex).Account) & ".cnt", UCase$(UserList(UserIndex).Account), "Conectada", "0")
        Else
            Print #n, "Logged=1"
        End If
        
        Print #n, "Position=" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y
        

        '[STATS]
        With .Stats
        
129         Print #n, "[STATS]"
            Print #n, "GLD=" & CStr(.GLD)
            Print #n, "BANCO=" & CStr(.Banco)
            Print #n, "MaxHP=" & CStr(.MaxHP)
            Print #n, "MinHP=" & CStr(.MinHP)
            Print #n, "MaxSTA=" & CStr(.MaxSta)
            Print #n, "MinSTA=" & CStr(.MinSta)
            Print #n, "MaxMAN=" & CStr(.MaxMAN)
            Print #n, "MinMAN=" & CStr(.MinMAN)
            Print #n, "MaxHIT=" & CStr(.MaxHIT)
            Print #n, "MinHIT=" & CStr(.MinHIT)
            Print #n, "MaxAGU=" & CStr(.MaxAGU)
            Print #n, "MinAGU=" & CStr(.MinAGU)
            Print #n, "MaxHAM=" & CStr(.MaxHam)
            Print #n, "MinHAM=" & CStr(.MinHam)
            Print #n, "SkillPtsLibres=" & CStr(.SkillPts)
            Print #n, "EXP=" & CStr(.Exp)
            Print #n, "ELV=" & CStr(.ELV)
            Print #n, "ELU=" & CStr(.ELU)
            
                        
            '[MUERTES]
            Print #n, "[MUERTES]"
            Print #n, "UserMuertes=" & CStr(.UsuariosMatados)
            Print #n, "NpcsMuertes=" & CStr(.NPCsMuertos)
            
        End With
        
        
        '[BancoInventory]
        Print #n, "[BancoInventory]"
        Print #n, "CantidadItems=" & CStr(.BancoInvent.NroItems)

        Dim loopd As Long
        
        For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
            Print #n, "Obj" & loopd & "=" & .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount
        Next loopd
                

        '[Inventory]
        With .Invent
            Print #n, "[Inventory]"
            Print #n, "CantidadItems=" & CStr(.NroItems)

            For loopc = 1 To MAX_INVENTORY_SLOTS
                Print #n, "Obj" & loopc & "=" & .Object(loopc).ObjIndex & "-" & .Object(loopc).Amount & "-" & .Object(loopc).Equipped
            Next
            
            Print #n, "WeaponEqpSlot=" & CStr(.WeaponEqpSlot)
            Print #n, "NudiEqpSlot=" & CStr(.NudiEqpSlot)
            Print #n, "ArmourEqpSlot=" & CStr(.ArmourEqpSlot)
            Print #n, "CascoEqpSlot=" & CStr(.CascoEqpSlot)
            Print #n, "EscudoEqpSlot=" & CStr(.EscudoEqpSlot)
            Print #n, "BarcoSlot=" & CStr(.BarcoSlot)
            Print #n, "MonturaSlot=" & CStr(.MonturaSlot)
            Print #n, "MunicionSlot=" & CStr(.MunicionEqpSlot)
            Print #n, "AnilloSlot=" & CStr(.AnilloEqpSlot)
            Print #n, "MagicSlot=" & CStr(.MagicSlot)
            
        End With


        '[HECHIZOS]
        Print #n, "[HECHIZOS]"

        Dim cad As String
        
        For loopc = 1 To MAXUSERHECHIZOS
            cad = .Stats.UserHechizos(loopc)
            Print #n, "H" & loopc & "=" & cad
        Next
            

        '[AMIGOS]
        Print #n, "[AMIGOS]"
        
        For loopc = 1 To MAXAMIGOS
            Print #n, "NOMBRE" & loopc & "=" & CStr(.Amigos(loopc).Nombre)
        Next
        
                
        '[AMIGOS]
        Print #n, "[CORREO]"
        
        For loopc = 1 To Max_Correos
        
            Print #n, "Carta" & loopc & "=" & CStr(.Correos(loopc).Carta)
            Print #n, "Emisor" & loopc & "=" & CStr(.Correos(loopc).Emisor)
            Print #n, "Leida" & loopc & "=" & CStr(.Correos(loopc).Leida)
            Print #n, "Objeto" & loopc & "=" & CStr(.Correos(loopc).ObjetoIndex & "-" & .Correos(loopc).ObjetoCantidad)
            
        Next
        
        Close #n
        
        'Devuelve el head de muerto
        If .flags.Muerto = 1 Then
            .Char.Head = iCabezaMuerto
        End If
        
    End With
    
    Exit Sub

ErrHandler:
        Close #n
        Call RegistrarError(Err.Number, Err.description, "ES.SaveUser", Erl)
        Resume Next
End Sub
Sub BackUPnPc(npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NpcNumero As Integer
    Dim npcfile   As String
    Dim loopc     As Integer
    
    NpcNumero = Npclist(npcindex).Numero
    
    'If NpcNumero > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
   ' npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(npcindex)
        'General
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", .Name)
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", .desc)
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(.Char.Head))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(.Char.body))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(.Char.heading))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(.Movement))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Level", val(.Leveles))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(.Attackable))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(.Comercia))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(.TipoItems))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(.GiveEXP))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(.GiveGLD))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(.InvReSpawn))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(.NPCtype))
        
        'Stats
        Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(.Stats.def))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(.Stats.MaxHIT))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(.Stats.MaxHP))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(.Stats.MinHIT))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(.Stats.MinHP))
        
        'Flags
        Call WriteVar(npcfile, "NPC" & NpcNumero, "RespawnTime", val(.flags.Respawn))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(.flags.BackUp))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(.flags.Domable))
        
        'Inventario
        Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(.Invent.NroItems))
 
        If .Invent.NroItems > 0 Then

            For loopc = 1 To MAX_INVENTORY_SLOTS
                Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & loopc, .Invent.Object(loopc).ObjIndex & "-" & _
                        .Invent.Object(loopc).Amount)
            Next loopc

        End If

    End With

End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando apuestas.dat"

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))
    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo el archivo apuestas.dat"

End Sub

Function criminal(ByVal UserIndex As Integer) As Byte
If esRene(UserIndex) = 1 Then
    criminal = 1
ElseIf esCiuda(UserIndex) Or esArmada(UserIndex) = 1 Then
    criminal = 2
ElseIf esCaos(UserIndex) = 1 Then
    criminal = 3
ElseIf esMili(UserIndex) Or esRepu(UserIndex) = 1 Then
    criminal = 4
Else
    criminal = 5
End If
End Function
Public Sub EfectoIncinerado(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
Dim n As Integer
Dim reproducesonido As Byte
 With UserList(UserIndex)
        If .Counters.Incinerado < 50 Then
            .Counters.Incinerado = .Counters.Incinerado + DeltaTick
            If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub
                 Call WriteLocaleMsg(UserIndex, 48)
                 If reproducesonido = 0 Then
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(78, .Pos.X, .Pos.Y))
                 reproducesonido = 1
                 End If
                n = RandomNumber(45, 65)
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - n
                If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
                Call WriteUpdateHP(UserIndex)
 
          Else
            .Counters.Incinerado = RandomNumber(-100, 100) ' Invi variable :D
            .flags.Incinerado = 0

            If .flags.Incinerado = 0 Then
                Call WriteLocaleMsg(UserIndex, 389)
                reproducesonido = 0
                End If

        End If

End With

 

End Sub
Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    '**********************************************************
    'Author: Jopi
    'Guarda una descripcion detallada del error en Errores.log
    '**********************************************************
        
        On Error GoTo RegistrarError_Err
    
        
        
        'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
100     If Componente = HistorialError.Componente And _
           Numero = HistorialError.ErrorCode Then
       
           'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
            'x lo que no hace falta registrar el error.
102         If HistorialError.Contador = 10 Then Exit Sub
        
            'Agregamos el error al historial.
104         HistorialError.Contador = HistorialError.Contador + 1
        
        Else 'Si NO es igual, reestablecemos el contador.

106         HistorialError.Contador = 0
108         HistorialError.ErrorCode = Numero
110         HistorialError.Componente = Componente
            
        End If
    
        'Registramos el error en Errores.log
112     Dim File As Integer: File = FreeFile
        
114     Open DocPath & "Errores.log" For Append As #File
    
116         Print #File, "Error: " & Numero
118         Print #File, "Descripcion: " & Descripcion
        
120         If LenB(Linea) <> 0 Then
122             Print #File, "Linea: " & Linea
            End If
        
124         Print #File, "Componente: " & Componente
126         Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
128         Print #File, vbNullString
        
130     Close #File
    
132     Debug.Print "Error: " & Numero & vbNewLine & "Descripcion: " & Descripcion & vbNewLine & "Componente: " & Componente & vbNewLine & "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine

134     frmMain.AgregarConsola "Hay un problema con la linea: " & Linea & "    : " & Time$ & " " & Componente
        
        #If Debugger = 1 Then
136            MsgBox "Hay un problema con la linea: " & Linea & "    : " & Time$ & " " & Componente
        #End If
        
138     Exit Sub

RegistrarError_Err:
140     Call RegistrarError(Err.Number, Err.description, "ES.RegistrarError", Erl)
        
End Sub
Sub CargarCiudades()
    
    '***************************************************
    'Author: Mermas
    'Last Modification: 07/07/21
    'Jopi: Uso de clsIniManager para cargar los valores.
    'Mermas: optimizaciones de ciudades.dat
    '***************************************************
    
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando Ciudades.dat"
    
    Dim Lector As clsIniManager: Set Lector = New clsIniManager
    
    Call Lector.Initialize(DatPath & "Ciudades.dat")
        
    ReDim Ciudades(1 To Lector.GetValue("INIT", "CIUDADES"))
        
    NUMCIUDADES = UBound(Ciudades)
        
    Dim NameCity As String
        
    With tCiudades.Nix
         NameCity = "NIX"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
        
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
    
    With tCiudades.Illiandor
         NameCity = "ILLIANDOR"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
    
    With tCiudades.Ullathorpe
         NameCity = "ULLATHORPE"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
            
    With tCiudades.Banderbill
         NameCity = "BANDERBILL"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
            
    With tCiudades.Rinkel
         NameCity = "RINKEL"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
            
    With tCiudades.DungeonNewbie
         NameCity = "DUNGEONNEWBIE"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
            
    With tCiudades.Lindos
         NameCity = "LINDOS"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
            
    With tCiudades.Arghal
         NameCity = "ARGHAL"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
    
    With tCiudades.Tiama
         NameCity = "TIAMA"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
    
    With tCiudades.Orac
         NameCity = "ORAC"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
    
    With tCiudades.Suramei
         NameCity = "SURAMEI"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
              
    With tCiudades.Nueva
         NameCity = "NUEVA"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
                
        .Dead_Map = Lector.GetValue(NameCity, "MAPA_DEAD")
        .Dead_X = Lector.GetValue(NameCity, "X_DEAD")
        .Dead_Y = Lector.GetValue(NameCity, "Y_DEAD")
    End With
    
    With tCiudades.Prision
         NameCity = "PRISION"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
    End With
    
    With tCiudades.Libertad
         NameCity = "LIBERTAD"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
    End With
    
    With tCiudades.Intermundia
         NameCity = "INTERMUNDIA"
        .Map = Lector.GetValue(NameCity, "MAPA")
        .X = Lector.GetValue(NameCity, "X")
        .Y = Lector.GetValue(NameCity, "Y")
    End With
    
    Set Lector = Nothing
    
12  Ciudades(eCiudad.cUllathorpe) = tCiudades.Ullathorpe
13  Ciudades(eCiudad.cIlliandor) = tCiudades.Illiandor
14  Ciudades(eCiudad.cNix) = tCiudades.Nix
15  Ciudades(eCiudad.cBanderbill) = tCiudades.Banderbill
16  Ciudades(eCiudad.cRinkel) = tCiudades.Rinkel
17  Ciudades(eCiudad.cDungeonNewbie) = tCiudades.DungeonNewbie
18  Ciudades(eCiudad.cLindos) = tCiudades.Lindos
20  Ciudades(eCiudad.cARGHAL) = tCiudades.Arghal
19  Ciudades(eCiudad.cTIAMA) = tCiudades.Tiama
21  Ciudades(eCiudad.cORAC) = tCiudades.Orac
22  Ciudades(eCiudad.cSURAMEI) = tCiudades.Suramei
23  Ciudades(eCiudad.cNueva) = tCiudades.Nueva
24  Ciudades(eCiudad.cPrision) = tCiudades.Prision
25  Ciudades(eCiudad.cLibertad) = tCiudades.Libertad
26  Ciudades(eCiudad.cIntermundia) = tCiudades.Intermundia

    If frmMain.Visible Then frmMain.AgregarConsola "Se cargaron las ciudades.dat"

    Exit Sub

CargarCiudades_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.CargarCiudades", Erl)
        Resume Next
End Sub


Public Sub loadAdministrativeUsers()
    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM
    If frmMain.Visible Then frmMain.AgregarConsola "Cargando Administradores/Dioses/Gms."

    'Si esta mierda tuviese array asociativos el codigo seria tan lindo.
    Dim buf  As Integer

    Dim i    As Long

    Dim Name As String
       
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager

    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    ' Especiales
    buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next i
    
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    Set ServerIni = Nothing

    If frmMain.Visible Then frmMain.AgregarConsola "Los Administradores/Dioses/Gms se han cargado correctamente."

End Sub
Function cuantos(ByVal Cadena As String, ByVal caracter As String) As Integer
Dim i As Integer, num As Integer
For i = 1 To Len(Cadena)
If mid(Cadena, i, 1) = caracter Then num = num + 1
Next
cuantos = num
End Function

Public Function QueClaseEs(ByVal index As Byte) As String
'Estas funciones lo que hacen es arreglar los index de razas o clases, por ejemplo en IAO el Clerigo es el eClass = 4, y acá es el eClass = 1, lo que hace es acomodarlo
        Select Case index
        Case 1: QueClaseEs = "CLERIGO"
        Case 2: QueClaseEs = "MAGO"
        Case 3: QueClaseEs = "GUERRERO"
        Case 4: QueClaseEs = "ASESINO"
        Case 5: QueClaseEs = "LADRON"
        Case 6: QueClaseEs = "BARDO"
        Case 7: QueClaseEs = "DRUIDA"
        Case 8: QueClaseEs = "GLADIADOR"
        Case 9: QueClaseEs = "PALADIN"
        Case 10: QueClaseEs = "CAZADOR"
        Case 11: QueClaseEs = "PESCADOR"
        Case 12: QueClaseEs = "HERRERO"
        Case 13: QueClaseEs = "LEÑADOR"
        Case 14: QueClaseEs = "MINERO"
        Case 15: QueClaseEs = "CARPINTERO"
        Case 16: QueClaseEs = "SASTRE"
        Case 17: QueClaseEs = "MERCENARIO"
        Case 18: QueClaseEs = "NIGROMANTE"
        End Select
End Function

Public Function QueRazaEs(ByVal index As Byte) As Byte
'Estas funciones lo que hacen es arreglar los index de razas o clases, por ejemplo en IAO la raza Enano es el eRaza = 4, y acá es el eRaza= 2, lo que hace es acomodarlo
        Select Case index
        Case 1: QueRazaEs = 1  'Humano
        Case 2: QueRazaEs = 5  'Elfo
        Case 3: QueRazaEs = 2  'Drow
        Case 4: QueRazaEs = 3  'Gnomo
        Case 5: QueRazaEs = 4  'Enano
        Case 6: QueRazaEs = 6 'Orco
        End Select
End Function


Public Function QueRazaNombre(ByVal index As Byte) As String
'Estas funciones lo que hacen es arreglar los index de razas o clases, por ejemplo en IAO el Clerigo es el eClass = 4, y acá es el eClass = 1, lo que hace es acomodarlo
        Select Case index
        Case 1: QueRazaNombre = "Humano"
        Case 2: QueRazaNombre = "Elfo"
        Case 3: QueRazaNombre = "Drow"
        Case 4: QueRazaNombre = "Gnom"
        Case 5: QueRazaNombre = "Enano"
        Case 6: QueRazaNombre = "Oco"
        End Select
End Function
Public Sub AgregarParticula(ByVal UserIndex As Integer, ByVal Particle As Integer, ByVal Time As Long, ByVal Remove As Boolean, ByVal DesdeMessage As Boolean)

2
    If DesdeMessage = True Then
    
       If Remove = False Then
3             UserList(CharList(UserIndex)).Char.ParticulaFx = Particle
4             UserList(CharList(UserIndex)).Char.Loops = Time
5        Else
6            UserList(CharList(UserIndex)).Char.ParticulaFx = 0
7            UserList(CharList(UserIndex)).Char.Loops = 0
8        End If
     
     Else
     
       If Remove = False Then
              UserList(UserIndex).Char.ParticulaFx = Particle
             UserList(UserIndex).Char.Loops = Time
        Else
            UserList(UserIndex).Char.ParticulaFx = 0
            UserList(UserIndex).Char.Loops = 0
        End If
     
     
     End If
End Sub

Public Function GuardarConsultaIni(ByVal User As String, ByVal Consulta As String, ByVal Tipo As Byte)
        
           On Error GoTo GuardarConsultaIni_Err
           
100        Dim UserFile As String
102        Dim n As Integer
104        Dim NumConsultas As Integer
105        Dim Paso As Byte

103        Paso = 0

106        UserFile = DocConsultas & User & ".ini"
        
108        If Not FileExist(UserFile, vbArchive) Then
        
110            n = FreeFile
111            Paso = 1
112            Open UserFile For Binary Access Write As n
               
114            Put n, , "[INIT]" & vbCrLf
116            Put n, , "Nombre=" & User & vbCrLf
118            Put n, , "NumConsultas=0" & vbCrLf & vbCrLf
            
120            Put n, , "[CONSULTAS]" & vbCrLf
            
122            Close #n

123            Paso = 0

124        End If
            
126        If FileExist(UserFile, vbArchive) Then
        
128            NumConsultas = CInt(GetVar(UserFile, "INIT", "NumConsultas"))
130            NumConsultas = NumConsultas + 1
            
132            Call WriteVar(UserFile, "CONSULTAS", "C" & NumConsultas, Format(Time, "hh:mm") & "-" & Tipo & "-" & Consulta & "-" & "0")
133            Call WriteVar(UserFile, "INIT", "NumConsultas", NumConsultas)

134        End If

136        Exit Function
        
GuardarConsultaIni_Err:
137     If Paso = 1 Then Close #n
138     Call RegistrarError(Err.Number, Err.description, "ES.GuardarConsultaIni", Erl)
140     Resume Next
        
End Function
 
