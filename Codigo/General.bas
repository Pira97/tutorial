Attribute VB_Name = "General"
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


Global LeerNPCs As New clsIniManager

    Public PorcentajeSkill As Byte
    Public mapasegundos As Integer
Sub DarCuerpoDesnudo(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/14/07
    'Da cuerpo desnudo a un usuario
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************

    Dim CuerpoDesnudo As Integer

    With UserList(UserIndex)
    
      If .flags.Navegando = 1 Or .flags.Montando = 1 Then Exit Sub
      
        Select Case .Genero

            Case eGenero.Hombre

                Select Case .raza

                    Case eRaza.Humano
                        CuerpoDesnudo = 21

                    Case eRaza.Drow
                        CuerpoDesnudo = 32

                    Case eRaza.Elfo
                        CuerpoDesnudo = 21

                    Case eRaza.gnomo
                        CuerpoDesnudo = 53

                    Case eRaza.enano
                        CuerpoDesnudo = 53
                        
                    Case eRaza.Orco
                        CuerpoDesnudo = 248
                    
                End Select

            Case eGenero.Mujer

                Select Case .raza

                    Case eRaza.Humano
                        CuerpoDesnudo = 39

                    Case eRaza.Drow
                        CuerpoDesnudo = 40

                    Case eRaza.Elfo
                        CuerpoDesnudo = 39

                    Case eRaza.gnomo
                        CuerpoDesnudo = 60

                    Case eRaza.enano
                        CuerpoDesnudo = 60

                    Case eRaza.Orco
                        CuerpoDesnudo = 249
                        
                End Select

        End Select
        
 
        .Char.body = CuerpoDesnudo
    
        .flags.Desnudo = 1

    End With

End Sub

Sub Bloquear(ByVal toMap As Boolean, _
             ByVal sndIndex As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal b As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
    '***************************************************

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)

    End If

End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then

        With MapData(Map, X, Y)

            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or ( _
                    .Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And .Graphic(2) = 0 Then
                HayAgua = True
            Else
                HayAgua = False

            End If

        End With

    Else
        HayAgua = False

    End If

End Function

Private Function HayLava(ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer) As Boolean

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/12/07
    '***************************************************
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False

        End If

    Else
        HayLava = False

    End If

End Function
Function HayCura(ByVal UserIndex As Integer) As Boolean
'******************************
'Adaptacion a 13.0: Kaneidra
'Last Modification: 15/05/2012
'******************************
 
Dim X As Integer, Y As Integer
 
For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
       
            If MapData(UserList(UserIndex).Pos.Map, X, Y).npcindex > 0 Then
                    If Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).npcindex).NPCtype = 1 Then
                        If Distancia(UserList(UserIndex).Pos, Npclist(MapData(UserList(UserIndex).Pos.Map, X, Y).npcindex).Pos) < 10 Then
                        HayCura = True
                        Exit Function
                    End If
                End If
            End If
           
        Next X
Next Y
 
HayCura = False
 
End Function
Sub Main()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************


    On Error GoTo Handler
        
        #If Debugger = 1 Then
        frmMain.Command4.Enabled = True
        #End If
        
100    Call BanIpCargar
    
102    frmCargando.Show
    
      'Desconectamos las cuentas por si quedó alguna bugeada / Mermas
103    Call DesconectarCuenta

      ' Constants & vars
104    frmCargando.Label1(2).Caption = "Cargando constantes..."
106    Call LoadConstants
108    DoEvents
    
    ' Arrays
110    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
112    Call LoadArrays
    
     'Cargamos Indices de objetos al nacer
113    Call LoadObjNacimiento
    ' Server.ini & Apuestas.dat & Ciudades.dat
114    frmCargando.Label1(2).Caption = "Cargando Server.ini"
116    Call LoadSini
118    Call CargarCiudades
120    Call CargaApuestas
    
    ' Npcs.dat
122    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
124    Call CargaNpcsDat
 
 
    ' Obj.dat
126    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
128    Call LoadOBJData
 
    ' Hechizos.dat
130    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
132    Call CargarHechizos
   
   
     ' Objetos de Herreria
134    frmCargando.Label1(2).Caption = "Cargando Objetos de Herreria"
136    Call LoadArmasHerreria
138    Call LoadArmadurasHerreria
140    Call LoadCascosHerreria      ' $ Shermie80 $
142    Call LoadEscudosHerreria     ' $ Shermie80 $
 
    
    ' Objetos de Capinteria
144    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpinteria"
146    Call LoadObjCarpintero
    
    
148    frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
150    Call LoadObjDruida
    
152    frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"
154    Call LoadObjSastre
       
    ' Balance.dat
156    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
158    Call LoadBalance
    
160    If BootDelBackUp Then
162        frmCargando.Label1(2).Caption = "Cargando BackUp"
164        Call CargarBackUp
166    Else
168        frmCargando.Label1(2).Caption = "Cargando Mapas"
170        Call LoadMapData
172    End If
    
    
186   MapInfo(tCiudades.Prision.Map).ResuSinEfecto = 1
        ''//


188    Call SonidosMapas.LoadSoundMapInfo

190    Dim loopc As Integer
    
    'Resetea las conexiones de los usuarios
192    For loopc = 1 To MaxUsers
194        UserList(loopc).ConnID = -1
196        UserList(loopc).ConnIDValida = False
198        Set UserList(loopc).incomingData = New clsByteQueue
200        Set UserList(loopc).outgoingData = New clsByteQueue
202    Next loopc
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    
    ' Timers
204    Call InitMainTimers
 
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Configuracion de los sockets
    
206    Call SecurityIp.InitIpTables(1000)
    
208    If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
210      Call IniciaWsApi(frmMain.hwnd)
212      SockListen = ListenForConnect(Puerto, hWndMsg, "")
      
214    If SockListen <> -1 Then
216        Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen) ' Guarda el socket escuchando
    Else
218        MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly
220    End If
    
    
222    If frmMain.Visible Then frmMain.AgregarConsola "Escuchando conexiones entrantes ..."
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
224    Unload frmCargando
        
        LogServerStartTime
        
    'Ocultar
226      If HideMe Then
228        Call frmMain.InitMain(1)
230      Else
232        Call frmMain.InitMain(0)
234      End If
    
236    tInicioServer = GetTickCount() And &H7FFFFFFF
 
    'Este ultimo es para saber siempre los records en el frmMain
238 frmMain.Label1.Caption = "Número de record de usuarios : " & recordusuarios
    
 
        Exit Sub
Handler:
240    Call RegistrarError(Err.Number, Err.description, "General.Main", Erl)
242    Resume Next

End Sub

Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************

    FileExist = LenB(dir$(File, FileType)) <> 0

End Function

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
    '*****************************************************************
    'Gets a field from a string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    'Gets a field from a delimited string
    '*****************************************************************

    Dim i          As Long
    Dim LastPos    As Long
    Dim CurrentPos As Long
    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

End Function

Function MapaValido(ByVal Map As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    MapaValido = Map >= 1 And Map <= NumMaps

End Function

Sub MostrarNumUsers()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    frmMain.CantUsuarios.Caption = "Número de usuarios jugando: " & NumUsers

End Sub

Public Sub LogCriticEvent(desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogIndex(ByVal index As Integer, ByVal desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogError(desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogStatic(desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogTarea(desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & desc
    Close #nfile

    Exit Sub

ErrHandler:

End Sub

Public Sub LogClanes(ByVal str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogDesarrollo(ByVal str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************ç

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogAsesinato(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogHackAttemp(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogCheating(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogAntiCheat(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Arg As String
    Dim i   As Integer
    
    For i = 1 To 33
    
        Arg = ReadField(i, cad, 44)
    
        If LenB(Arg) = 0 Then Exit Function
    
    Next i
    
    ValidInputNP = True

End Function

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 15/11/2009
    '15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '**************************************************************

    With UserList(UserIndex)

        If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> 2 And _
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> 4 Then Intemperie = True
        Else
            Intemperie = False

        End If

    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False

End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    For i = 1 To MAXMASCOTAS

        With UserList(UserIndex)

            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = Npclist(.MascotasIndex( _
                            i)).Contadores.TiempoExistencia - DeltaTick

                    If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotasIndex( _
                            i), 0)

                End If

            End If

        End With

    Next i

End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
    '***************************************************
    'Autor: Unkonwn
    'Last Modification: 23/11/2009
    'If user is naked and it's in a cold map, take health points from him
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim modifi As Integer
    Dim Mapa As String
    Mapa = UserList(UserIndex).Pos.Map
    With UserList(UserIndex)

        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + DeltaTick
        Else
        If MapInfo(.Pos.Map).Terreno = Nieve Then
            'If Mapa = 217 Or Mapa = 218 Or Mapa = 219 Or Mapa = 220 Or Mapa = 221 Or Mapa = 222 Or Mapa = 223 Or Mapa = 224 Or Mapa = 225 Or Mapa = 226 Or Mapa = 227 Or Mapa = 228 Or Mapa = 229 Or Mapa = 230 Or Mapa = 231 Or Mapa = 232 Or Mapa = 233 Or Mapa = 250 Then
                Call WriteLocaleMsg(UserIndex, 46)
                modifi = Porcentaje(.Stats.MaxHP, 5)
                .Stats.MinHP = .Stats.MinHP - modifi
                
                If .Stats.MinHP < 1 Then
                    .Stats.MinHP = 0
                    Call UserDie(UserIndex)

                End If
                
                Call WriteUpdateHP(UserIndex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(UserIndex, modifi)
                Call WriteUpdateSta(UserIndex)

            End If
            
            .Counters.Frio = 0

        End If

    End With

End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo EfectoInvisibilidad_Err
    
    With UserList(UserIndex)

        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + DeltaTick
        Else
            .Counters.Invisibilidad = 0
            .flags.Invisible = 0

            If .flags.Oculto = 0 Then
                Call WriteLocaleMsg(UserIndex, 307)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
             End If
             
        End If

    End With

    Exit Sub

EfectoInvisibilidad_Err:
182     Call RegistrarError(Err.Number, Err.description, "General.EfectoInvisibilidad", Erl)
184     Resume Next
        
End Sub

Public Sub EfectoParalisisNpc(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(npcindex)

        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0

        End If

    End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)

        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - DeltaTick

            If .Counters.Ceguera <= 0 Then
                If .flags.Estupidez = 1 Then
                    .flags.Estupidez = 0
                    Call WriteDumbNoMore(UserIndex)
                Else
                    Call WriteBlindNoMore(UserIndex)
                End If
            End If
        End If

    End With

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
 
 

    With UserList(UserIndex)
    
        If .Counters.Paralisis > 0 Then
        
            Dim CasterIndex As Integer

            CasterIndex = .flags.ParalizedByIndex
        
            ' Only aplies to non-magic clases
            If .Stats.MaxMAN = 0 Then

                ' Paralized by user?
                If CasterIndex <> 0 Then
                
                    ' Close? => Remove Paralisis
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                        
                        ' Caster dead? => Remove Paralisis
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                    
                    ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then

                        ' Out of vision range? => Reduce paralisis counter
                        If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                            ' Aprox. 1500 ms
                            .Counters.Paralisis = IntervaloParalizadoReducido
                            Exit Sub

                        End If

                    End If
                
                    ' Npc?
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex
                    
                    ' Paralized by npc?
                    If CasterIndex <> 0 Then
                    
                        If .Counters.Paralisis > IntervaloParalizadoReducido Then

                            ' Out of vision range? => Reduce paralisis counter
                            If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                ' Aprox. 1500 ms
                                .Counters.Paralisis = IntervaloParalizadoReducido
                                Exit Sub

                            End If

                        End If

                    End If
                    
                End If

            End If
            
            .Counters.Paralisis = .Counters.Paralisis - DeltaTick

            If .Counters.Paralisis <= 0 Then
                Call RemoveParalisis(UserIndex)
            End If

        End If

    End With

End Sub
Public Sub RemoveParalisis(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .Counters.Paralisis = 0
        Call WriteParalizeOK(UserIndex)
    End With


End Sub
Public Sub RecStamina(ByVal UserIndex As Integer, _
                      ByVal DeltaTick As Single, _
                      ByRef EnviarStats As Boolean, _
                      ByVal Intervalo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    With UserList(UserIndex)
    
        ' $ Shermie80 / fix
       ' If UserList(UserIndex).flags.Trabajando Then Exit Sub
        
        If UserList(UserIndex).flags.Desnudo Then
            If Not UserList(UserIndex).flags.Montando Then Exit Sub
        End If
    
    
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 2 And _
                MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 4 Then Exit Sub
        
        Dim massta As Integer

        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + DeltaTick
               
            Else
                EnviarStats = True
                .Counters.STACounter = 0
                
                 massta = 5 * RandomNumber(1, Porcentaje(.Stats.MaxSta + .Stats.UserSkills(eSkill.Supervivencia), 5))
                .Stats.MinSta = .Stats.MinSta + (massta)

                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta

                End If

            End If

        End If

    End With
    
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer, ByVal DeltaTick As Single)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .Counters.Veneno < IntervaloVeneno Then
            .Counters.Veneno = .Counters.Veneno + DeltaTick
        Else
            Call WriteLocaleMsg(UserIndex, 47, "", 0, 5)
            .Counters.Veneno = 0
            n = RandomNumber(1, 5)
            .Stats.MinHP = .Stats.MinHP - n

            If .Stats.MinHP < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)

        End If

    End With

End Sub

Public Function TieneSacri(ByVal UserIndex As Integer) As Byte
On Error Resume Next
    Dim i As Long
    Dim ObjInd As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjInd = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ObjInd > 0 Then
            If UserList(UserIndex).Invent.Object(i).Equipped = 1 And ObjData(ObjInd).EfectoMagico = eMagicType.Sacrificio Then
                TieneSacri = CByte(i)
                Exit Function
            End If
        End If
    Next i
    
    TieneSacri = 0

End Function

Public Sub DuracionPociones(ByVal UserIndex As Integer, ByVal DeltaTick As Single)

    '***************************************************
    'Author: ??????
    'Last Modification: 11/27/09 (Budi)
    'Cuando se pierde el efecto de la poción updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
    '***************************************************
    With UserList(UserIndex)

        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
            .flags.DuracionEfecto = .flags.DuracionEfecto - DeltaTick

            If .flags.DuracionEfecto = 0 Then
                .flags.TomoPocion = False
                .flags.TipoPocion = 0
                'volvemos los atributos al estado normal
                Dim loopX As Long
                For loopX = 1 To NUMATRIBUTOS
                .Stats.UserAtributos(loopX) = .Stats.UserAtributosBackUP(loopX)
                Next
                Call WriteUpdateDexterity(UserIndex)
                Call WriteUpdateStrenght(UserIndex)
            End If

        End If

    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByVal DeltaTick As Single, ByRef fenviarAyS As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)

        If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        
        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + DeltaTick
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - 10
                
                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1

                End If
                
                fenviarAyS = True

            End If

        End If
        
        'hambre
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + DeltaTick
            Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - 10

                If .Stats.MinHam <= 0 Then
                    .Stats.MinHam = 0
                    .flags.Hambre = 1

                End If

                fenviarAyS = True

            End If

        End If

    End With

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, _
                 ByVal DeltaTick As Single, _
                 ByRef EnviarStats As Boolean, _
                 ByVal Intervalo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 1 And MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 2 And _
                MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 4 Then Exit Sub
                
        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
        If .Stats.MinHP < .Stats.MaxHP Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + DeltaTick
            Else
                mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                
                .Counters.HPCounter = 0
                .Stats.MinHP = .Stats.MinHP + mashit

                If .Stats.MinHP > .Stats.MaxHP Then .Stats.MinHP = .Stats.MaxHP
                Call WriteLocaleMsg(UserIndex, 389)
                EnviarStats = True

            End If

        End If

    End With

End Sub

Public Sub CargaNpcsDat(Optional ByVal ForzarActualizacionNpcsExistentes As Boolean = False)
    '***************************************************
    'Author: Unknown
    'Last Modification: 06/07/2020 (Cuicui)
    ' 06/07/2020 (Cuicui) - Actualizamos la informacion de los NPC's ya spawneados.
    '***************************************************
    
100    On Error GoTo CargaNpcsDat_Err
    
102    If frmMain.Visible Then frmMain.AgregarConsola "Cargando NPCs.dat."
    
    ' Leemos el NPCs.dat y lo almacenamos en la memoria.
104    Set LeerNPCs = New clsIniManager

106    Call LeerNPCs.Initialize(DatPath & "NPCs.dat")

    '' Actualizamos la informacion de los NPC's ya spawneados.
110    If ForzarActualizacionNpcsExistentes Then

112        Dim i As Long

114        For i = 1 To MAXNPCS

116            If Npclist(i).flags.NPCActive Then
118                Call ReloadNPCByIndex(i)
120            End If

122            DoEvents

124        Next i

126    End If
    
128    If frmMain.Visible Then frmMain.AgregarConsola "Se cargo el archivo NPCs.dat."
    

    Exit Sub

CargaNpcsDat_Err:
130         Call RegistrarError(Err.Number, Err.description, "General.CargaNpcsDat", Erl)
132         Resume Next
        
End Sub

Sub PasarSegundo()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler
    Dim UserIndex  As Integer
    Dim i As Long
 
    Call FuncionCR
    
    For i = 1 To LastUser
    
        If UserList(i).Counters.TiempoDeMapeo > 0 Then
            UserList(i).Counters.TiempoDeMapeo = UserList(i).Counters.TiempoDeMapeo - 1
        End If
        
        If UserList(i).flags.UserLogged Then

            Call PasarSegundotelep(i)
            
            Call PasarSegundoCarcel(i)

            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
            
               UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
               
                If UserList(i).Counters.Salir < 1 Then
                    Call WriteDisconnect(i)
                    Call FlushBuffer(i)
                    Call CloseSocket(i)
                Else
                    Call WriteLocaleMsg(i, 491, UserList(i).Counters.Salir)
                End If
                
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)

    Resume Next

End Sub
 
 
 
Sub GuardarUsuarios()
    
    On Error GoTo Handler
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(180)) 'Grabando personajes
        
    Dim i As Integer
    
    For i = 1 To LastUser
    
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i)
        End If
        
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(436)) 'Personajes grabados

    If frmMain.Visible Then frmMain.AgregarConsola "Guardado PJs completo" & " " & Time

    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.description, "General.GuardarUsuarios")
    Resume Next
End Sub

Public Sub FreeNPCs()
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all NPC Indexes
    '***************************************************
    Dim loopc As Long
    
    ' Free all NPC indexes
    For loopc = 1 To MAXNPCS
        Npclist(loopc).flags.NPCActive = False
    Next loopc

End Sub

Public Sub FreeCharIndexes()
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all char indexes
    '***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

End Sub

Public Function Tilde(data As String) As String
 
    Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function
Public Sub CargarELU()

      On Error GoTo Handler
       If frmMain.Visible Then frmMain.AgregarConsola "Cargando elu.dat"
100    Dim Leer     As clsIniManager
    
102    Set Leer = New clsIniManager
    
104    Call Leer.Initialize(DatPath & "Niveles.dat")
    
106    Dim loopc As Long
    
108    For loopc = 1 To (STAT_MAXELV - 1)
110          levelELU(loopc) = Leer.GetValue("INIT", "Nivel" & loopc)

112    Next loopc
    
114    Set Leer = Nothing
If frmMain.Visible Then frmMain.AgregarConsola "Se cargo elu.dat. Operacion Realizada con exito."
116        Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.description, "General.CargarELU")
    Resume Next
End Sub
Public Sub PasarSegundoCarcel(ByVal i As Integer)
                If Not UserList(i).Pos.Map = 0 Then
                    'Counter de piquete
                    If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).Trigger = eTrigger.ANTIPIQUETE Then
                            If UserList(i).flags.Muerto = 0 Then
                                 UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                                 UserList(i).Counters.ContadorPiquete = UserList(i).Counters.ContadorPiquete + 1
                                If UserList(i).Counters.ContadorPiquete = 6 Then
                                    Call WriteLocaleMsg(i, 70)
                                    UserList(i).Counters.ContadorPiquete = 0
                                End If
                                If UserList(i).Counters.PiqueteC >= 30 Then
                                     UserList(i).Counters.PiqueteC = 0
                                     UserList(i).Counters.ContadorPiquete = 0
                                    Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                                End If
                        Else
                             UserList(i).Counters.PiqueteC = 0
                        End If
                    Else
                         UserList(i).Counters.PiqueteC = 0
                    End If
                End If
End Sub

Public Function EquipaSkin(ByVal UserIndex As Integer, ByVal Equipo As Integer, ByVal BackOrNext As Integer)
 
    With UserList(UserIndex)
    
    If Not .Donador.activo = 1 Then
        Call WriteLocaleMsg(UserIndex, 377)
        Exit Function
    End If
    
    If .flags.Montando = 1 Then
        Call WriteLocaleMsg(UserIndex, 21)
        Exit Function
    End If

    If .flags.Navegando = 1 Then
        Call WriteLocaleMsg(UserIndex, 20)
        Exit Function
    End If
 
 
    If BackOrNext = 1 Then ' Siguiente
    
        Select Case Equipo ' Que parte del cuerpo tiene
        
            Case 1 'Armaduras
            
                If ObjData(.Invent.ArmourEqpObjIndex).CantidadSkin > 0 Then
                    
                    If .Char.BodySkin > 0 Then
                    
                        If .Char.BodySkin <= ObjData(.Invent.ArmourEqpObjIndex).CantidadSkin Then
                            .Char.BodySkin = .Char.BodySkin + 1
                            .Char.body = ObjData(.Invent.ArmourEqpObjIndex).TieneSkin(.Char.BodySkin - 1)
                             Exit Function
                        End If
                    
                    Else
                    
                    .Char.body = ObjData(.Invent.ArmourEqpObjIndex).TieneSkin(0)
                    .Char.BodySkin = 1
                     Exit Function
                    End If
                    
                    'WriteMarcamosSkin UserIndex, 2 'marcamos tilde en el frMskIN
                End If
                
              
            Case 2
            
        End Select
        
    Else ' Anterior
        
    
        Select Case Equipo ' Que parte del cuerpo tiene
        
            Case 1 'Armaduras
            
                 If ObjData(.Invent.ArmourEqpObjIndex).CantidadSkin > 0 Then
                
                    If .Char.BodySkin > 1 Then
                      
                        .Char.BodySkin = .Char.BodySkin - 1
                        .Char.body = ObjData(.Invent.ArmourEqpObjIndex).TieneSkin(.Char.BodySkin - 1)
                         Exit Function
                        
                    Else
                    
                        .Char.BodySkin = 0
                        .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                         Exit Function
                        
                    End If
                    
                End If
                
 
            Case 2
            
        End Select
        
    
    End If
    
 
     
    
    End With
    
 End Function
 Public Sub MovimientoFriend(ByVal UserIndex As Integer)

        If UserList(UserIndex).flags.CantidadAmigos = 0 Then Exit Sub

        Dim i As Byte
        Dim slot As Byte
   
        For i = 1 To UserList(UserIndex).flags.CantidadAmigos
           If UserList(UserIndex).Amigos(UserList(UserIndex).flags.CantidadAmigos).index <= 0 Then
           UserList(UserIndex).flags.CheckAmigos = 0
           Debug.Print "se desconectaron todos tus amigos" & UserList(UserIndex).Name
           Exit Sub
           End If
        Next i
   
        For i = 1 To UserList(UserIndex).flags.CantidadAmigos
        If UserList(UserIndex).Amigos(i).index > 0 Then
        slot = BuscarSlotAmigoNameSlot(UserList(UserIndex).Amigos(i).index, UserList(UserIndex).Name)
        Writemostrarubicacion UserList(UserIndex).Amigos(i).index, UserList(UserIndex).Name, slot, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y
        End If
        Next i
End Sub
 

Public Function ColorNick(Status As Byte) As Byte

Select Case Status
Case 1 'Rene
ColorNick = 1
Case 2, 5 'Ciuda, Armi
ColorNick = 2
Case 3, 6  'Repu /Mili
ColorNick = 4
Case 4 'Caos
ColorNick = 3
Case Else 'Gris x las dudas
ColorNick = 1
End Select

End Function
 
Private Sub InitMainTimers()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Initializes Main Timers.
    '*****************************************************************
    On Error Resume Next

    Call frmMain.CargarLabelsMain
    
 
    With frmMain
        .AutoSave.Enabled = True
        .GameTimer.Enabled = True
        .packetResend.Enabled = True
        .TIMER_AI.Enabled = True
        .Auditoria.Enabled = True
    End With
    
    
    LastGameTick = GetTickCount
    
End Sub
 
 

Private Sub LoadConstants()

    On Error GoTo Handler
       
       
100    LastBackup = Format(Now, "Short Time")
102    Minutos = Format(Now, "Short Time")
    
    'Dir
104    IniPath = App.Path & "\"
106    DatPath = App.Path & "\Dat\"
108    AccountPath = App.Path & "\Cuentas\"
109    DocConsultas = App.Path & "\Documentacion\Soporte\"
110    CharPath = App.Path & "\Charfile\"
111    DocPath = App.Path & "\Documentacion\"
 
    ' Verifico si existe la carpeta donde se guardan las cuentas.
112    If LenB(dir$(AccountPath, vbDirectory)) = 0 Then
114        Call MkDir(AccountPath)
116    End If
    
    ' Verifico si existe la carpeta donde se guardan los personajes.
118    If LenB(dir$(CharPath, vbDirectory)) = 0 Then
120        Call MkDir(CharPath)
122    End If
    
 
138    LogError ("///////////////////////////////// Se levantó el servidor /////////////////////////////////")
    
    
140    LevelSkill(1).LevelValue = 3
142    LevelSkill(2).LevelValue = 5
144    LevelSkill(3).LevelValue = 7
146    LevelSkill(4).LevelValue = 10
148    LevelSkill(5).LevelValue = 13
150    LevelSkill(6).LevelValue = 15
152    LevelSkill(7).LevelValue = 17
154    LevelSkill(8).LevelValue = 20
156    LevelSkill(9).LevelValue = 23
158    LevelSkill(10).LevelValue = 25
160    LevelSkill(11).LevelValue = 27
162    LevelSkill(12).LevelValue = 30
164    LevelSkill(13).LevelValue = 33
166    LevelSkill(14).LevelValue = 35
168    LevelSkill(15).LevelValue = 37
170    LevelSkill(16).LevelValue = 40
172    LevelSkill(17).LevelValue = 43
174    LevelSkill(18).LevelValue = 45
176    LevelSkill(19).LevelValue = 47
178    LevelSkill(20).LevelValue = 50
180    LevelSkill(21).LevelValue = 53
182    LevelSkill(22).LevelValue = 55
184    LevelSkill(23).LevelValue = 57
186    LevelSkill(24).LevelValue = 60
188    LevelSkill(25).LevelValue = 63
190    LevelSkill(26).LevelValue = 65
192    LevelSkill(27).LevelValue = 67
194    LevelSkill(28).LevelValue = 70
196    LevelSkill(29).LevelValue = 73
198    LevelSkill(30).LevelValue = 75
200    LevelSkill(31).LevelValue = 77
202    LevelSkill(32).LevelValue = 80
204    LevelSkill(33).LevelValue = 83
206    LevelSkill(34).LevelValue = 85
208    LevelSkill(35).LevelValue = 87
210    LevelSkill(36).LevelValue = 90
212    LevelSkill(37).LevelValue = 93
214    LevelSkill(38).LevelValue = 95
216    LevelSkill(39).LevelValue = 97
218    LevelSkill(40).LevelValue = 100
220    LevelSkill(41).LevelValue = 100
222    LevelSkill(42).LevelValue = 100
224    LevelSkill(43).LevelValue = 100
226    LevelSkill(44).LevelValue = 100
228    LevelSkill(45).LevelValue = 100
230    LevelSkill(46).LevelValue = 100
232    LevelSkill(47).LevelValue = 100
234    LevelSkill(48).LevelValue = 100
236    LevelSkill(49).LevelValue = 100
238    LevelSkill(50).LevelValue = 100
    
240    ListaRazas(eRaza.Humano) = "Humano"
242    ListaRazas(eRaza.Elfo) = "Elfo"
246    ListaRazas(eRaza.Drow) = "Drow"
248    ListaRazas(eRaza.gnomo) = "Gnomo"
250    ListaRazas(eRaza.enano) = "Enano"
252    ListaRazas(eRaza.Orco) = "Orco"
    
254    ListaClases(eClass.Mago) = "Mago"
256    ListaClases(eClass.Clerigo) = "Clerigo"
258    ListaClases(eClass.Guerrero) = "Guerrero"
260    ListaClases(eClass.Asesino) = "Asesino"
262    ListaClases(eClass.ladron) = "Ladron"
264    ListaClases(eClass.Bardo) = "Bardo"
266    ListaClases(eClass.Druida) = "Druida"
268    ListaClases(eClass.Paladin) = "Paladin"
270    ListaClases(eClass.Cazador) = "Cazador"
272    ListaClases(eClass.PescadoR) = "Pescador"
274    ListaClases(eClass.Herrero) = "Herrero"
276    ListaClases(eClass.Leñador) = "Leñador"
278    ListaClases(eClass.Minero) = "Minero"
280    ListaClases(eClass.Carpintero) = "Carpintero"
282    ListaClases(eClass.Mercenario) = "Mercenario"
284    ListaClases(eClass.Nigromante) = "Nigromante"
286    ListaClases(eClass.Sastre) = "Sastre"
288    ListaClases(eClass.Gladiador) = "Gladiador"
    
290    SkillsNames(eSkill.Suerte) = "Suerte"
292    SkillsNames(eSkill.magia) = "Magia"
294    SkillsNames(eSkill.robar) = "Robar"
296    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
298    SkillsNames(eSkill.armas) = "Combate con armas"
300    SkillsNames(eSkill.Meditar) = "Meditar"
302    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
304    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
306    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
308    SkillsNames(eSkill.talar) = "Talar arboles"
310    SkillsNames(eSkill.comerciar) = "Comercio"
312    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
314    SkillsNames(eSkill.pesca) = "Pesca"
316    SkillsNames(eSkill.mineria) = "Mineria"
318    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
320    SkillsNames(eSkill.Herreria) = "Herreria"
322    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
324    SkillsNames(eSkill.domar) = "Domar animales"
326    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
328    SkillsNames(eSkill.Wrestling) = "Artes marciales"
330    SkillsNames(eSkill.Navegacion) = "Navegación"
332    SkillsNames(eSkill.Resistencia) = "Resistencia mágica"
334    SkillsNames(eSkill.ArmasArrojadizas) = "Armas arrojadizas"
336    SkillsNames(eSkill.alquimia) = "Alquimia"
338    SkillsNames(eSkill.botanica) = "Botanica"
340    SkillsNames(eSkill.Sastreria) = "Sastreria"
342    SkillsNames(eSkill.Equitacion) = "Equitacion"
    
344    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
346    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
348    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
350    ListaAtributos(eAtributos.Carisma) = "Carisma"
352    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    'Bordes del mapa
354    MinXBorder = XMinMapSize + (XWindow \ 2)
356    MaxXBorder = XMaxMapSize - (XWindow \ 2)
358    MinYBorder = YMinMapSize + (YWindow \ 2)
360    MaxYBorder = YMaxMapSize - (YWindow \ 2)

362    Call seguridad_clones_construir

364    MaxUsers = 0
    
    Exit Sub
        
Handler:
366    Call RegistrarError(Err.Number, Err.description, "General.LoadConstants", Erl)
368    Resume Next

End Sub

Private Sub LoadArrays()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Loads all arrays
    '*****************************************************************
    On Error GoTo Handler

    ' Load guilds info
100    Call LoadGuildsDB
    
    ' Load forbidden words
102    Call CargarForbidenWords

    'Cargamos niveles
104    Call CargarELU
    
    Exit Sub
        
Handler:
106    Call RegistrarError(Err.Number, Err.description, "General.LoadArrays", Erl)
108    Resume Next

End Sub

 
Public Sub LogServerStartTime()
    
    'Log
240    Dim n As Integer
242    n = FreeFile
244    Open App.Path & "\logs\Main.log" For Append Shared As #n
246    Print #n, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
248    Close #n
    
End Sub
Public Function OroLleno(ByVal UserIndex As Integer, ByVal Valor1 As Long, Optional valor2 As Long = 0) As Boolean
OroLleno = False

With UserList(UserIndex)
Dim RemainingAmountToMaximumGold As Long
RemainingAmountToMaximumGold = 2147483647 - .Stats.GLD

If Valor1 > 2147483647 Then 'And RemainingAmountToMaximumGold >= valor2 Then
OroLleno = True
Exit Function
Else
OroLleno = False
End If

End With

End Function
 


Sub PasarSegundotelep(ByVal i As Integer)

    On Error GoTo Handler
 
    With UserList(i)
    
    If .Counters.CreoTeleport Then
    
        .Counters.TimeTeleport = .Counters.TimeTeleport + 1

        Select Case .Counters.TimeTeleport 'Cases de segundos
        
        Case 5
          'Mermas nuevo, chequeamos si a la hora de crear el TP no pusieron cosas xd
        
        Dim Mapa As Integer, X As Integer, Y As Integer
        Dim Cancelo As Boolean
        
        Mapa = .flags.DondeTiroMap
        X = .flags.DondeTiroX
        Y = .flags.DondeTiroY
        
        Cancelo = False
        
          If .Pos.Map = tCiudades.Prision.Map Or .Pos.Map = Ciudades(eCiudad.cIntermundia).Map Or MapInfo(.Pos.Map).Pk = False Then
            Call WriteLocaleMsg(i, 448)
            Cancelo = True
          End If
          
        If Not LegalPos(Mapa, X, Y) Then
            Cancelo = True
        End If
        
         If MapData(Mapa, X, Y).ObjInfo.ObjIndex Then
             Call WriteLocaleMsg(i, 257)
             Cancelo = True
         End If
        
         If MapData(Mapa, X, Y).TileExit.Map Then
             Call WriteLocaleMsg(i, 257)
             Cancelo = True
         End If
        
         If MapData(Mapa, X, Y).Blocked Then
            Call WriteLocaleMsg(i, 257)
            Cancelo = True
         End If
         
         If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then
            Cancelo = True
        End If
        
        If Cancelo = True Then
            Call SendData(SendTarget.ToPCArea, i, PrepareMessageEfectoTerrenoParticula(34, X, Y, 5000))
            Call ControlarPortalLum(i, 0)
            Exit Sub
        End If
        
        If .Pos.Map = .flags.DondeTiroMap Then 'Si se va del mapa lo cancela
        
            MapData(Mapa, X, Y).TileExit.Map = Ciudades(eCiudad.cIntermundia).Map
            MapData(Mapa, X, Y).TileExit.X = Ciudades(eCiudad.cIntermundia).X
            MapData(Mapa, X, Y).TileExit.Y = Ciudades(eCiudad.cIntermundia).Y
             
            Dim ET As Obj: ET.Amount = 1: ET.ObjIndex = 672

            Call MakeObj(ET, Mapa, X, Y)
        
            Call SendData(SendTarget.ToPCArea, i, PrepareMessageEfectoTerrenoParticula(34, X, Y, 5000))
            
            .flags.CasteandoPortal = False
            
        End If
        
        Case Is >= 15: Call ControlarPortalLum(i, 1)
        
        End Select
        
    End If
    
    End With
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.description, "General.PasarSegundotelep", Erl)
    Resume Next

End Sub




Sub ControlarPortalLum(ByVal UserIndex As Integer, ByVal Modo As Byte)
    
    On Error GoTo Handler
    
    
    If UserList(UserIndex).Counters.CreoTeleport = True And UserList(UserIndex).flags.CasteandoPortal = False Then
       Call EraseObj(672, UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY)
       
       MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.Map = 0
       MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.X = 0
       MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.Y = 0
       
       UserList(UserIndex).flags.DondeTiroMap = 0
       UserList(UserIndex).flags.DondeTiroX = 0
       UserList(UserIndex).flags.DondeTiroY = 0
       UserList(UserIndex).Counters.TimeTeleport = 0
       UserList(UserIndex).Counters.CreoTeleport = False
Debug.Print 2

    ElseIf UserList(UserIndex).flags.CasteandoPortal = True Or Modo = 0 Then
        Call WriteLocaleMsg(UserIndex, 356)
                
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoTerrenoParticula(52, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY, 0))
                
        UserList(UserIndex).flags.CasteandoPortal = False
        
        UserList(UserIndex).flags.DondeTiroMap = 0
        UserList(UserIndex).flags.DondeTiroX = 0
        UserList(UserIndex).flags.DondeTiroY = 0
        UserList(UserIndex).Counters.TimeTeleport = 0
        UserList(UserIndex).Counters.CreoTeleport = False
        Debug.Print 1
    End If
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.description, "General.ControlarPortalLum", Erl)
    Resume Next

End Sub
Public Function FuncionCR() 'Cuenta regresiva

    If CuentaRegresivaTimer > 0 Then
    
        If CuentaRegresivaTimer > 1 Then
            
            If mapasegundos > 0 Then
                Call SendData(SendTarget.toMap, mapasegundos, PrepareMessageLocaleMsg(434, CuentaRegresivaTimer - 1))
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(434, CuentaRegresivaTimer - 1))
            End If
            
            
        Else
        
            If mapasegundos > 0 Then
                Call SendData(SendTarget.toMap, mapasegundos, PrepareMessageLocaleMsg(435))
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(435))
            End If
            
        mapasegundos = 0
       
        End If
        
        
    CuentaRegresivaTimer = CuentaRegresivaTimer - 1
    
    End If
End Function
Public Function DesconectarCuenta()
    
    On Error GoTo ErrorHandler
    
    Dim sArchivo As String, Cuenta As String
    Dim lngPos As Integer
    
    sArchivo = dir(App.Path & "\Cuentas" & "\*.cnt")
    Do While sArchivo <> vbNullString
        lngPos = (InStr(1, sArchivo, ".cnt") - 1)
        Cuenta = mid(sArchivo, 1, lngPos)
        Call WriteVar(App.Path & "\Cuentas\" & Cuenta & ".cnt", UCase$(Cuenta), "Conectada", "0")
        sArchivo = dir
    Loop

    Exit Function
ErrorHandler:
    Call RegistrarError(Err.Number, Err.description, "General.DesconectarCuenta", Erl)
    Resume Next
End Function

Function LoadObjNacimiento()

    On Error GoTo ErrorHandler
    
      If frmMain.Visible Then frmMain.AgregarConsola "Cargando Objetos iniciales al nacer..."
      
10    Dim Archivo As clsIniManager
    
20    Set Archivo = New clsIniManager
30    Call Archivo.Initialize(DatPath & "NACIMIENTO" & ".ini")
    
40    Dim i As Integer
50    Dim j As Integer
    
60    Dim CantidadObj As Integer
    
62    For i = 1 To NUMCLASES
64        CantidadObj = val(Archivo.GetValue(i, "CantidadObj"))
        
          ReDim OBJNacimiento.Clase(i).ObjIndex(1 To CantidadObj)
          ReDim OBJNacimiento.Clase(i).Amount(1 To CantidadObj)
          ReDim OBJNacimiento.Clase(i).Equipped(1 To CantidadObj)
          
66        For j = 1 To CantidadObj
67            OBJNacimiento.Clase(i).ObjIndex(j) = val(Archivo.GetValue(i, "ObjIndex" & j))
68            OBJNacimiento.Clase(i).Amount(j) = val(Archivo.GetValue(i, "ObjAmount" & j))
              OBJNacimiento.Clase(i).Equipped(j) = val(Archivo.GetValue(i, "ObjEquipped" & j))
69        Next j
        
70    Next i
        
71    Set Archivo = Nothing

      If frmMain.Visible Then frmMain.AgregarConsola "Se cargo Nacimiento.ini con éxito."
      
72    Exit Function
    
ErrorHandler:
    Set Archivo = Nothing
    Call RegistrarError(Err.Number, Err.description, "General.LoadObjNacimiento", Erl)
    Resume Next
End Function

