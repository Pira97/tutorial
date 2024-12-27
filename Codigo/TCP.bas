Attribute VB_Name = "TCP"
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

Enum lStat
    Incinerado = &H1
    Envenenado = &H2
    Comerciand = &H4
    Trabajando = &H8
    Combatiendo = &H10
    Ciego = &H20
    Inactivo = &H40
    Resucitando = &H80
    Saliendo = &H100
End Enum

Enum lStatEx
    Paralizado = &H1
    Inmovilizado = &H2
    Hombre = &H4
    Mujer = &H8
End Enum
 
Sub DarCuerpo(ByVal UserIndex As Integer)

    On Error GoTo Error_Err
    
    Dim NewBody As Integer
    Dim UserRaza As Byte
    Dim UserGenero As Byte
    
    UserGenero = UserList(UserIndex).Genero
    UserRaza = UserList(UserIndex).raza
    
 
    Select Case UserGenero
    
        Case eGenero.Hombre
        
            Select Case UserRaza
            
                Case eRaza.Humano
                    NewBody = 1
                    
                Case eRaza.Elfo
                    NewBody = 2
                    
                Case eRaza.Drow
                    NewBody = 3
                    
                Case eRaza.enano
                    NewBody = 52
                    
                Case eRaza.gnomo
                    NewBody = 52
                    
                Case eRaza.Orco
                    NewBody = 252
                    
                End Select
                
        Case eGenero.Mujer
        
            Select Case UserRaza
            
                Case eRaza.Humano
                    NewBody = 1
                    
                Case eRaza.Elfo
                    NewBody = 2
                    
                Case eRaza.Drow
                    NewBody = 3
                    
                Case eRaza.gnomo
                    NewBody = 138
                    
                Case eRaza.enano
                    NewBody = 138
                    
                Case eRaza.Orco
                    NewBody = 253
                    
                End Select
                
    End Select
        
    UserList(UserIndex).Char.body = NewBody
    
    Exit Sub

Error_Err:
    Call RegistrarError(Err.Number, Err.description, "TCP.DarCuerpo", Erl)
    Resume Next
    
End Sub

Private Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserGenero As Byte, ByVal Head As Integer) As Boolean

    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= 1 And Head <= 30)
                    
                Case eRaza.enano
                    ValidarCabeza = (Head >= 301 And Head <= 315)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= 101 And Head <= 121)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= 202 And Head <= 212)
 
                Case eRaza.gnomo
                    ValidarCabeza = (Head >= 401 And Head <= 409)
                    
                Case eRaza.Orco
                    ValidarCabeza = (Head >= 501 And Head <= 514)

            End Select
    
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= 70 And Head <= 80)

                Case eRaza.enano
                    ValidarCabeza = (Head >= 370 And Head <= 373)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= 170 And Head <= 189)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= 270 And Head <= 278)
 
                Case eRaza.gnomo
                    ValidarCabeza = (Head >= 470 And Head <= 481)
                
                Case eRaza.Orco
                    ValidarCabeza = (Head >= 570 And Head <= 573)

            End Select

    End Select
        
End Function

Function ValidarNombre(Nombre As String) As Boolean
    
    If Len(Nombre) < 1 Or Len(Nombre) > 18 Then Exit Function
    
    Dim temp As String
    temp = UCase$(Nombre)
    
    Dim i As Long, Char As Integer, LastChar As Integer
    For i = 1 To Len(temp)
        Char = Asc(mid$(temp, i, 1))
        
        If (Char < 65 Or Char > 90) And Char <> 32 Then
            Exit Function
        
        ElseIf Char = 32 And LastChar = 32 Then
            Exit Function
        End If
        
        LastChar = Char
    Next

    If Asc(mid$(temp, 1, 1)) = 32 Or Asc(mid$(temp, Len(temp), 1)) = 32 Then
        Exit Function
    End If
    
    ValidarNombre = True

End Function

Function AsciiValidos(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte
    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
    
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function

        End If
    
    Next i

    AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte
    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
    
        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function

        End If
    
    Next i

    Numeric = True

End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    For i = 1 To UBound(ForbidenNames)

        If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function

        End If

    Next i

    NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Integer

    For loopc = 1 To NUMSKILLS

        If UserList(UserIndex).Stats.UserSkills(loopc) < 0 Then
            Exit Function

            If UserList(UserIndex).Stats.UserSkills(loopc) > 100 Then UserList(UserIndex).Stats.UserSkills(loopc) = 100

        End If

    Next loopc

    ValidateSkills = True
    
End Function
Function AsignarAtributos(ByVal UserIndex As Integer)

    On Error GoTo Error_Err
    
1    Dim Fuerza As Byte, Agilidad As Byte, Inteligencia As Byte, Carisma As Byte, Constitucion As Byte
    
2    With UserList(UserIndex)
    
3        Select Case .Clase
        
            Case eClass.Clerigo, eClass.Asesino, eClass.Bardo, eClass.Druida, eClass.Paladin, eClass.Nigromante
4              Fuerza = 14
5              Agilidad = 14
6              Inteligencia = 18
7              Carisma = 10
8              Constitucion = 18
            
            Case eClass.Mago
9              Fuerza = 6
11             Agilidad = 18
10             Inteligencia = 18
12             Carisma = 10
13             Constitucion = 18
               
            Case Else 'Clases sin magia
14             Fuerza = 18
15             Agilidad = 18
16             Inteligencia = 6
17             Carisma = 10
18             Constitucion = 18
        
        End Select
        
19      .Stats.UserAtributos(eAtributos.Fuerza) = Fuerza + ModRaza(.raza).Fuerza
20      .Stats.UserAtributos(eAtributos.Agilidad) = Agilidad + ModRaza(.raza).Agilidad
21      .Stats.UserAtributos(eAtributos.Inteligencia) = IIf(Inteligencia + ModRaza(.raza).Inteligencia < 0, 0, Inteligencia + ModRaza(.raza).Inteligencia)
23      .Stats.UserAtributos(eAtributos.Carisma) = Carisma + ModRaza(.raza).Carisma
22      .Stats.UserAtributos(eAtributos.Constitucion) = Constitucion + ModRaza(.raza).Constitucion
        
61      .Stats.UserAtributosBackUP(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza)
62      .Stats.UserAtributosBackUP(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad)
63      .Stats.UserAtributosBackUP(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia)
64      .Stats.UserAtributosBackUP(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma)
65      .Stats.UserAtributosBackUP(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion)
    
    End With
    
    Exit Function

Error_Err:
    Call RegistrarError(Err.Number, Err.description, "TCP.AsignarAtributos", Erl)
    Resume Next
End Function
Function ConnectNewUser(ByVal UserIndex As Integer, ByRef Name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Cabeza As Integer, ByVal UserCuenta As String, ByVal Hogar As eCiudad, ByRef MensajeAdvertencia As Integer, ByRef CierraConexion As Boolean) As Boolean
                   
    On Error GoTo ConnectNewuser_Err
     
     CierraConexion = True
     ConnectNewUser = False
     MensajeAdvertencia = 36
     
1    Dim i As Long
2    Dim loopc As Long
3    Dim totalskpts As Long

4    With UserList(UserIndex)
       
        If ObtenerCantidadDePersonajesByUserIndex(UserIndex) >= 10 Then
            MensajeAdvertencia = 10    'No puedes crear mas personajes, has llegado a tu límite de diez personajes.
            CierraConexion = True
            Exit Function
        End If

5        If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
6            MensajeAdvertencia = (4)  ' Nombre inválido
             CierraConexion = False
7            Exit Function
8        End If
    
        If Trim(Name) = "" Then
            Name = RTrim$(Name)
            MensajeAdvertencia = (4)  ' Nombre inválido
            CierraConexion = False
            Exit Function
        End If
    
        If Len(Name) < 2 Then
            MensajeAdvertencia = (34) 'Corto
            CierraConexion = False
            Exit Function
        End If
            
        If Len(Name) > 30 Then
            MensajeAdvertencia = (35)  'El nombre es muy largo.
            CierraConexion = False
            Exit Function
        End If
        
17       If UserList(UserIndex).flags.UserLogged Then
18            Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).ip)
              Call frmMain.AgregarConsola("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).ip)
              Call CloseSocketSL(UserIndex)
              Call Cerrar_Usuario(UserIndex)
              CierraConexion = True
20            Exit Function
21       End If
 
        ' Nombre válido
        If Not ValidarNombre(Name) Then
            MensajeAdvertencia = (4)  ' Nombre inválido
            CierraConexion = False
            Exit Function
        End If
        
112     If Not NombrePermitido(Name) Then
113         MensajeAdvertencia = (74)  'Prohibido
            CierraConexion = False
            Exit Function
        End If

        '¿Existe el personaje?
114     If PersonajeExiste(Name) Then
115         MensajeAdvertencia = (7)  'Ya existe
            CierraConexion = False
            Exit Function
        End If
        
        If UserRaza <= 0 Or UserRaza > NUMRAZAS Then
            MensajeAdvertencia = (75)  'Raza valida
            CierraConexion = True
            Exit Function
        End If
        
        If UserSexo < eGenero.Hombre Or UserSexo > eGenero.Mujer Then
            MensajeAdvertencia = (76)  'Genero valido
            CierraConexion = True
            Exit Function
        End If
        
        If UserClase <= 0 Or UserClase > NUMCLASES Then
            MensajeAdvertencia = (77)  'Clase valida
            CierraConexion = True
            Exit Function
        End If
        
        If Hogar <= 0 Or Hogar > 2 Then 'Ponemos 2, porque al nacimiento solo hay Nix o Illi
            MensajeAdvertencia = (78)  'Hogar valido
            CierraConexion = True
            Exit Function
        End If
        
        If Not ValidarCabeza(UserRaza, UserSexo, Cabeza) Then
            MensajeAdvertencia = (79) 'Cabeza valida
            CierraConexion = True
            Exit Function
        End If
    
        'Si pasó los chequeos generales hacemos los chequeos por cuenta, por si un vivo quiere logear cosas con datos invalidos :p
        If seguridad_clones_validar(UserList(UserIndex).ip) = False Then
            MensajeAdvertencia = (80) 'Estás intentando crear muchos personajes, vuelva a intentarlo más tarde.
            CierraConexion = True
            Exit Function
        End If
        
        'Flags en 0
        Call UsuarioNuevoFlags(UserIndex)
        
40      .Name = Name
41      .Clase = UserClase
42      .raza = UserRaza
43      .Genero = UserSexo

        Call AsignarAtributos(UserIndex)
          
        Dim j As Integer
          
        For j = 1 To NUMATRIBUTOS
            If .Stats.UserAtributos(j) = 0 Then
                MensajeAdvertencia = (80)  'Atributos invalidos
                CierraConexion = True
                Exit Function
            End If
        Next j
          
46      Call ResetFacciones(UserIndex)


44        If Hogar = cIlliandor Then
            .Faccion.Status = 3
          Else
            .Faccion.Status = 2
          End If
          
          .Hogar = cDungeonNewbie
          
          
          '.Hogar = cRinkel
          
          .Stats.SkillPts = 10
          .Char.heading = eHeading.SOUTH
          
          Call DarCuerpo(UserIndex)
          
          .Char.Head = Cabeza
          .OrigChar.Head = .Char.Head
          .OrigChar = .Char
          
          .Char.WeaponAnim = NingunArma
          .Char.ShieldAnim = NingunEscudo
          .Char.CascoAnim = NingunCasco

          Dim MiInt As Long
          
98        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)

99       .Stats.MaxHP = 15 + MiInt
100      .Stats.MinHP = 15 + MiInt

101       MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)

102       If MiInt = 1 Then MiInt = 2

103       .Stats.MaxSta = 20 * MiInt
104       .Stats.MinSta = 20 * MiInt

105       .Stats.MaxAGU = 100
106       .Stats.MinAGU = 100

107       .Stats.MaxHam = 100
108       .Stats.MinHam = 100
    
109       Select Case UserClase
    
            Case eClass.Mago
111             MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
                .Stats.UserHechizos(1) = 2
                .Stats.MaxMAN = MiInt
                .Stats.MinMAN = MiInt
                
                
                 
             
                
    
1114        Case eClass.Clerigo, eClass.Druida, eClass.Bardo, eClass.Asesino, eClass.Nigromante, eClass.Paladin
                .Stats.UserHechizos(1) = 2
                
                If Not eClass.Paladin Then
                    .Stats.MaxMAN = 50
116                 .Stats.MinMAN = 50
117             End If

            Case Else
            
118             .Stats.MinMAN = 0
119             .Stats.MaxMAN = 0
120
          End Select
          
          .Stats.MaxHIT = 2
126       .Stats.MinHIT = 1
    
127       .Stats.GLD = 0
128       .Donador.CreditoDonador = 0

129       .Stats.Exp = 0
130       .Stats.ELU = 300
131       .Stats.ELV = 1

          Call RellenarInventario(UserIndex)
          
          'Posicion de comienzo (Primera vez que logea)
          .Pos = tCiudades.DungeonNewbie
        
        
1232    Call SaveNewUser(UserIndex)

        Call AddUserInAccount(Name, UserCuenta)
        
1233    If Not ConnectNewUserOnline(UserIndex, Name, UserCuenta, CierraConexion) Then
            MensajeAdvertencia = (36)
            CierraConexion = True
            Exit Function
        End If
 
        ConnectNewUser = True
        
    End With

    Exit Function

ConnectNewuser_Err:
     
    CierraConexion = True
    ConnectNewUser = False
    MensajeAdvertencia = 36

    Call RegistrarError(Err.Number, Err.description, "TCP.ConnectNewUser", Erl)
    Resume Next
End Function
Sub CloseSocket(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
        
        Call aDos.RestarConexion(UserList(UserIndex).ip)

        If UserIndex = LastUser Then

            Do Until UserList(LastUser).flags.UserLogged
      
                LastUser = LastUser - 1

                If LastUser < 1 Then Exit Do
            Loop

        End If
    
        If UserList(UserIndex).ConnID <> -1 Then
            Call CloseSocketSL(UserIndex)

        End If
     
        'Es el mismo user al que está revisando el centinela??
        'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
        ' y lo podemos loguear
        If Centinela.RevisandoUserIndex = UserIndex Then Call modCentinela.CentinelaUserLogout
    
         'Empty buffer for reuse
        Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.length)
    
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseUser(UserIndex)

            If NumUsers > 0 Then NumUsers = NumUsers - 1
            Call MostrarNumUsers
            
        Else
            Call ResetUserSlot(UserIndex)

        End If
    
        UserList(UserIndex).ConnID = -1
        UserList(UserIndex).ConnIDValida = False
    
        Exit Sub

ErrHandler:

    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    Call ResetUserSlot(UserIndex)
    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)

End Sub

 
'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

 
        If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
            Call BorraSlotSock(UserList(UserIndex).ConnID)
            Call WSApiCloseSocket(UserList(UserIndex).ConnID)
            UserList(UserIndex).ConnIDValida = False

        End If

 

End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
 
Public Function EnviarDatosASlot(ByVal UserIndex As Integer, _
                                 ByRef Datos As String) As Long
    
       '***************************************************
       'Author: Unknown
       'Last Modification: 01/10/07
       'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
       'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
       '***************************************************
       
        On Error GoTo Err
    
        Dim Ret As Long
    
        Ret = WsApiEnviar(UserIndex, Datos)
    
        If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
            ' Close the socket avoiding any critical error
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            
        End If
        
        Exit Function
        
Err:

End Function

Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim X As Integer, Y As Integer

    For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    EstaPCarea = False

End Function

Function HayPCarea(Pos As WorldPos) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim X As Integer, Y As Integer

    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function

                End If

            End If

        Next X
    Next Y

    HayPCarea = False

End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim X As Integer, Y As Integer

    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    HayOBJarea = False

End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean

    ValidateChr = UserList(UserIndex).Char.Head <> 0 And UserList(UserIndex).Char.body <> 0 And ValidateSkills(UserIndex)

End Function

Function ConnectNewUserOnline(ByVal UserIndex As Integer, ByRef Name As String, ByRef Account As String, ByRef CierraConexion As Boolean) As Boolean
    
    On Error GoTo ErrHandler
    
100    ConnectNewUserOnline = False
    
102    Dim tStr As String
    
104    With UserList(UserIndex)
         
106       If .flags.UserLogged Then
108            Call LogCheating("El usuario " & .Name & " ha intentado logear a " & Name & " desde la IP " & .ip)
110            Call CloseSocketSL(UserIndex)
112            Call Cerrar_Usuario(UserIndex)
114            CierraConexion = True
116            Exit Function
118       End If
 
122       'Reseteamos los FLAGS
124       .flags.Escondido = 0
126       .flags.TargetNPC = 0
128       .flags.TargetNpcTipo = eNPCType.Comun
130       .flags.TargetObj = 0
132       .flags.TargetUser = 0
134       .Char.FX = 0
 
          'Reseteamos los privilegios
136       .flags.Privilegios = 0
138       .GuildIndex = 0

          'Cargamos los datos del personaje
139       Dim PersonajeConError As Boolean

140       If Not LoadUser(UserIndex, Name, PersonajeConError) Then
              Call WriteShowMessageBox(UserIndex, 83) 'Error en la lectura del personaje 138
142           CierraConexion = False
143           Exit Function
144       End If
        
146       .LogOnTime = Now

148       If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
150       If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
152       If .Invent.WeaponEqpSlot = 0 And .Invent.NudiEqpSlot = 0 And .Invent.AnilloEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        
154       Call UpdateUserInv(True, UserIndex, 0)
156       Call UpdateUserHechizos(True, UserIndex, 0)
            
          'Sin estupidez
166       Call WriteDumbNoMore(UserIndex)
        
170       If Not MapaValido(.Pos.Map) Then
172           .Pos.Map = Ciudades(eCiudad.cIntermundia).Map
174           .Pos.X = Ciudades(eCiudad.cIntermundia).X
176           .Pos.Y = Ciudades(eCiudad.cIntermundia).Y
178       End If
        
          'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
          'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
180       If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).npcindex <> 0 Then
182           Dim FoundPlace As Boolean
184           Dim esAgua     As Boolean
186           Dim tX         As Long
188           Dim tY         As Long
       
190           FoundPlace = False
192           esAgua = HayAgua(.Pos.Map, .Pos.X, .Pos.Y)
       
194           For tY = .Pos.Y - 1 To .Pos.Y + 1
196               For tX = .Pos.X - 1 To .Pos.X + 1

198                   If esAgua Then
                        
                       'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
200                       If LegalPos(.Pos.Map, tX, tY, True, False) Then
202                           FoundPlace = True
204                           Exit For
206                        End If

208                    Else

                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
210                        If LegalPos(.Pos.Map, tX, tY, False, True) Then
212                            FoundPlace = True
214                            Exit For
216                        End If

218                    End If

220               Next tX

222             If FoundPlace Then Exit For

224           Next tY
       
226           If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
228               .Pos.X = tX
230               .Pos.Y = tY
232           Else

              'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
234              If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then
236                  Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
238               End If

240           End If

242       End If
        
          'Nombre de sistema
244       .Name = Name
246       .Account = Account
 
248       .showName = True 'Por default los nombres son visibles

          'Info
276       Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
278       Call WriteChangeMap(UserIndex, .Pos.Map)

282       .Counters.IdleCount = 0
    
          'Crea  el personaje del usuario
284       Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
   
286       Call WriteUserCharIndexInServer(UserIndex)
      
288       Call WriteUpdateUserStats(UserIndex)

296       Call SendMOTD(UserIndex)
        
298       Call SendUpdate(UserIndex)
 
308       .flags.UserLogged = True
        
310       MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
316       Call WriteLevelUp(UserIndex, .Stats.SkillPts)

332       .flags.SeguroResu = True

344       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
   
          'Aca las variables que afectan al login
356       If Lloviendo Then
358           Call WriteRainToggle(UserIndex, Queclima)
360       End If

374       Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(489, .Name)) 'Se ha conectado
         
386       NumUsers = NumUsers + 1
        
376       If NumUsers > recordusuarios Then
378           Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(6, NumUsers)) ' Record
380           recordusuarios = NumUsers
382           Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
384       End If
 
388       Call MostrarNumUsers

396       Call WriteLoggedMessage(UserIndex)
        
390       Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")
        
392       Call WriteVar(AccountPath & .Account & ".cnt", .Account, "Conectada", "1")
          
          ConnectNewUserOnline = True
          
       End With
 
    
    Exit Function

ErrHandler:
    ConnectNewUserOnline = False
    Call WriteShowMessageBox(UserIndex, 84) 'Error en la lectura del personaje 139
    Call RegistrarError(Err.Number, Err.description, "TCP.ConnectNewUserOnline", Erl)
    
End Function

Function ConnectUser(ByVal UserIndex As Integer, ByRef Name As String, ByRef Account As String, ByRef CierraConexion As Boolean) As Boolean
    
    On Error GoTo ErrHandler
    
100    ConnectUser = False
    
102    Dim tStr As String
    
104    With UserList(UserIndex)
         
106       If .flags.UserLogged Then
108            Call LogCheating("El usuario " & .Name & " ha intentado logear a " & Name & " desde la IP " & .ip)
110            Call CloseSocketSL(UserIndex)
112            Call Cerrar_Usuario(UserIndex)
114            CierraConexion = True
116            Exit Function
118       End If
 
122       'Reseteamos los FLAGS
124       .flags.Escondido = 0
126       .flags.TargetNPC = 0
128       .flags.TargetNpcTipo = eNPCType.Comun
130       .flags.TargetObj = 0
132       .flags.TargetUser = 0
134       .Char.FX = 0
 
          'Reseteamos los privilegios
136       .flags.Privilegios = 0
138       .GuildIndex = 0

          'Cargamos los datos del personaje
139       Dim PersonajeConError As Boolean

140       If Not LoadUser(UserIndex, Name, PersonajeConError) Then
              Call WriteShowMessageBox(UserIndex, 83) 'Error en la lectura del personaje 138
142           CierraConexion = False
143           Exit Function
144       End If
 
120       Call WriteLoggedSuccessful(UserIndex)
        
146       .LogOnTime = Now

148       If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
150       If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
152       If .Invent.WeaponEqpSlot = 0 And .Invent.NudiEqpSlot = 0 And .Invent.AnilloEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        
154       Call UpdateUserInv(True, UserIndex, 0)
156       Call UpdateUserHechizos(True, UserIndex, 0)
        
158       If .flags.Paralizado Then
160          Call WriteParalizeOK(UserIndex)
162       End If

164       If .flags.Estupidez = 0 Then
166           Call WriteDumbNoMore(UserIndex)
168       End If
        
170       If Not MapaValido(.Pos.Map) Then
172           .Pos.Map = Ciudades(eCiudad.cIntermundia).Map
174           .Pos.X = Ciudades(eCiudad.cIntermundia).X
176           .Pos.Y = Ciudades(eCiudad.cIntermundia).Y
178       End If

          'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
          'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
180       If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).npcindex <> 0 Then
182           Dim FoundPlace As Boolean
184           Dim esAgua     As Boolean
186           Dim tX         As Long
188           Dim tY         As Long
       
190           FoundPlace = False
192           esAgua = HayAgua(.Pos.Map, .Pos.X, .Pos.Y)
       
194           For tY = .Pos.Y - 1 To .Pos.Y + 1
196               For tX = .Pos.X - 1 To .Pos.X + 1

198                   If esAgua Then
                        
                       'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
200                       If LegalPos(.Pos.Map, tX, tY, True, False) Then
202                           FoundPlace = True
204                           Exit For
206                        End If

208                    Else

                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
210                        If LegalPos(.Pos.Map, tX, tY, False, True) Then
212                            FoundPlace = True
214                            Exit For
216                        End If

218                    End If

220               Next tX

222             If FoundPlace Then Exit For

224           Next tY
       
226           If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
228               .Pos.X = tX
230               .Pos.Y = tY
232           Else

              'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
234              If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then
236                  Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
238               End If

240           End If

242       End If
        
          'Nombre de sistema
244       .Name = Name
246       .Account = Account
 
248       .showName = True 'Por default los nombres son visibles
   
          'If in the water, and has a boat, equip it!
250       If .Invent.BarcoObjIndex > 0 And (HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then

252           Dim Barco As ObjData
254           Barco = ObjData(.Invent.BarcoObjIndex)

256           .Char.Head = 0

258           If .flags.Muerto <> 0 Then
260               .Char.body = iFragataFantasmal
262           Else
264               .Char.body = iBarca
266           End If

268           .flags.Navegando = 1
274       End If

          'Info
276       Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
278       Call WriteChangeMap(UserIndex, .Pos.Map)

282       .Counters.IdleCount = 0
    
          'Crea  el personaje del usuario
284       Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
   
286       Call WriteUserCharIndexInServer(UserIndex)
      
288       Call WriteUpdateUserStats(UserIndex)

296       Call SendMOTD(UserIndex)
        
298       Call SendUpdate(UserIndex)
 
308       .flags.UserLogged = True
        
310       MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
312       If .Stats.SkillPts > 0 Then
314           Call WriteSendSkills(UserIndex)
316           Call WriteLevelUp(UserIndex, .Stats.SkillPts)
318       End If
        
320       If .flags.Navegando = 1 Then
322           Call WriteNavigateToggle(UserIndex)
324       End If

326       If (.flags.Muerto = 0) Then
328           .flags.SeguroResu = False
330       Else
332           .flags.SeguroResu = True
334       End If
  
336       If .GuildIndex > 0 Then
           'welcome to the show baby...
338           If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
               'Call WriteMensajes(UserIndex, eMensajes.Mensaje387)
340           End If
342       End If
   
344       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
   
346       Call modGuilds.SendGuildNews(UserIndex)
   
348       tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)

350       If LenB(tStr) <> 0 Then
352           Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
354       End If

          'Aca las variables que afectan al login
356       If Lloviendo Then
358           Call WriteRainToggle(UserIndex, Queclima)
360       End If
              
362       If .flags.RecibioCorreo > 0 Then
364           WriteMensajeSigno UserIndex, 1
366       End If
 
368       If .Counters.Pena > 0 Then
370           Call WriteLocaleMsg(UserIndex, 146, .Counters.Pena)
372       End If
        
374       Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(489, .Name))
         
386       NumUsers = NumUsers + 1
        
376       If NumUsers > recordusuarios Then
378           Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(6, NumUsers))
380           recordusuarios = NumUsers
382           Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
384       End If

388       Call MostrarNumUsers
        
390       Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")
        
392       Call WriteVar(AccountPath & .Account & ".cnt", .Account, "Conectada", "1")

396       Call WriteLoggedMessage(UserIndex)

          'Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
394       Call IntervaloPermiteSerAtacado(UserIndex, True)

          ConnectUser = True
          
       End With
 
    
    Exit Function

ErrHandler:
    ConnectUser = False
    Call WriteShowMessageBox(UserIndex, 84) 'Error en la lectura del personaje 139
    Call RegistrarError(Err.Number, Err.description, "TCP.ConnectUser", Erl)
    
End Function



Sub ResetFacciones(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    With UserList(UserIndex).Faccion
        .Status = 0

       .CiudadanosMatados = 0
       .RenegadosMatados = 0
       .RepublicanosMatados = 0
        
       .Rango = 0
    End With

End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '05/20/2007 Integer - Agregue todas las variables que faltaban.
    '*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0

    End With

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .Loops = 0
        .heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .ParticulaFx = 0
        .Arma_Aura = 0
        .Body_Aura = 0
        .Escudo_Aura = 0
        .Head_Aura = 0
        .Anillo_Aura = 0
        .Otra_Aura = 0
    End With

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex)
        .Name = vbNullString
        .desc = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .Clase = 0
        .Genero = 0
        .Hogar = 0
        .raza = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0

        End With
        
    End With

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0

    End If

    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)

    End If

    UserList(UserIndex).GuildIndex = 0

End Sub

Sub UsuarioNuevoFlags(ByVal UserIndex As Integer)
Dim i As Long
    With UserList(UserIndex)

26        .flags.RecibioCorreo = 0
27        .flags.CantidadAmigos = 0
28        .flags.Muerto = 0
29        .flags.Escondido = 0
36        .flags.CantidadCorreos = 0
37        .Donador.activo = 0

        
For i = 1 To MAXAMIGOS
  .Amigos(i).Nombre = "Vacío"
  .Amigos(i).index = 0
Next i


End With
 With UserList(UserIndex)
        .Casamiento.Candidato = 0
        .Casamiento.Casado = 0
        .Casamiento.Pareja = ""
        
    End With
    
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 06/28/2008
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '03/29/2006 Maraxus - Reseteo el CentinelaOK también.
    '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
    '*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Montando = 0
        .Oculto = 0
        .Envenenado = 0
        .Incinerado = 0
        .Invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Meditando = 0
        .Trabajando = 0
        .Lingoteando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .OwnedNpc = 0
        .Ignorado = False
        .ParalizedBy = vbNullString
        .ParalizedByIndex = 0
        .ParalizedByNpcIndex = 0
        .CasteandoPortal = False
        .Resucitando = False
    End With

    'Resto de "flags"
    With UserList(UserIndex)
    
        .Casamiento.Candidato = 0
        .Casamiento.Casado = 0
        .Casamiento.Pareja = ""
        
    End With
    
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long

    For loopc = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(loopc) = 0
    Next loopc

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For loopc = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(loopc) = 0
        UserList(UserIndex).MascotasType(loopc) = 0
    Next loopc

End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long
    
    For loopc = 1 To MAX_BANCOINVENTORY_SLOTS
        UserList(UserIndex).BancoInvent.Object(loopc).Amount = 0
        UserList(UserIndex).BancoInvent.Object(loopc).Equipped = 0
        UserList(UserIndex).BancoInvent.Object(loopc).ObjIndex = 0
    Next loopc
    
    UserList(UserIndex).BancoInvent.NroItems = 0

End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Long

    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).ConnID = -1
    
    Call ResetFacciones(UserIndex)
    Call ResetContadores(UserIndex)
    Call ResetGuildInfo(UserIndex)
    Call ResetCharInfo(UserIndex)
    Call ResetBasicUserInfo(UserIndex)
    Call ResetUserFlags(UserIndex)
    Call LimpiarInventario(UserIndex)
    Call ResetUserSpells(UserIndex)
    Call ResetUserPets(UserIndex)
    Call ResetUserBanco(UserIndex)
    Call ResetUserExtras(UserIndex)

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    Dim errordesc As String
    Dim Map   As Integer
    Dim i     As Integer
    Dim aN    As Integer
    
    Map = UserList(UserIndex).Pos.Map
    

    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Attackable = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString
    End If

    aN = UserList(UserIndex).flags.NPCAtacado

    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
        
    End If

    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.NPCAtacado = 0
    
    
    errordesc = "ERROR AL DESMONTAR"
    
    If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MonturaSlot)
    End If

    errordesc = "ERROR AL SACAR MIMETISMO"

    If UserList(UserIndex).Invent.MagicIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MagicSlot)
    End If

    errordesc = "ERROR AL ENVIAR PARTICULA O FX"
    
    UserList(UserIndex).Char.FX = 0
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, 0, 0, True, True))
    
    
    UserList(UserIndex).flags.UserLogged = False
    UserList(UserIndex).Counters.Saliendo = False
        
    errordesc = "ERROR AL ENVIAR INVI"
    
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

    errordesc = "ERROR AL CERRAR PORTAL TP"
    
    Call ControlarPortalLum(UserIndex, 1)
    
    errordesc = "ERROR AL CERRAR AMIGOS"
    Call ObtenerIndexAmigos(UserIndex, True)
          
    errordesc = "ERROR AL GRABAR PJ Y CERRAR CUENTA"
    
    Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Logged", "0")
    Call WriteVar(AccountPath & UserList(UserIndex).Account & ".cnt", UserList(UserIndex).Account, "Conectada", "0")
    
    Call SaveUser(UserIndex, True)

    errordesc = "ERROR AL DESCONTAR USER DE MAPA"
  
    If MapInfo(Map).NumUsers > 0 Then
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    End If
 
    errordesc = "ERROR AL ERASEUSERCHAR"

    If UserList(UserIndex).Char.CharIndex > 0 Then
        Call EraseUserChar(UserIndex, UserList(UserIndex).flags.AdminInvisible = 1, False)
    End If
    
    errordesc = "ERROR Update Map Users"
    
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0
    
    errordesc = "ERROR AL LIMPIAR CONSULTAS GM"
    
    'Reseteamos consultas
    If Ayuda.Existe(UserList(UserIndex).Name) Then
        Dim UserFile As String
        
        UserFile = DocConsultas & UserList(UserIndex).Name & ".ini"
        
        Call Ayuda.Quitar(UserList(UserIndex).Name)
        
        Dim NumConsultas As Integer
        Dim Consulta As String
        Dim Reemplazo As String
        
        NumConsultas = val(GetVar(UserFile, "INIT", "NumConsultas"))
        Consulta = GetVar(UserFile, "CONSULTAS", "C" & NumConsultas)
        Reemplazo = "1"
        
        Consulta = StrReverse(mid(StrReverse(Consulta), Len(Reemplazo) + 1)) + Reemplazo

        Call WriteVar(UserFile, "CONSULTAS", "C" & NumConsultas, Consulta)

    End If
    
    errordesc = "ERROR AL RESETEAR FLAGS Name:" & UserList(UserIndex).Name & " cuenta:" & UserList(UserIndex).Account
    
    Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(490, UserList(UserIndex).Name)) 'Se ha desconectado
    
    Call ResetUserSlot(UserIndex)
    
    Exit Sub

ErrHandler:
    Debug.Print errordesc & " " & Time
    Call RegistrarError(Err.Number, Err.description, "TCP.CloseUser", Erl)
    Resume Next ' TODO: Provisional hasta solucionar bugs graves
    
End Sub

Sub ReloadSokcet()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

        Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
        If NumUsers <= 0 Then
            Call WSApiReiniciarSockets
        Else

            '       Call apiclosesocket(SockListen)
            '       SockListen = ListenForConnect(Puerto, hWndMsg, "")
        End If


    Exit Sub
ErrHandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long
    
    For loopc = 1 To LastUser

        If UserList(loopc).flags.UserLogged And UserList(loopc).ConnID >= 0 And UserList(loopc).ConnIDValida Then
            If UserList(loopc).flags.Privilegios And PlayerType.User Then
                Call CloseSocket(loopc)

            End If

        End If

    Next loopc

End Sub
 
Public Sub ResetUserExtras(ByVal UserIndex As Integer)
'***************************************************
'Author: Bateman
'***************************************************
  Dim i As Integer
  For i = 1 To MAXAMIGOS
  UserList(UserIndex).Amigos(i).Nombre = vbNullString
  UserList(UserIndex).Amigos(i).index = 0
  Next i
UserList(UserIndex).QuienAmigo = vbNullString
End Sub
Function Generate_Char_Stat(ByVal UserIndex As Integer) As Integer

    With UserList(UserIndex)
    
        If .flags.Incinerado = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Incinerado
        End If
        
        If .flags.Envenenado > 0 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Envenenado
        End If
    
        If .flags.Comerciando = True Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Comerciand
        End If

        If .flags.Trabajando = True Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Trabajando
        End If

        If .Counters.TiempoDeMapeo > 0 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Combatiendo
        End If
        
        If .flags.Ceguera = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Ciego
        End If
    
        If .Counters.IdleCount > 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Inactivo
        End If
        
        If .flags.Resucitando = True Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Resucitando
        End If
        
        If .Counters.Saliendo = True Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Saliendo
        End If
        
    End With
    
End Function
 
Function Generate_Char_StatEx(ByVal UserIndex As Integer) As Byte

With UserList(UserIndex)
    If .flags.Paralizado = 1 Then
       Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Paralizado
    End If

    If .flags.Inmovilizado = 1 Then
        Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Inmovilizado
    End If
    
    If .Genero = eGenero.Hombre Then
        Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Hombre
    Else
        Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Mujer
    End If
End With
End Function
Function Generate_Char_StatExNpcs(ByVal npcindex As Integer) As Byte

With Npclist(npcindex)
    If .flags.Paralizado = 1 Then
       Generate_Char_StatExNpcs = Generate_Char_StatExNpcs Or lStatEx.Paralizado
    End If

    If .flags.Inmovilizado = 1 Then
        Generate_Char_StatExNpcs = Generate_Char_StatExNpcs Or lStatEx.Inmovilizado
    End If

End With
End Function
Function LoadUser(UserIndex As Integer, Name As String, ByRef Error As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    'Cargamos el personaje
    
1    Dim UserFile As New clsIniManager
2    Set UserFile = New clsIniManager

3    Call UserFile.Initialize(CharPath & UCase$(Name) & ".chr")
    
4    Dim loopc As Long
5    Dim ln As String

     LoadUser = False
     
     With UserList(UserIndex)
    
        'Load Counters
6        With .Counters
9            .Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
7            .TiempoDeMapeo = 0
8        End With
        
         'Load Flags
10       Dim priv As Long 'Para muchas verificaciones el compilador de vb usa long para tener mas eficencia
        
        '[DONADOR]
        .Donador.activo = CByte(UserFile.GetValue("DONADOR", "DONADOR"))
        .Donador.CreditoDonador = CLng(UserFile.GetValue("DONADOR", "PUNTOS"))
        
22       If val(GetVar(AccountPath & UserList(UserIndex).Account & ".cnt", UserList(UserIndex).Account, "Donador")) Then .Donador.activo = 1

11        With .flags
12            .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
13            .Montando = CByte(UserFile.GetValue("FLAGS", "Montando"))
14            .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
15            .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
16            .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
17            .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
18            .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
19            .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
21            .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            
24            .CheckAmigos = 0
        
26            .CantidadAmigos = CByte(UserFile.GetValue("FLAGS", "CantidadAmigos"))
            
25            .CantidadCorreos = CByte(UserFile.GetValue("FLAGS", "Correos"))
             
27            .Incinerado = CByte(UserFile.GetValue("FLAGS", "Incinerado"))
28
29            .MuertesUsuario = CLng(UserFile.GetValue("FLAGS", "Murio"))
30            .RecibioCorreo = CByte(UserFile.GetValue("FLAGS", "Recibiocorreo"))
            
        End With
        
          '[CASAMIENTO]
          .Casamiento.Casado = CByte(UserFile.GetValue("CASAMIENTO", "Casado"))
          .Casamiento.Pareja = CStr(UserFile.GetValue("CASAMIENTO", "Pareja"))
        
        'Cargamos datos faccionarios
31        With .Faccion
32           priv = CLng(UserFile.GetValue("FACCIONES", "Status"))
33          .CiudadanosMatados = CInt(UserFile.GetValue("FACCIONES", "CiudMatados"))
34          .RenegadosMatados = CInt(UserFile.GetValue("FACCIONES", "ReneMatados"))
35          .RepublicanosMatados = CInt(UserFile.GetValue("FACCIONES", "RepuMatados"))
36          .Rango = CInt(UserFile.GetValue("FACCIONES", "RANGO"))
37          .CaosMatados = CInt(UserFile.GetValue("FACCIONES", "CaosMatados"))
38          .ArmadaMatados = CInt(UserFile.GetValue("FACCIONES", "ArmiMatados"))
39          .MilicianosMatados = CInt(UserFile.GetValue("FACCIONES", "MiliMatados"))
        End With
        
        
40        Call LogGM(Name, "Se conecto con ip:" & .ip)
 
41        If EsAdmin(Name) Then
42          .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
43          .Faccion.Status = 9
          ElseIf EsDios(Name) Then
44          .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
45          .Faccion.Status = 9
          ElseIf EsSemiDios(Name) Then
46          .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
            .Faccion.Status = 8
          ElseIf EsConsejero(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            .Faccion.Status = 7
          ElseIf EsRolesMaster(Name) Then
49          .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
50          .Faccion.Status = 7
          Else
51            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
52            .flags.AdminPerseguible = True
53            .Faccion.Status = priv
          
          End If
        
54        If .flags.Paralizado = 1 Then .Counters.Paralisis = IntervaloParalizado
        
        'Datos del personaje
        '''///
        
55        .Genero = CByte(UserFile.GetValue("INIT", "Genero"))
56        .Clase = CByte(UserFile.GetValue("INIT", "Clase"))
58        .raza = CByte(UserFile.GetValue("INIT", "Raza"))
57        .Hogar = CByte(UserFile.GetValue("INIT", "Hogar"))
59        .desc = CStr(UserFile.GetValue("INIT", "Desc"))
60        .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
62        .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
63        .Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
64        .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
65        .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
        With .OrigChar
66            .Head = CInt(UserFile.GetValue("INIT", "Head"))
68            .body = CInt(UserFile.GetValue("INIT", "Body"))
67            .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
69            .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
70            .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
71            .heading = UserList(UserIndex).Char.heading
        End With
        
72        If .flags.Muerto = 0 Then
73            .Char = .OrigChar
          Else
74            .Char.body = iCuerpoMuerto
75            .Char.Head = iCabezaMuerto
76            .Char.WeaponAnim = NingunArma
79            .Char.ShieldAnim = NingunEscudo
78            .Char.CascoAnim = NingunCasco
77        End If
        
 
80        .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
81        .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
82        For loopc = 1 To MAX_BANCOINVENTORY_SLOTS
92            ln = UserFile.GetValue("BancoInventory", "Obj" & loopc)
91            .BancoInvent.Object(loopc).ObjIndex = CInt(ReadField(1, ln, 45))
90            .BancoInvent.Object(loopc).Amount = CInt(ReadField(2, ln, 45))
89        Next loopc


        'Lista de objetos
83        For loopc = 1 To MAX_INVENTORY_SLOTS
85            ln = UserFile.GetValue("Inventory", "Obj" & loopc)
86            .Invent.Object(loopc).ObjIndex = CInt(ReadField(1, ln, 45))
87            .Invent.Object(loopc).Amount = CInt(ReadField(2, ln, 45))
88            .Invent.Object(loopc).Equipped = CByte(ReadField(3, ln, 45))

84        Next loopc

        With .Invent
        
        'Obtiene el indice-objeto del arma
93        .NudiEqpSlot = CByte(UserFile.GetValue("Inventory", "NudiEqpSlot"))
        
94        If .NudiEqpSlot > 0 Then
        
95            .NudiEqpObjIndex = .Object(.NudiEqpSlot).ObjIndex
        
              
96            If UserList(UserIndex).flags.Muerto = 0 Then
97                UserList(UserIndex).Char.Arma_Aura = ObjData(.NudiEqpObjIndex).Aura
98            End If
            
99        End If
        
        'Obtiene el indice-objeto del arma
100        .WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))

101        If .WeaponEqpSlot > 0 Then
102            .WeaponEqpObjIndex = .Object(.WeaponEqpSlot).ObjIndex
              
103            If UserList(UserIndex).flags.Muerto = 0 Then
104                UserList(UserIndex).Char.Arma_Aura = ObjData(.WeaponEqpObjIndex).Aura
105            End If
            
106        End If
 
        'Obtiene el indice-objeto del escudo
111        .EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))

113        If .EscudoEqpSlot > 0 Then
112            .EscudoEqpObjIndex = .Object(.EscudoEqpSlot).ObjIndex
        
114            If UserList(UserIndex).flags.Muerto = 0 Then
116                UserList(UserIndex).Char.Escudo_Aura = ObjData(.EscudoEqpObjIndex).Aura
115            End If
            
117        End If

        'Obtiene el indice-objeto del anillo magico
118        .MagicSlot = CByte(UserFile.GetValue("Inventory", "MagicSlot"))

119        If .MagicSlot > 0 Then
120            .MagicIndex = .Object(.MagicSlot).ObjIndex
        
121            If UserList(UserIndex).flags.Muerto = 0 Then
122                UserList(UserIndex).Char.Anillo_Aura = ObjData(.MagicIndex).Aura
123            End If
            
124        End If
        
        'Obtiene el indice-objeto del casco
125        .CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))

126        If .CascoEqpSlot > 0 Then
127            .CascoEqpObjIndex = .Object(.CascoEqpSlot).ObjIndex
        
128            If UserList(UserIndex).flags.Muerto = 0 Then
129                UserList(UserIndex).Char.Head_Aura = ObjData(.CascoEqpObjIndex).Aura
130            End If
            
131        End If
        
        'Obtiene el indice-objeto barco
132        .BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))

133        If .BarcoSlot > 0 Then .BarcoObjIndex = .Object(.BarcoSlot).ObjIndex
 
 
        'Obtiene el indice-objeto municion
134        .MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))

135        If .MunicionEqpSlot > 0 Then .MunicionEqpObjIndex = .Object(.MunicionEqpSlot).ObjIndex

        
        'Obtiene el indice-objeto anilo
136        .AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))

137        If .AnilloEqpSlot > 0 Then .AnilloEqpObjIndex = .Object(.AnilloEqpSlot).ObjIndex
            
           
            
        'Obtiene el indice-objeto del armadura
138        .ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))

139        If .ArmourEqpSlot > 0 Then
140            .ArmourEqpObjIndex = .Object(.ArmourEqpSlot).ObjIndex
141            UserList(UserIndex).flags.Desnudo = 0
            
142            If UserList(UserIndex).flags.Muerto = 0 Then
143                UserList(UserIndex).Char.Body_Aura = ObjData(.ArmourEqpObjIndex).Aura
144            End If
            
145        Else
146            UserList(UserIndex).flags.Desnudo = 1
147        End If
        
  
        'Obtiene el indice-objeto montura
107        .MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
        
108        If .MonturaSlot > 0 Then
109            .MonturaObjIndex = .Object(.MonturaSlot).ObjIndex
110        End If
        
        
        End With
        
148        ln = UserFile.GetValue("Guild", "GUILDINDEX")

        If IsNumeric(ln) Then
149            .GuildIndex = CInt(ln)
        Else
150            .GuildIndex = 0
        End If
    
    
151    For loopc = 1 To Max_Correos
              With .Correos(loopc)
153                    .Carta = CStr(UserFile.GetValue("CORREO", "Carta" & loopc))
154                    .Emisor = CStr(UserFile.GetValue("CORREO", "Emisor" & loopc))
155                    .Leida = CByte(UserFile.GetValue("CORREO", "Leida" & loopc))
156                    .ObjetoIndex = CInt(ReadField(1, UserFile.GetValue("CORREO", "Objeto" & loopc), 45))
                       .ObjetoCantidad = CInt(ReadField(2, UserFile.GetValue("CORREO", "Objeto" & loopc), 45))

              End With
        Next loopc
            
152        For loopc = 1 To MAXAMIGOS
            .Amigos(loopc).Nombre = CStr(UserFile.GetValue("AMIGOS", "NOMBRE" & loopc))
501           Next loopc
        
        
        'Comienza LoadUserStats
        With .Stats

602            For loopc = 1 To NUMATRIBUTOS
603                .UserAtributos(loopc) = CByte(UserFile.GetValue("ATRIBUTOS", "AT" & loopc))
604                .UserAtributosBackUP(loopc) = CByte(.UserAtributos(loopc))
605            Next loopc
            
606            For loopc = 1 To NUMSKILLS
607                .UserSkills(loopc) = CByte(UserFile.GetValue("SKILLS", "SK" & loopc))
608                .EluSkills(loopc) = CLng(UserFile.GetValue("SKILLS", "ELUSK" & loopc))
609                .ExpSkills(loopc) = CLng(UserFile.GetValue("SKILLS", "EXPSK" & loopc))
            Next loopc
        
            For loopc = 1 To MAXUSERHECHIZOS
610                .UserHechizos(loopc) = CInt(UserFile.GetValue("Hechizos", "H" & loopc))
            Next loopc
        
611            .GLD = CLng(UserFile.GetValue("STATS", "GLD"))
612            .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
613            .MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
614            .MinHP = CInt(UserFile.GetValue("STATS", "MinHP"))
        
615            .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
616            .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
617            .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
618            .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
619            .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
620            .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
621            .MaxAGU = CInt(UserFile.GetValue("STATS", "MaxAGU"))
622            .MinAGU = CInt(UserFile.GetValue("STATS", "MinAGU"))
        
623            .MaxHam = CInt(UserFile.GetValue("STATS", "MaxHAM"))
624            .MinHam = CInt(UserFile.GetValue("STATS", "MinHAM"))
        
625            .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

626            .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
627            .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
628            .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
            
629            .UsuariosMatados = CInt(UserFile.GetValue("MUERTES", "UserMuertes"))
630            .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

        End With
        
    End With
    
631 Set UserFile = Nothing
    
632    If Not ValidateChr(UserIndex) Then
633        LoadUser = False
634        Exit Function
635    End If
    
636    LoadUser = True
    
    Exit Function

ErrorHandler:

    Set UserFile = Nothing
    LoadUser = False
    Error = True
    
    Call RegistrarError(Err.Number, Err.description, "TCP.LoadUser", Erl)
    
End Function

Sub RellenarInventario(ByVal UserIndex As String)
        
        On Error GoTo RellenarInventario_Err
        
          With UserList(UserIndex)
          
          Dim i As Integer
          
          Dim NumItems As Integer
          NumItems = UBound(OBJNacimiento.Clase(.Clase).ObjIndex())


          For i = 1 To NumItems
            
            Select Case i
            
                Case 4
            
                    Select Case .raza
                       Case eRaza.Elfo
                          .Invent.Object(i).ObjIndex = 464
                          
                       Case eRaza.Drow
                          .Invent.Object(i).ObjIndex = 465
                          
                       Case eRaza.enano
                          .Invent.Object(i).ObjIndex = 466
                          
                       Case eRaza.gnomo
                          .Invent.Object(i).ObjIndex = 466
                          
                       Case eRaza.Orco
                          .Invent.Object(i).ObjIndex = 1087
                       
                       Case Else 'Humano
                          .Invent.Object(i).ObjIndex = 463
                          
                    End Select
            
                    .Invent.ArmourEqpSlot = i
                    .Invent.ArmourEqpObjIndex = .Invent.Object(i).ObjIndex
                    .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                       
                Case 3
                    .Invent.Object(i).ObjIndex = OBJNacimiento.Clase(.Clase).ObjIndex(i)
                    .Invent.WeaponEqpSlot = i
                    .Invent.WeaponEqpObjIndex = .Invent.Object(i).ObjIndex
                    .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
                    
                Case Else
            
                    .Invent.Object(i).ObjIndex = OBJNacimiento.Clase(.Clase).ObjIndex(i)
                
                End Select
            
            .Invent.Object(i).Amount = OBJNacimiento.Clase(.Clase).Amount(i)
            .Invent.Object(i).Equipped = OBJNacimiento.Clase(.Clase).Equipped(i)
            
          Next i
          
         'Seteo la cantidad de items
208      .Invent.NroItems = NumItems

        End With
   
        
        Exit Sub

RellenarInventario_Err:
210     Call RegistrarError(Err.Number, Err.description, "TCP.RellenarInventario", Erl)
212     Resume Next
        
End Sub


Sub SaveNewUser(ByVal UserIndex As Integer)
        
        On Error GoTo SaveNewUser_Err

        Dim OldUserHead As Long
        Dim UserFile    As String
        
        With UserList(UserIndex)
        
100     UserFile = CharPath & UCase$(.Name) & ".chr"
        
        
        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
        
102     If .Clase = 0 Or .Stats.ELV = 0 Then
104         Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
            Exit Sub
        End If
    
106     If FileExist(UserFile, vbNormal) Then
108         If .flags.Muerto = 1 Then
110             OldUserHead = UserList(UserIndex).Char.Head
112              .Char.Head = GetVar(UserFile, "INIT", "Head")

            End If
            
        End If
    
        Dim loopc As Integer

        Dim n

        Dim Datos$

114     n = FreeFile

116     Open UserFile For Binary Access Write As n

        '[FLAGS]
122     Put n, , "[FLAGS]" & vbCrLf
126     Put n, , "Muerto=0" & vbCrLf
128     Put n, , "Escondido=0" & vbCrLf
130     Put n, , "Hambre=0" & vbCrLf
132     Put n, , "Sed=0" & vbCrLf
134     Put n, , "Desnudo=0" & vbCrLf
        Put n, , "Ban=0" & vbCrLf
136     Put n, , "Navegando=0" & vbCrLf
137     Put n, , "Montando=0" & vbCrLf
138     Put n, , "Envenenado=0" & vbCrLf
140     Put n, , "Paralizado=0" & vbCrLf
144     Put n, , "Incinerado=0" & vbCrLf
152     Put n, , "Correos=0" & vbCrLf
146     Put n, , "Murio=0" & vbCrLf
147     Put n, , "RecibioCorreo=0" & vbCrLf
        Put n, , "CantidadAmigos=0" & vbCrLf


        '[COUNTERS]
        Put n, , "[COUNTERS]" & vbCrLf
        Put n, , "Pena=0" & vbCrLf


        '[FACCIONES]
835     Put n, , "[FACCIONES]" & vbCrLf
        Put n, , "Rango=0" & vbCrLf
        Put n, , "Status=" & .Faccion.Status & vbCrLf
        
        Put n, , "CiudMatados=0" & vbCrLf
        Put n, , "ReneMatados=0" & vbCrLf
        Put n, , "RepuMatados=0" & vbCrLf
        Put n, , "CaosMatados=0" & vbCrLf
        Put n, , "ArmiMatados=0" & vbCrLf
        Put n, , "MiliMatados=0" & vbCrLf
        
        '[DONADOR]
834     Put n, , "[DONADOR]" & vbCrLf
        Put n, , "Donador=0" & vbCrLf
        Put n, , "Puntos=0" & vbCrLf
  
  
        '[ATRIBUTOS]
352     Put n, , "[ATRIBUTOS]" & vbCrLf

        '¿Fueron modificados los atributos del usuario?
354     For loopc = 1 To UBound(.Stats.UserAtributos)
356         Put n, , "AT" & loopc & "=" & CStr(.Stats.UserAtributos(loopc)) & vbCrLf
        Next

        '[SKILLS]
276     Put n, , "[SKILLS]" & vbCrLf

278     For loopc = 1 To UBound(.Stats.UserSkills)
280         Put n, , "SK" & loopc & "=0" & vbCrLf
            Put n, , "ELUSK" & loopc & "=0" & vbCrLf
            Put n, , "EXPSK" & loopc & "=0" & vbCrLf
            
        Next
  
        '[CASAMIENTO]
314     Put n, , "[CASAMIENTO]" & vbCrLf
        Put n, , "PAREJA=" & vbCrLf
        Put n, , "CASADO=0" & vbCrLf
        

        '[INIT]
        Put n, , "[INIT]" & vbCrLf

316     Put n, , "Genero=" & .Genero & vbCrLf
318     Put n, , "Raza=" & .raza & vbCrLf
320     Put n, , "Hogar=" & .Hogar & vbCrLf
322     Put n, , "Clase=" & .Clase & vbCrLf
324     Put n, , "Desc=" & .desc & vbCrLf
326     Put n, , "Heading=" & CStr(.Char.heading) & vbCrLf
328     Put n, , "Head=" & CStr(.Char.Head) & vbCrLf
        Put n, , "Body=" & CStr(.Char.body) & vbCrLf
330     Put n, , "Arma=" & CStr(.Char.WeaponAnim) & vbCrLf
332     Put n, , "Escudo=" & CStr(.Char.ShieldAnim) & vbCrLf
334     Put n, , "Casco=" & CStr(.Char.CascoAnim) & vbCrLf

        Dim TempDate As Date
        TempDate = Now - .LogOnTime
        .LogOnTime = Now
        .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        .UpTime = .UpTime
        
        Put n, , "UpTime=" & .UpTime & vbCrLf
        
336     Put n, , "Position=" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & vbCrLf

        '[STATS]
204     Put n, , "[STATS]" & vbCrLf
        Put n, , "GLD=" & CStr(.Stats.GLD) & vbCrLf
        Put n, , "BANCO=0" & vbCrLf
208     Put n, , "MaxHP=" & CStr(.Stats.MaxHP) & vbCrLf
210     Put n, , "MinHP=" & CStr(.Stats.MinHP) & vbCrLf
212     Put n, , "MaxSTA=" & CStr(.Stats.MaxSta) & vbCrLf
214     Put n, , "MinSTA=" & CStr(.Stats.MinSta) & vbCrLf
216     Put n, , "MaxMAN=" & CStr(.Stats.MaxMAN) & vbCrLf
218     Put n, , "MinMAN=" & CStr(.Stats.MinMAN) & vbCrLf
220     Put n, , "MaxHIT=" & CStr(.Stats.MaxHIT) & vbCrLf
222     Put n, , "MinHIT=" & CStr(.Stats.MinHIT) & vbCrLf
224     Put n, , "MaxAGU=" & CStr(.Stats.MaxAGU) & vbCrLf
226     Put n, , "MinAGU=" & CStr(.Stats.MinAGU) & vbCrLf
228     Put n, , "MaxHAM=" & CStr(.Stats.MaxHam) & vbCrLf
230     Put n, , "MinHAM=" & CStr(.Stats.MinHam) & vbCrLf
232     Put n, , "SkillPtsLibres=" & CStr(.Stats.SkillPts) & vbCrLf
234     Put n, , "EXP=" & CStr(.Stats.Exp) & vbCrLf
236     Put n, , "ELV=" & CStr(.Stats.ELV) & vbCrLf
238     Put n, , "ELU=" & CStr(.Stats.ELU) & vbCrLf
        
        
        
        '[MUERTES]
999     Put n, , "[MUERTES]" & vbCrLf
        Put n, , "UserMuertes=0" & vbCrLf
        Put n, , "NpcsMuertes=0" & vbCrLf
        
        
        '[BANCO]
382     Put n, , "[BancoInventory]" & vbCrLf & "CantidadItems=0" & vbCrLf

        Dim loopd As Integer

384     For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
386         Put n, , "Obj" & loopd & "=" & .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount & vbCrLf
388     Next loopd

        
        
    
        '[INVENTARIO]
284     Put n, , "[Inventory]" & vbCrLf & "CantidadItems=" & val(.Invent.NroItems) & vbCrLf

286     For loopc = 1 To MAX_INVENTORY_SLOTS
288         Put n, , "Obj" & loopc & "=" & .Invent.Object(loopc).ObjIndex & "-" & .Invent.Object(loopc).Amount & "-" & .Invent.Object(loopc).Equipped & vbCrLf
        Next
        
        
290     Put n, , "WeaponEqpSlot=" & CStr(.Invent.WeaponEqpSlot) & vbCrLf
310     Put n, , "NudiEqpSlot=" & CStr(.Invent.NudiEqpSlot) & vbCrLf
294     Put n, , "ArmourEqpSlot=" & CStr(.Invent.ArmourEqpSlot) & vbCrLf
296     Put n, , "CascoEqpSlot=" & CStr(.Invent.CascoEqpSlot) & vbCrLf
298     Put n, , "EscudoEqpSlot=" & CStr(.Invent.EscudoEqpSlot) & vbCrLf
300     Put n, , "BarcoSlot=" & CStr(.Invent.BarcoSlot) & vbCrLf
302     Put n, , "MonturaSlot=" & CStr(.Invent.MonturaSlot) & vbCrLf
304     Put n, , "MunicionSlot=" & CStr(.Invent.MunicionEqpSlot) & vbCrLf
306     Put n, , "AnilloSlot=" & CStr(.Invent.AnilloEqpSlot) & vbCrLf
308     Put n, , "MagicSlot=" & CStr(.Invent.MagicSlot) & vbCrLf
    
312
                
                
        '[HECHIZOS]
408     Put n, , "[HECHIZOS]" & vbCrLf

        Dim cad As String

410     For loopc = 1 To MAXUSERHECHIZOS
412         cad = .Stats.UserHechizos(loopc)
414         Put n, , "H" & loopc & "=" & cad & vbCrLf
        Next
      
416
                
 
        '[AMIGOS]
588     Put n, , "[AMIGOS]" & vbCrLf

        Dim cantAmigos As Integer
        
        For cantAmigos = 1 To MAXAMIGOS
            Put n, , "Nombre" & cantAmigos & "=Vacío" & vbCrLf
        Next
        
        
        
                
        '[CORREO]
        Put n, , "[CORREO]" & vbCrLf
        
        Dim CantCorreos As Integer
        
        For CantCorreos = 1 To Max_Correos
            Put n, , "Carta" & CantCorreos & "=0" & vbCrLf
            Put n, , "Emisor" & CantCorreos & "=0" & vbCrLf
            Put n, , "Leida" & CantCorreos & "=0" & vbCrLf
            Put n, , "Objeto" & CantCorreos & "=0-0" & vbCrLf
        Next
        
438     Close #n
    
        'Devuelve el head de muerto
440     If .flags.Muerto = 1 Then
442         .Char.Head = 0
        End If
    
        End With
        
        Exit Sub

SaveNewUser_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.SaveNewUser", Erl)
        Resume Next
        
End Sub


