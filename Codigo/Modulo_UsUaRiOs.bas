Attribute VB_Name = "UsUaRiOs"
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


Sub DoResucitar(ByVal UserIndex As Integer, Optional ByVal IntervaloRevivir As Integer = 2500)

    With UserList(UserIndex)
    
        If Not .flags.Resucitando Then Exit Sub
        
        Dim TActual As Long: TActual = GetTickCount() And &H7FFFFFFF
        
        If TActual - UserList(UserIndex).Counters.IntervaloRevive < IntervaloRevivir Then
            Exit Sub
        Else
            Call DarVida(UserIndex)
        End If
        
    End With
End Sub
Sub RevivirUsuario(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Resucitando <> 1 Then
        UserList(UserIndex).flags.Resucitando = True
        UserList(UserIndex).Counters.IntervaloRevive = GetTickCount
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_RESUCITAR, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, 222, 10, False, True))
        
    End If
    
End Sub
Sub DarVida(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler


    With UserList(UserIndex)
        
    .flags.Muerto = 0
    .flags.Resucitando = False
    
    .Stats.MinHP = .Stats.MaxHP
        
        If .flags.Navegando = 1 Then
            .Char.Head = 0
            .Char.body = iBarca
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            
           If .Invent.EscudoEqpObjIndex > 0 Then
               .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
           Else
               .Char.ShieldAnim = NingunEscudo
           End If
    
           If .Invent.CascoEqpObjIndex > 0 Then
               .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
           Else
               .Char.CascoAnim = NingunCasco
           End If
    
       
           If .Invent.NudiEqpObjIndex > 0 Then
               .Char.WeaponAnim = ObjData(.Invent.NudiEqpObjIndex).WeaponAnim
           Else
               .Char.WeaponAnim = NingunArma
           End If
               
    
           If .Invent.WeaponEqpObjIndex > 0 Then
               .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
           Else
               .Char.WeaponAnim = NingunArma
           End If
           
            If .Invent.ArmourEqpObjIndex > 0 Then
                .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
                .Char.Head = .OrigChar.Head
            End If
            
        End If
        Call WriteLocaleMsg(UserIndex, 90)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_RESUCITADO, .Pos.X, .Pos.Y))
        Call ChangeUserCharTodo(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

        Call WriteUpdateHP(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(.Char.CharIndex, 22, 0, True, True))
        End With
        
        Exit Sub
        
ErrorHandler:
    Call RegistrarError(Err.Number, Err.description, "Usuarios.DarVida", Erl)
    Resume Next
    
    
End Sub

Public Sub ChangeUserCharTodo(ByVal UserIndex As Integer, _
                          ByVal body As Integer, _
                          ByVal Head As Integer, _
                          ByVal heading As Byte, _
                          ByVal Arma As Integer, _
                          ByVal Escudo As Integer, _
                          ByVal casco As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, _
                Arma, Escudo, .FX, .Loops, casco))

    End With

End Sub

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean, ByVal Desvanecer As Boolean)

    '*************************************************
    'Author: Unknown
    'Last modified: 08/01/2009
    '08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
    '*************************************************

    On Error GoTo ErrorHandler
 
    With UserList(UserIndex)
        CharList(.Char.CharIndex) = 0
        
        If .Char.CharIndex = LastChar Then

            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1

                If LastChar <= 1 Then Exit Do
            Loop

        End If
        
        ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex, Desvanecer))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex, Desvanecer))
        End If
        
        Call QuitarUser(UserIndex, .Pos.Map)
        
        MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
        .Char.CharIndex = 0

    End With
    
    NumChars = NumChars - 1
    Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)

End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Tararira
    'Last modified: 04/07/2009
    'Refreshes the status and tag of UserIndex.
    '04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
    '*************************************************
    Dim ClanTag   As String
    With UserList(UserIndex)
Dim Barco As ObjData
        If .GuildIndex > 0 Then
            ClanTag = modGuilds.GuildName(.GuildIndex)
            ClanTag = " <" & ClanTag & ">"

        End If
  
    
        If .showName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .Name & ClanTag, .Faccion.Status, .Donador.activo))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, _
                    vbNullString, 0, 0))
        End If
        
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            Barco = ObjData(.Invent.Object(.Invent.BarcoSlot).ObjIndex)
            .Char.body = Barco.Ropaje
            Call ChangeUserCharTodo(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
        
        Dim NuevaA As Boolean
        Dim Gi     As Integer
        Dim tStr   As String
        
        Gi = .GuildIndex

        If Gi > 0 Then
        
           NuevaA = False
         
          If Not modGuilds.m_ValidarPermanencia(UserIndex, True, NuevaA) Then
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje004)
          End If

          If NuevaA Then
               'Call SendData(SendTarget.ToGuildMembers, Gi, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación " & modGuilds.GuildAlignment(Gi) & "!", FontTypeNames.FONTTYPE_GUILD))
                tStr = modGuilds.GuildName(Gi)
                
               Call LogClanes("¡El clan " & tStr & " cambio de alineación!")
          End If
           
         End If

    End With

End Sub

Public Sub MakeUserChar(ByVal toMap As Boolean, _
                        ByVal sndIndex As Integer, _
                        ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 15/01/2010
    '23/07/2009: Budi - Ahora se envía el nick
    '15/01/2010: ZaMa - Ahora se envia el color del nick.
    '*************************************************

    On Error GoTo ErrHandler

    Dim CharIndex  As Integer
    Dim ClanTag    As String
    Dim UserName   As String
    Dim Privileges As Byte
    
    With UserList(UserIndex)
    
        If InMapBounds(Map, X, Y) Then

            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = UserIndex

            End If
            
            'Place character on map if needed
            If toMap Then MapData(Map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = modGuilds.GuildName(.GuildIndex)

                End If
                
                Privileges = .Faccion.Status
                
                'Preparo el nick
                If .showName Then
                    UserName = .Name

                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or _
                                PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                        Else

                            If (.flags.Invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName
                            Else

                                If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"

                            End If

                        End If

                    

                End If
 
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, .Char.CharIndex, X, Y, _
                        .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, UserName, .Faccion.Status, .Donador.activo, .Char.ParticulaFx, .Char.Arma_Aura, .Char.Body_Aura, .Char.Escudo_Aura, .Char.Head_Aura, .Char.Otra_Aura, .Char.Anillo_Aura)
            Else
                'Hide the name and clan - set privs as normal user
                Call AgregarUser(UserIndex, .Pos.Map)

            End If

        End If

    End With

    Exit Sub

ErrHandler:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(UserIndex)

End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)

    Dim Pts              As Integer
    Dim AumentoHIT       As Integer
    Dim AumentoMANA      As Integer
    Dim AumentoSTA       As Integer
    Dim AumentoHP        As Integer
    Dim WasNewbie        As Boolean
    Dim Promedio         As Double
    Dim aux              As Integer
    Dim DistVida(1 To 5) As Integer
    Dim Gi               As Integer 'Guild Index
    Dim PasoDeNivel      As Boolean
    
    On Error GoTo ErrHandler
    
   
    WasNewbie = EsNewbie(UserIndex)
 
    With UserList(UserIndex)
    
        Do While .Stats.Exp >= .Stats.ELU And .Stats.ELV < STAT_MAXELV
        
 
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub
            End If
    
             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
                   

                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, 127, 3, False, False))
                      
       
       
    
 
    
    
                    

                 
             
             Call WriteLocaleMsg(UserIndex, 186)
             
            .Stats.Exp = .Stats.Exp - .Stats.ELU
             
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
            .Stats.ELU = levelELU(.Stats.ELV)

            Pts = Pts + 5

            'Calculo subida de vida
            Promedio = ModVida(.Clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)
            
            If Promedio - Int(Promedio) = 0.5 Then
            
                'Es promedio semientero
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 1.5

                End If

            Else
                'Es promedio entero
                
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 2
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(4) Then
                    AumentoHP = Promedio - 1
                Else
                    AumentoHP = Promedio - 2

                End If
                
            End If
            
            .Stats.ELV = .Stats.ELV + 1
            
            If UserList(UserIndex).Stats.ELV > 50 Then
             Call WriteConsoleMsg(UserIndex, "Servidor>Ya no Seguiras Obteniendo Bonificaciones ya eres Nivel 50.", FontTypeNames.FONTTYPE_INFO)
             Call WriteUpdateUserStats(UserIndex)
              Exit Sub
            End If
 
            Select Case .Clase
            
                Case eClass.Guerrero
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Cazador
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Paladin
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
                    AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.ladron
                    AumentoHIT = 1
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mago
                    AumentoHIT = 1
                    AumentoMANA = 2.8 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.Leñador
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLeñador
                
                Case eClass.Minero
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTMinero
                
                Case eClass.PescadoR
                    AumentoHIT = 1
                    AumentoSTA = AumentoSTPescador
                
                Case eClass.Clerigo
                    AumentoHIT = 2
                    AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druida
                    AumentoHIT = 2
                    AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Asesino
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
                    AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bardo
                    AumentoHIT = 2
                    AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Herrero, eClass.Carpintero
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Gladiador
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 40, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Nigromante
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 40, 1, 3)
                    AumentoMANA = 2.5 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case eClass.Mercenario
                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 30, 2, 3)
                    AumentoSTA = AumentoSTDef
                    
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP

            If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA

            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            
            'Actualizamos Golpe Máximo
            .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT

            If .Stats.ELV < 36 Then
                If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then .Stats.MaxHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then .Stats.MaxHIT = STAT_MAXHIT_OVER36
            End If
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT

            If .Stats.ELV < 36 Then
                If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then .Stats.MinHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then .Stats.MinHIT = STAT_MAXHIT_OVER36
            End If
            
            'Notificamos al user
            If AumentoHP > 0 Then
            Call WriteLocaleMsg(UserIndex, 197, AumentoHP)
            End If

            If AumentoSTA > 0 Then
                Call WriteLocaleMsg(UserIndex, 198, AumentoSTA)
            End If

            If AumentoMANA > 0 Then
                Call WriteLocaleMsg(UserIndex, 199, AumentoMANA)
            End If

            If AumentoHIT > 0 Then
                Call WriteLocaleMsg(UserIndex, 200, AumentoHIT)
            End If
            
            .Stats.MinHP = .Stats.MaxHP
            
            PasoDeNivel = True
            
            'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
             If Not EsNewbie(UserIndex) And WasNewbie Then
                Call QuitarNewbieObj(UserIndex)
                
                If .Pos.Map = 37 Or .Pos.Map = 208 Then
                'UCase$(MapInfo(.Pos.Map).Zona) = "NEWBIE" Then
                    If esCiuda(UserIndex) Then
                        UserList(UserIndex).Hogar = cNix
                        Call WarpUserChar(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)
                    ElseIf esRepu(UserIndex) Then
                        UserList(UserIndex).Hogar = cIlliandor
                        Call WarpUserChar(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)
                    Else
                        UserList(UserIndex).Hogar = cRinkel
                        Call WarpUserChar(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)
                    End If
                    
                    Call WarpUserChar(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)
                    
                    Call WriteLocaleMsg(UserIndex, 304)
                End If
            End If
            
            Call FlushBuffer(UserIndex)
            DoEvents
            
        Loop

    If PasoDeNivel Then
          
        If .Stats.ELV >= STAT_MAXELV Then
            .Stats.Exp = 0
            .Stats.ELU = 0
            Exit Sub
        End If
        
        Call UpdateUserInv(True, UserIndex, 0, True)
        Call WriteUpdateUserStatsForLevel(UserIndex)
        
        If Pts > 0 Then
            .Stats.SkillPts = .Stats.SkillPts + Pts
             Call WriteLevelUp(UserIndex, Pts)
             
            Call WriteLocaleMsg(UserIndex, 187, Pts)
            Call FlushBuffer(UserIndex)
        End If
    
       If UserList(UserIndex).flags.Meditando Then
           If (ParticleToLevel(UserIndex) <> ParticleToLevel(UserIndex, True)) And (ParticleToLevel(UserIndex, True) <> 0) Then
               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), 0, True, True))
               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex, True), -1, False, True))
           End If
       End If
       
    End If
 
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & _
            Err.description)

End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 Or UserList(UserIndex).flags.Vuela = 1

End Function

Function LegalWalk(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal heading As eHeading, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True) As Boolean
        
        On Error GoTo LegalWalk_Err
        

100      If Map <= 0 Or Map > NumMaps Then Exit Function
        
102     If X < MinXBorder Or X > MaxXBorder Then Exit Function
        
104     If Y < MinYBorder Or Y > MaxYBorder Then Exit Function
        
  
        With MapData(Map, X, Y)
        
110         If .UserIndex <> 0 Then
112             If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                    Exit Function
                End If
            End If
 
114         If Not PuedeTraslado Then
116             If .TileExit.Map > 0 Then Exit Function
            End If
            
118         If Not PuedeAgua Then
120             If (.Blocked And Not HayAgua(Map, X, Y)) <> 0 Then Exit Function
            End If
            
122         If Not PuedeTierra Then
124             If (.Blocked And HayAgua(Map, X, Y)) = 0 Then Exit Function
            End If
            
            If (.Blocked And 2 ^ (heading - 1)) <> 0 Then Exit Function
            
        End With
        
        LegalWalk = True
        
         
 
        
        Exit Function

LegalWalk_Err:
130     Call RegistrarError(Err.Number, Err.description, "Extra.LegalWalk", Erl)
132     Resume Next
        
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

    On Error GoTo MoveUserChar_Err
    
    Dim sailing       As Boolean
    Dim nPos          As WorldPos
    Dim CasperIndex   As Integer
    Dim OppositeHeading As eHeading
    
    With UserList(UserIndex)
   
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = .Pos
    
    Call HeadtoPos(nHeading, nPos)
    
    
     If .flags.Montando = 1 Then
        If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger >= 20) Or (MapInfo(.Pos.Map).Zona = "DUNGEON") Then
            Call DoEquita(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
            Debug.Print "estaba motnando" & Time
        End If
    End If
    
    'If UserList(UserIndex).flags.CheckAmigos = 1 Then
    '    Call MovimientoFriend(UserIndex)
    'End If
    
        'If LegalWalk(.Pos.Map, nPos.X, nPos.Y, nHeading, .flags.Navegando = 1, .flags.Navegando = 0, .flags.montando) Then
        If MoveToLegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
            
            If MapData(nPos.Map, nPos.X, nPos.Y).TileExit.Map <> 0 And .Counters.TiempoDeMapeo > 0 Then
                If .flags.Muerto = 0 Then
                    Call WriteLocaleMsg(UserIndex, 488, .Counters.TiempoDeMapeo)
                    Call WritePosUpdate(UserIndex)
                    Exit Sub
                End If
            End If
                
            'si no estoy solo en el mapa...
            If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
    
                CasperIndex = MapData(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y).UserIndex
    
                'Si hay un usuario, y paso la validacion, entonces es un casper
                If CasperIndex > 0 Then
                
                    If .flags.AdminInvisible = 0 Then
                        
                        Call WritePosUpdate(CasperIndex)
                        OppositeHeading = InvertHeading(nHeading)
                        Call HeadtoPos(OppositeHeading, UserList(CasperIndex).Pos)
                        
                        If UserList(CasperIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(UserList(CasperIndex).Char.CharIndex, UserList(CasperIndex).Pos.X, UserList(CasperIndex).Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(UserList(CasperIndex).Char.CharIndex, UserList(CasperIndex).Pos.X, UserList(CasperIndex).Pos.Y))
                        End If
                         
                        Call WriteForceCharMove(CasperIndex, OppositeHeading)
                        
                        'Update map and char
                        '.Pos = CasPerPos
                         UserList(CasperIndex).Char.heading = OppositeHeading
                         MapData(UserList(CasperIndex).Pos.Map, UserList(CasperIndex).Pos.X, UserList(CasperIndex).Pos.Y).UserIndex = CasperIndex
                         
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.CheckUpdateNeededUser(CasperIndex, OppositeHeading)
                        
                    Else
                        Call WritePosUpdate(UserIndex)
                        Exit Sub
                    End If
                      
                End If
      
                If .flags.AdminInvisible = 0 Then
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                Else
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                End If
                
            End If
                    
            'Update map and user pos
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex Then
                MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
            End If
            
            UserList(UserIndex).Pos = nPos
            UserList(UserIndex).Char.heading = nHeading
            MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
            
                
            'Actualizamos las áreas de ser necesario
            Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
            
        Else
            Call WritePosUpdate(UserIndex)
            
        End If
    
    If UserList(UserIndex).Counters.Trabajando Then UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
    End With
    
    Exit Sub

MoveUserChar_Err:
192     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.MoveUserChar", Erl)
194     Resume Next
        
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

    '*************************************************
    'Author: ZaMa
    'Last modified: 30/03/2009
    'Returns the heading opposite to the one passed by val.
    '*************************************************
    Select Case nHeading

        Case eHeading.EAST
            InvertHeading = WEST

        Case eHeading.WEST
            InvertHeading = EAST

        Case eHeading.SOUTH
            InvertHeading = NORTH

        Case eHeading.NORTH
            InvertHeading = SOUTH

    End Select

End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal slot As Byte, ByRef Object As UserObj)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    UserList(UserIndex).Invent.Object(slot) = Object
    Call WriteChangeInventorySlot(UserIndex, slot)

End Sub
 
Function NextOpenCharIndex() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long
    
    For loopc = 1 To MAXCHARS

        If CharList(loopc) = 0 Then
            NextOpenCharIndex = loopc
            NumChars = NumChars + 1
            
            If loopc > LastChar Then LastChar = loopc
            
            Exit Function

        End If

    Next loopc

End Function

Function NextOpenUser() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long
    
    For loopc = 1 To MaxUsers + 1

        If loopc > MaxUsers Then Exit For
        If (UserList(loopc).ConnID = -1 And UserList(loopc).flags.UserLogged = False) Then Exit For
    Next loopc
    
    NextOpenUser = loopc

End Function

Public Sub SendUserStatsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildI As Integer
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, _
                FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Salud: " & .Stats.MinHP & "/" & .Stats.MaxHP & "  Maná: " & .Stats.MinMAN & _
                "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, _
                FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" _
                    & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & _
                    ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, _
                    FontTypeNames.FONTTYPE_INFO)

        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData( _
                        .Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData( _
                        .Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, _
                        FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData( _
                        .Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, _
                        FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & _
                    "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(SendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
        
        GuildI = .GuildIndex

        If GuildI > 0 Then
            Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)

            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(SendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)

            End If

            'guildpts no tienen objeto
        End If
        
         Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(SendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
       ' Call WriteConsoleMsg( SendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
         
        Call WriteConsoleMsg(SendIndex, "Oro: " & .Stats.GLD & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & _
                .Pos.Map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & _
                .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & _
                ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos( _
                eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)

    End With

End Sub
Sub SendUserInvTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(SendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            If .Invent.Object(j).ObjIndex > 0 Then
            
                Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & _
                        " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    End With

End Sub

Sub SendUserInvTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j        As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd   As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", _
                FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, _
                        FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    Else
        Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserSkillsTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Integer
    
    Call WriteConsoleMsg(SendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(SendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), _
                FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(SendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

End Sub

Private Function EsMascotaCiudadano(ByVal npcindex As Integer, _
                                    ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If Npclist(npcindex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not esRene(Npclist(npcindex).MaestroUser)

        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(npcindex).MaestroUser, "¡¡" & UserList(UserIndex).Name & _
                    " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Function

Sub NPCAtacado(ByVal npcindex As Integer, ByVal UserIndex As Integer)
    '**********************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
    '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
    '06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
    '02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
    '**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(npcindex).flags.AttackedBy = UserList(UserIndex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = npcindex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(npcindex).flags.AttackedFirstBy = vbNullString Then

        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

        Npclist(npcindex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf Npclist(npcindex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then

        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

    End If
    
    If Npclist(npcindex).MaestroUser > 0 Then
        If Npclist(npcindex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, Npclist(npcindex).MaestroUser)

        End If

    End If
    
    If EsMascotaCiudadano(npcindex, UserIndex) Then

        Npclist(npcindex).Movement = TipoAI.NPCDEFENSA
        Npclist(npcindex).Hostile = 1
        
    Else
    
        If Npclist(npcindex).MaestroUser <> UserIndex Then
            'hacemos que el npc se defienda
            Npclist(npcindex).Movement = TipoAI.NPCDEFENSA
            Npclist(npcindex).Hostile = 1

        End If

    End If

End Sub

Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR Or UserList( _
                    UserIndex).Clase = eClass.Asesino

        End If

    End If

End Function
Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
'*************************************************
'Author: Unknown
'Last modified: 30/01/2012
'11/19/2009 Pato   - Implement the new system to train the skills.
'30/01/2012 maTih - Modifico la subida de skills fáciles.
'*************************************************
    With UserList(UserIndex)
    
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
        
            With .Stats
                If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
               
                Dim Lvl As Integer
                Lvl = .ELV
               
                If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
               
                If .UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub

                If RandomNumber(1, 100) > PorcentajeSkill Then Exit Sub
                    .ExpSkills(Skill) = .EluSkills(Skill)
                   
                    If .ExpSkills(Skill) >= .EluSkills(Skill) Then
                        .UserSkills(Skill) = .UserSkills(Skill) + 1
                        
                        
                        Call WriteLocaleMsg(UserIndex, 454, "*" & Skill & "%" & .UserSkills(Skill))
                        .Exp = .Exp + 10
                        If .Exp > MAXEXP Then .Exp = MAXEXP
                        
                        Call WriteLocaleMsg(UserIndex, 140, 10)
                        Call WriteUpdateExp(UserIndex)
                        Call CheckUserLevel(UserIndex)
                        Call FlushBuffer(UserIndex)
                    End If
            End With
        End If
    End With
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)


    On Error GoTo ErrorHandler

    Dim i  As Long
    Dim aN As Integer
    Dim Drops As Obj
    Dim DropObjs As Integer
    
    With UserList(UserIndex)
        
        If TieneObjetos(1601, 1, UserIndex) Then
            UserList(UserIndex).Stats.MinHP = 0
            UserList(UserIndex).flags.Muerto = 1
            UserList(UserIndex).Char.body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).flags.SeguroResu = True
            If UserList(UserIndex).flags.Paralizado <> 0 Or UserList(UserIndex).flags.Inmovilizado <> 0 Then
                UserList(UserIndex).Counters.Paralisis = 0: UserList(UserIndex).flags.Paralizado = 0: UserList(UserIndex).flags.Inmovilizado = 0
                Call WriteParalizeOK(UserIndex)
            End If
            
            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, False)
            Call TirarTodo(UserIndex)
            Call RevivirUsuario(UserIndex)

            Exit Sub
      End If
        'Sonido
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)

        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHP = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .Counters.Veneno = 0
        
        '.flags.Metamorfosis = 0
        .flags.Incinerado = 0
        
        .flags.Muerto = 1
        
        .flags.MuertesUsuario = .flags.MuertesUsuario + 1

        ' No se activa en arenas
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
           ' Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
          '  Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)

        End If
        
        aN = .flags.AtacadoPorNpc

        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Attackable = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString

        End If
        
        aN = .flags.NPCAtacado

        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString

            End If

        End If

        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        Call PerdioNpc(UserIndex)

        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Call RemoveParalisis(UserIndex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), 0, True, True))
            Call WriteMeditateToggle(UserIndex)
        End If
        
        '<<<< Invisible >>>>
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
        
        
        If (TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE And MapInfo(UserList(UserIndex).Pos.Map).Pk = True) Then
            
             DropObjs = TieneSacri(UserIndex)
             If DropObjs = 0 Then
                If Not EsNewbie(UserIndex) Then
                    Call TirarTodo(UserIndex)
                 Else
                    Call TirarTodosLosItemsNoNewbies(UserIndex)
                 End If
             Else
                
                'ReMod Marius Pendiente de sacrificio con 3 estados
                        'Tiramos el sacri 3/3 2/3 1/3
                    If .Invent.Object(DropObjs).ObjIndex = 1081 Then 'Pendiente del Sacrificio 3/3
                            'Debug.Print "2/3"
                        Drops.ObjIndex = 1498 '2/3
                        Drops.Amount = 1
                            
                        'Tilelibre UserList(UserIndex).Pos, newpos, Drops, True, True
                        Call TirarItemAlPiso(UserList(UserIndex).Pos, Drops)
                    ElseIf .Invent.Object(DropObjs).ObjIndex = 1498 Then 'Pendiente del Sacrificio 2/3
                            'Debug.Print "1/3"
                        Drops.ObjIndex = 1499 '1/3
                        Drops.Amount = 1
                            
                        'Tilelibre UserList(UserIndex).Pos, newpos, Drops, True, True
                        Call TirarItemAlPiso(UserList(UserIndex).Pos, Drops)
                    ElseIf .Invent.Object(DropObjs).ObjIndex = 1499 Then 'Pendiente del Sacrificio 1/3
                        'Debug.Print "0/3"
                        'No cae nada, se destruyó el pendiente
                    Else
                        'Debug.Print "Error"
                    End If
                        
                    'Le saca del inventario el pendiente
                    Call QuitarUserInvItem(UserIndex, DropObjs, 1)
                    Call UpdateUserInv(False, UserIndex, DropObjs)
                        
                    'Call DropObj(UserIndex, DropObjs, 1, NewPos.map, NewPos.x, NewPos.Y)
                    '\ReMod
                End If
        End If
        
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
            .Char.Body_Aura = 0
        End If
        
        'desequipar arma
        If .Invent.NudiEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.NudiEqpSlot)
            .Char.Arma_Aura = 0
        End If
            
        'desequipar nudillo
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 1))
            .Char.Arma_Aura = 0
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
            .Char.Head_Aura = 0
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipamos items macigos
        If .Invent.MagicIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)
            .Char.Anillo_Aura = 0
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
            .Char.Escudo_Aura = 0
        End If
 
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.Loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.Loops = 0
        End If
        
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then

            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
            
        End If
        
        Call WriteUpdateDexterity(UserIndex)
        Call WriteUpdateStrenght(UserIndex)
       
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            .Char.body = iFragataFantasmal
        End If
        
        '<<<< Montando >>>>
        If .flags.Montando = 1 Then
            .flags.Montando = 0
            Call WriteMontateToggle(UserIndex)
        End If


        For i = 1 To MAXMASCOTAS

            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
                ' Si estan en agua o zona segura
            Else
                .MascotasType(i) = 0

            End If

        Next i
        
        .NroMascotas = 0
        
        
        '<< Actualizamos clientes >>
        Call ChangeUserCharTodo(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(UserIndex)
      
        '<<Cerramos comercio seguro>>
        'Call LimpiarComercioSeguro(UserIndex)
        
        If TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE Then
            .flags.Resucitando = True
            Call RevivirUsuario(UserIndex)
        End If
        
        Call ControlarPortalLum(UserIndex, 1)
        
        If .Char.ParticulaFx > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, .Char.ParticulaFx, 0, True, False))
        End If
            
            
        Call AgregarParticula(UserIndex, 0, 0, True, False)
                   
    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
    
  ' $ Nuevo Sistema de Facciones $
    If UserList(Atacante).Stats.ELV > UserList(Muerto).Stats.ELV + 10 Then
        Call WriteLocaleMsg(Atacante, 330)
        Exit Sub
        
    End If
    
    If EsNewbie(Muerto) Then Exit Sub
   ' If UserList(Atacante).Pos.Map = Prision.Map Then Exit Sub
     
     'Si esta denudo el muerto no cuenta la muerte
    If UserList(Muerto).flags.Desnudo = 1 Then Exit Sub
    
     'Si esta denudo el muerto no cuenta la muerte
 
    
    With UserList(Atacante)
    
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        
        If esRene(Muerto) Then .Faccion.RenegadosMatados = .Faccion.RenegadosMatados + 1
        If esCiuda(Muerto) Then .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
        If esRepu(Muerto) Then .Faccion.RepublicanosMatados = .Faccion.RepublicanosMatados + 1
        
        
        If esArmada(Muerto) Then .Faccion.ArmadaMatados = .Faccion.ArmadaMatados + 1
        If esMili(Muerto) Then .Faccion.MilicianosMatados = .Faccion.MilicianosMatados + 1
        If esCaos(Muerto) Then .Faccion.CaosMatados = .Faccion.CaosMatados + 1
    
    End With
    
  ' $ Shermie80 $
    
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, _
              ByRef nPos As WorldPos, _
              ByRef Obj As Obj, _
              ByRef Agua As Boolean, _
              ByRef Tierra As Boolean)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
    '**************************************************************
    Dim loopc  As Integer
    Dim tX     As Long
    Dim tY     As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.Map = Pos.Map
    nPos.X = 0
    nPos.Y = 0
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If loopc > 15 Then
            Exit Do

        End If
        
        For tY = Pos.Y - loopc To Pos.Y + loopc
            For tX = Pos.X - loopc To Pos.X + loopc
                
                If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, _
                            tY).ObjInfo.ObjIndex <> Obj.ObjIndex)

                    If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.Amount + Obj.Amount > _
                            MAX_INVENTORY_OBJS)

                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = Pos.X + loopc
                        tY = Pos.Y + loopc

                    End If

                End If
            
            Next tX
        Next tY
        
        loopc = loopc + 1
    Loop

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 ByVal FX As Boolean, _
                 Optional ByVal Teletransported As Boolean)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 13/11/2009
    '15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
    '13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
    '**************************************************************
    Dim OldMap As Integer
    Dim OldX   As Integer
    Dim OldY   As Integer
 
    With UserList(UserIndex)
 

        
        OldMap = .Pos.Map
        OldX = .Pos.X
        OldY = .Pos.Y
        
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        Call WriteRemoveAllDialogs(UserIndex)
        
        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1, True)
        
        If OldMap <> Map Then
        
            If .flags.CasteandoPortal Then

                Call ControlarPortalLum(UserIndex, 2)
            End If
            
            Call WriteChangeMap(UserIndex, Map)
            If Lloviendo Then
            Call WriteRainToggle(UserIndex, Queclima)
            End If
           
             
            'Add Marius Cuando pasas de mapa y no esta permitido invi, te lo saca. Sacado de la 0.13.3 xD
            If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
                If MapInfo(Map).InviSinEfecto > 0 And .flags.Invisible = 1 Then
                    .flags.Invisible = 0
                    .Counters.Invisibilidad = 0
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteLocaleMsg(UserIndex, 447)
                End If
            End If
            '\Add
            
            'Update new Map Users
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0

            End If
        End If
        
        .Pos.X = X
        .Pos.Y = Y
        .Pos.Map = Map
   
        Call MakeUserChar(True, Map, UserIndex, Map, X, Y)
        Call WriteUserCharIndexInServer(UserIndex)
         
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.Invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
        
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, 236, 5, False, False))
        End If
            
        If .NroMascotas Then Call WarpMascotas(UserIndex)
        
        ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        ' Perdes el npc al cambiar de mapa
        Call PerdioNpc(UserIndex)
        
        ' Automatic toogle navigate
      ''  If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
     '       If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
        '        If .flags.Navegando = 0 Then
    '                .flags.Navegando = 1
   '
  '                  'Tell the client that we are navigating.
 '                   Call WriteNavigateToggle(UserIndex)
'
  '              End If

 '           Else
'
     '           If .flags.Navegando = 1 Then
    '                .flags.Navegando = 0
   '
  '                  'Tell the client that we are naviga'ting.
 '                   Call WriteNavigateToggle(UserIndex) '

        '        End If
        '
        '    End If

       ' End If

 
    End With

End Sub

Private Sub WarpMascotas(ByVal UserIndex As Integer)
    '************************************************
    'Author: Uknown
    'Last Modified: 11/05/2009
    '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
    '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
    '11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
    '************************************************
    Dim i                As Integer
    Dim petType          As Integer
    Dim PetRespawn       As Boolean
    Dim PetTiempoDeVida  As Integer
    Dim NroPets          As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp          As Boolean
    Dim index            As Integer
    Dim iMinHP           As Integer
    
    NroPets = UserList(UserIndex).NroMascotas
    canWarp = (MapInfo(UserList(UserIndex).Pos.Map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(UserIndex).MascotasIndex(i)
        
        If index > 0 Then

            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(UserIndex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHP
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(UserIndex).MascotasType(i) = petType

            End If

        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0

        End If
        
        If petType > 0 And canWarp Then
            index = SpawnNpc(petType, UserList(UserIndex).Pos, False, PetRespawn)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If index = 0 Then
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje163)
            Else
                UserList(UserIndex).MascotasIndex(i) = index

                ' Nos aseguramos de que conserve el hp, si estaba dañado
                Npclist(index).Stats.MinHP = IIf(iMinHP = 0, Npclist(index).Stats.MinHP, iMinHP)
            
                Npclist(index).MaestroUser = UserIndex
                Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(index)

            End If

        End If

    Next i
    
    If InvocadosMatados > 0 Then
        'Call WriteMensajes(UserIndex, eMensajes.Mensaje164)

    End If
    
    If Not canWarp Then
        'Call WriteMensajes(UserIndex, eMensajes.Mensaje165)

    End If
    
    UserList(UserIndex).NroMascotas = NroPets

End Sub

Public Sub WarpMascota(ByVal UserIndex As Integer, ByVal PetIndex As Integer)
    '************************************************
    'Author: ZaMa
    'Last Modified: 18/11/2009
    'Warps a pet without changing its stats
    '************************************************
    Dim petType   As Integer
    Dim npcindex  As Integer
    Dim iMinHP    As Integer
    Dim TargetPos As WorldPos
    
    With UserList(UserIndex)
        
        TargetPos.Map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
        
        npcindex = .MascotasIndex(PetIndex)
            
        'Store data and remove NPC to recreate it after warp
        petType = .MascotasType(PetIndex)
        
        ' Guardamos el hp, para restaurarlo cuando se cree el npc
        iMinHP = Npclist(npcindex).Stats.MinHP
        
        Call QuitarNPC(npcindex)
        
        ' Restauramos el valor de la variable
        .MascotasType(PetIndex) = petType
        .NroMascotas = .NroMascotas + 1
        npcindex = SpawnNpc(petType, TargetPos, False, False)
        
        'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
        ' Exception: Pets don't spawn in water if they can't swim
        If npcindex = 0 Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje166)
        Else
            .MascotasIndex(PetIndex) = npcindex

            With Npclist(npcindex)
                ' Nos aseguramos de que conserve el hp, si estaba dañado
                .Stats.MinHP = IIf(iMinHP = 0, .Stats.MinHP, iMinHP)
            
                .MaestroUser = UserIndex
                .Movement = TipoAI.SigueAmo
                .Target = 0
                .TargetNPC = 0

            End With
            
            Call FollowAmo(npcindex)

        End If

    End With

End Sub
Sub Cerrar_Usuario(ByVal UserIndex As Integer)

    On Error GoTo ErrorHand
    
      With UserList(UserIndex)

1         If .Counters.Saliendo Then
2             Call CancelExit(UserIndex)
3             Exit Sub
4         End If
          
6         Dim isNotVisible As Boolean

20            If .flags.UserLogged And Not .Counters.Saliendo Then
21
22                 .Counters.Saliendo = True
24                 .Counters.Salir = IntervaloCerrarConexion

26                 Call WriteEjecutarAccion(UserIndex, 1) ' Iniciar salida

27                 isNotVisible = (UserList(UserIndex).flags.Oculto Or UserList(UserIndex).flags.Invisible)
28
29                If isNotVisible Then
30                    .flags.Oculto = 0
31                    .flags.Invisible = 0
32                    .Counters.Invisibilidad = 0
33                    .Counters.TiempoOculto = 0
34                    Call WriteLocaleMsg(UserIndex, 307)
35                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
36                End If

37                If .flags.Trabajando = True Then
38                     Call WriteLocaleMsg(UserIndex, 391, vbNullString, 1)
39                     .flags.Trabajando = False
40                     .flags.Lingoteando = 0
41                End If
42
13                If Not MapInfo(.Pos.Map).Pk Or .flags.Muerto Or EsGm(UserIndex) Then
14                    Call WriteDisconnect(UserIndex)
15                    Call FlushBuffer(UserIndex)
16                    Call CloseSocket(UserIndex)
17                    Exit Sub
                  End If
          
                  Call WriteLocaleMsg(UserIndex, 203, UserList(UserIndex).Counters.Salir)
                    
44            End If
 
        
46    End With

47    Exit Sub

ErrorHand:
    Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.Cerrar_Usuario", Erl)
    Resume Next

End Sub

Public Sub CancelExit(ByVal UserIndex As Integer)


    If UserList(UserIndex).Counters.Saliendo Then
    
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            
            Call WriteEjecutarAccion(UserIndex, 1) '  'Cancelado cierre
            
            Call WriteLocaleMsg(UserIndex, 365, , 1)
            
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Pk, IntervaloCerrarConexion, 0)
        End If
    
    End If
    
    If UserList(UserIndex).Counters.IdleCount > 0 Then
        UserList(UserIndex).Counters.IdleCount = 0
    End If
    
End Sub


'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, _
                       ByVal UserIndexDestino As Integer, _
                       ByVal NuevoNick As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ViejoNick       As String
    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup

    End If

End Sub

Sub SendUserStatsTxtOFF(ByVal SendIndex As Integer, ByVal Nombre As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
        Call WriteLocaleMsg(SendIndex, 80)
    Else
        Call WriteConsoleMsg(SendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & _
                GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", _
                "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & _
                GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & _
                GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Maná: " & GetVar(CharPath & Nombre & _
                ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), _
                FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", _
                "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(SendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), _
                FontTypeNames.FONTTYPE_INFO)
        
             Dim TempSecs As Long
            Dim TempStr  As String
            TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
            TempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod _
                    86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(SendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
     
    End If

End Sub

Sub SendUserOROTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim CharFile As String
    
    On Error Resume Next

    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", _
                FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Public Function BodyIsBoat(ByVal body As Integer) As Boolean

    If body = iGraficos.iBarca Or body = iGraficos.iFragataFantasmal Or body = iGraficos.iGaleon Or body = iGraficos.iGalera Then
        BodyIsBoat = True
    End If

End Function
 
Public Function IsArena(ByVal UserIndex As Integer) As Boolean
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 10/11/2009
    'Returns true if the user is in an Arena
    '**************************************************************
    IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)

End Function

Public Sub PerdioNpc(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 18/01/2010 (ZaMa)
    'The user loses his owned npc
    '18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
    '**************************************************************

    Dim PetIndex As Long
    
    With UserList(UserIndex)

        If .flags.OwnedNpc > 0 Then
            Npclist(.flags.OwnedNpc).Owner = 0
            .flags.OwnedNpc = 0
            
            ' Dejan de atacar las mascotas
            If .NroMascotas > 0 Then

                For PetIndex = 1 To MAXMASCOTAS

                    If .MascotasType(PetIndex) > 0 Then Call FollowAmo(PetIndex)
                Next PetIndex

            End If

        End If

    End With

End Sub

Public Sub ApropioNpc(ByVal UserIndex As Integer, ByVal npcindex As Integer)
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 18/01/2010 (zaMa)
    'The user owns a new npc
    '18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
    '19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
    '**************************************************************

    With UserList(UserIndex)

        ' Los admins no se pueden apropiar de npcs
        If EsGm(UserIndex) Then Exit Sub
        
        'No aplica a zonas seguras
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then Exit Sub
        
        ' No aplica a algunos mapas que permiten el robo de npcs
        If MapInfo(.Pos.Map).RoboNpcsPermitido = 1 Then Exit Sub
        
        ' Pierde el npc anterior
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
        ' Si tenia otro dueño, lo perdio aca
        Npclist(npcindex).Owner = UserIndex
        .flags.OwnedNpc = npcindex

    End With
    
    ' Inicializo o actualizo el timer de pertenencia
    Call IntervaloPerdioNpc(UserIndex, True)

End Sub

Public Function GetDireccion(ByVal UserIndex As Integer, _
                             ByVal OtherUserIndex As Integer) As String
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Devuelve la direccion hacia donde esta el usuario
    '**************************************************************
    Dim X As Integer
    Dim Y As Integer
    
    X = UserList(UserIndex).Pos.X - UserList(OtherUserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    
    If X = 0 And Y > 0 Then
        GetDireccion = "Sur"
    ElseIf X = 0 And Y < 0 Then
        GetDireccion = "Norte"
    ElseIf X > 0 And Y = 0 Then
        GetDireccion = "Este"
    ElseIf X < 0 And Y = 0 Then
        GetDireccion = "Oeste"
    ElseIf X > 0 And Y < 0 Then
        GetDireccion = "NorEste"
    ElseIf X < 0 And Y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf X > 0 And Y > 0 Then
        GetDireccion = "SurEste"
    ElseIf X < 0 And Y > 0 Then
        GetDireccion = "SurOeste"

    End If

End Function

Public Function SameFaccion(ByVal UserIndex As Integer, _
                            ByVal OtherUserIndex As Integer) As Boolean
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Devuelve True si son de la misma faccion
    '**************************************************************
    SameFaccion = (esCaos(UserIndex) And esCaos(OtherUserIndex)) Or (esArmada(UserIndex) And esArmada(OtherUserIndex))

End Function

Public Function FarthestPet(ByVal UserIndex As Integer) As Integer

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 18/11/2009
    'Devuelve el indice de la mascota mas lejana.
    '**************************************************************
    On Error GoTo ErrHandler
    
    Dim PetIndex      As Integer
    Dim Distancia     As Integer
    Dim OtraDistancia As Integer
    
    With UserList(UserIndex)

        If .NroMascotas = 0 Then Exit Function
    
        For PetIndex = 1 To MAXMASCOTAS

            ' Solo pos invocar criaturas que exitan!
            If .MascotasIndex(PetIndex) > 0 Then

                ' Solo aplica a mascota, nada de elementales..
                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        ' Por si tiene 1 sola mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + Abs(.Pos.Y - Npclist( _
                                .MascotasIndex(PetIndex)).Pos.Y)
                    Else
                        ' La distancia de la proxima mascota
                        OtraDistancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + Abs(.Pos.Y - Npclist( _
                                .MascotasIndex(PetIndex)).Pos.Y)

                        ' Esta mas lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex

                        End If

                    End If

                End If

            End If

        Next PetIndex

    End With

    Exit Function
    
ErrHandler:
    Call LogError("Error en FarthestPet")

End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal UserIndex As Integer, _
                         ByVal Skill As Byte, _
                         ByVal Allocation As Boolean)
    '*************************************************
    'Author: Torres Patricio (Pato)
    'Last modified: 11/20/2009
    '
    '*************************************************

    With UserList(UserIndex).Stats

        If .UserSkills(Skill) < MAXSKILLPOINTS Then
            If Allocation Then
                .ExpSkills(Skill) = 0
            Else
                .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)

            End If
        
            .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
        Else
            .ExpSkills(Skill) = 0
            .EluSkills(Skill) = 0

        End If

    End With

End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, _
                               ByVal ObjIndex As Integer, _
                               ByVal Amount As Long) As Boolean
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 25/11/2009
    'Cheks Wether the user has the required amount of items in the inventory or not
    '**************************************************************

    Dim slot          As Long
    Dim ItemInvAmount As Long
    
    For slot = 1 To MAX_INVENTORY_SLOTS
        ' Si es el item que busco
        
        If UserList(UserIndex).Invent.Object(slot).ObjIndex = ObjIndex Then
            ' Lo sumo a la cantidad total
            ItemInvAmount = ItemInvAmount + UserList(UserIndex).Invent.Object(slot).Amount

        End If

    Next slot

    HasEnoughItems = Amount <= ItemInvAmount

End Function
Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento Or _
               OBJType = eOBJType.otCorreo Or _
               OBJType = eOBJType.otArboles

End Function
Public Function ParticleToLevel(ByVal UserIndex As Integer, Optional ByVal CambioStats As Boolean = False) As Integer

Dim nivel As Byte

nivel = UserList(UserIndex).Stats.ELV


If CambioStats = True Then nivel = UserList(UserIndex).Stats.ELV + 1
 

If EsGm(UserIndex) Then
ParticleToLevel = 305 '
Else


Select Case nivel

Case Is < 13
ParticleToLevel = 42

Case Is < 25
ParticleToLevel = 81

Case Is < 34
ParticleToLevel = 41

Case Is < 45
    Select Case UserList(UserIndex).Faccion.Status
    Case 1
    ParticleToLevel = 39
    Case 2
    ParticleToLevel = 40
    Case 3
    ParticleToLevel = 71
    Case 4
    ParticleToLevel = 37
    Case 5
    ParticleToLevel = 38
    Case 6
    ParticleToLevel = 66
    End Select

Case Is < 50
    Select Case UserList(UserIndex).Faccion.Status
    Case 1
    ParticleToLevel = 161
    Case 2
    ParticleToLevel = 160
    Case 3
    ParticleToLevel = 176
    Case 4
    ParticleToLevel = 162
    Case 5
    ParticleToLevel = 163
    Case 6
    ParticleToLevel = 176
    End Select
    
Case Is < 55
    Select Case UserList(UserIndex).Faccion.Status
    Case 1
    ParticleToLevel = 168
    Case 2
    ParticleToLevel = 169
    Case 3
    ParticleToLevel = 173
    Case 4
    ParticleToLevel = 170
    Case 5
    ParticleToLevel = 171
    Case 6
    ParticleToLevel = 173
    End Select
Case Is < 59
    Select Case UserList(UserIndex).Faccion.Status
    Case 1
    ParticleToLevel = 167
    Case 2
    ParticleToLevel = 166
    Case 3
    ParticleToLevel = 174
    Case 4
    ParticleToLevel = 165
    Case 5
    ParticleToLevel = 164
    Case 6
    ParticleToLevel = 174
    End Select
Case 60
    Select Case UserList(UserIndex).Faccion.Status
    Case 1
    ParticleToLevel = 36
    Case 2
    ParticleToLevel = 110
    Case 3
    ParticleToLevel = 109
    Case 4
    ParticleToLevel = 62
    Case 5
    ParticleToLevel = 113
    Case 6
    ParticleToLevel = 111
    End Select

End Select
End If


'Agregamos la particula al char
UserList(UserIndex).Char.Particles(1) = ParticleToLevel

End Function
 Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
 
 If UserList(attackerIndex).Stats.ELV > UserList(VictimIndex).Stats.ELV + 10 Then Exit Sub
    Dim DaExp As Integer
    DaExp = CInt(UserList(VictimIndex).Stats.ELV) * 2
    With UserList(attackerIndex)
        
        .Stats.Exp = .Stats.Exp + DaExp
       ' If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        'Lo mata¡
        Call WriteConsoleMsg(attackerIndex, "¡Has matado a " & UserList(VictimIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(VictimIndex, "¡" & .Name & " te ha matado!", FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(attackerIndex, "Has ganado " & CStr(DaExp) & " puntos de experiencia.", FontTypeNames.FONTTYPE_INFO)
             
        Call FlushBuffer(VictimIndex)
         'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(VictimIndex).Name)
        
        Call WriteUpdateExp(attackerIndex)
        
    End With
End Sub
