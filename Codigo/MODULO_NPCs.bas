Attribute VB_Name = "NPCs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) = npcindex Then
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
         
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
            Exit For

        End If

    Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

End Sub

Sub MuereNpc(ByVal npcindex As Integer, ByVal UserIndex As Integer)

    '********************************************************
    'Author: Unknown
    'Llamado cuando la vida de un NPC llega a cero.
    'Last Modify Date: 24/01/2007
    '24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
    '********************************************************
    On Error GoTo ErrHandler

    Dim MiNPC As npc
    MiNPC = Npclist(npcindex)

    'Quitamos el npc
    Call QuitarNPC(npcindex)
    
    If UserIndex > 0 Then ' Lo mato un usuario?

        With UserList(UserIndex)
        
            If MiNPC.flags.Snd3 > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, _
                        MiNPC.Pos.Y))

            End If

            .flags.TargetNPC = 0
            .flags.TargetNpcTipo = eNPCType.Comun
            
            'El user que lo mato tiene mascotas?
            If .NroMascotas > 0 Then
                Dim t As Integer

                For t = 1 To MAXMASCOTAS

                    If .MascotasIndex(t) > 0 Then
                        If Npclist(.MascotasIndex(t)).TargetNPC = npcindex Then
                            Call FollowAmo(.MascotasIndex(t))

                        End If

                    End If

                Next t

            End If
            
            '[KEVIN]
            If MiNPC.flags.ExpCount > 0 Then
 
            
            
                    .Stats.Exp = .Stats.Exp + MiNPC.flags.ExpCount

                    If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                      Call WriteLocaleMsg(UserIndex, 140, MiNPC.flags.ExpCount)
 
                   

               

                MiNPC.flags.ExpCount = 0

            End If
            Call WriteLocaleMsg(UserIndex, 184, MiNPC.Name)
            '[/KEVIN]
            
            If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1

            Call CheckUserLevel(UserIndex)
            
            If npcindex = .flags.ParalizedByNpcIndex Then
                Call RemoveParalisis(UserIndex)

            End If
            
        End With

    End If ' Userindex > 0
   
    If MiNPC.MaestroUser = 0 Then
        
         If MiNPC.GiveGLD > 0 Then Call NPCTirarOro(MiNPC, UserIndex)
    
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(UserIndex, MiNPC)
        'ReSpawn o no
        Call ReSpawnNpc(MiNPC)

    End If
    
    Exit Sub

ErrHandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.description)

End Sub

Private Sub ResetNpcFlags(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Clear the npc's flags
    
    With Npclist(npcindex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Domable = 0
        .Envenenado = 0
        .Status = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .Invisible = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0

    End With

End Sub

Private Sub ResetNpcCounters(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(npcindex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
        .Ataque = 0
    End With

End Sub

Private Sub ResetNpcCharInfo(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(npcindex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0

    End With

End Sub

Private Sub ResetNpcCriatures(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Long
    
    With Npclist(npcindex)

        For j = 1 To .NroCriaturas
            .Criaturas(j).npcindex = 0
        Next j
        
        .NroCriaturas = 0

    End With

End Sub
 
Private Sub ResetNpcMainInfo(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(npcindex)
        .Attackable = 0
        .CanAttack = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, npcindex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        
        .MaestroUser = 0
        .MaestroNpc = 0
        
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCtype = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .desc = vbNullString
        
        Dim j As Long

        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j

    End With
    
    Call ResetNpcCharInfo(npcindex)
    Call ResetNpcCriatures(npcindex)

End Sub

Public Sub QuitarNPC(ByVal npcindex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Now npcs lose their owner
    '***************************************************
    On Error GoTo ErrHandler

    With Npclist(npcindex)
        .flags.NPCActive = False
        
        .Owner = 0 ' Murio, no necesita mas dueños :P.
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(npcindex)

        End If

    End With
        
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(npcindex)
    Call ResetNpcFlags(npcindex)
    Call ResetNpcCounters(npcindex)
    
    Call ResetNpcMainInfo(npcindex)
    
    If npcindex = LastNPC Then

        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1

            If LastNPC < 1 Then Exit Do
        Loop

    End If
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1

    End If

    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarNPC")

End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal npcindex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 18/11/2009
    'Kills a pet
    '***************************************************
    On Error GoTo ErrHandler

    Dim i        As Integer
    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        ' Busco el indice de la mascota
        For i = 1 To MAXMASCOTAS

            If .MascotasIndex(i) = npcindex Then PetIndex = i
        Next i
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        ' Limpio el slot de la mascota
        .NroMascotas = .NroMascotas - 1
        .MascotasIndex(PetIndex) = 0
        .MascotasType(PetIndex) = 0
        
        ' Elimino la mascota
        Call QuitarNPC(npcindex)

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.description & " NpcIndex: " & npcindex _
            & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)

End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, _
                                  Optional PuedeAgua As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 3 And MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> _
                2 And MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> 1

    End If
    
End Function

Sub CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Crea un NPC del tipo NRONPC

    Dim Pos            As WorldPos
    Dim newpos         As WorldPos
    Dim altpos         As WorldPos
    Dim nIndex         As Integer
    Dim PosicionValida As Boolean
    Dim Iteraciones    As Long
    Dim PuedeAgua      As Boolean
    Dim PuedeTierra    As Boolean

    Dim Map            As Integer
    Dim X              As Integer
    Dim Y              As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Sub
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.Map = Mapa 'mapa
        altpos.Map = Mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana

            If newpos.X <> 0 And newpos.Y <> 0 Then
                altpos.X = newpos.X
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)

                If newpos.X <> 0 And newpos.Y <> 0 Then
                    altpos.X = newpos.X
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)

                End If

            End If

            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, PuedeAgua) And Not HayPCarea(newpos) And TestSpawnTrigger( _
                    newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.Map = newpos.Map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1

            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    Map = altpos.Map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).Pos.Map = Map
                    Npclist(nIndex).Pos.X = X
                    Npclist(nIndex).Pos.Y = Y
                    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
                    Exit Sub
                Else
                    altpos.X = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)

                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.Map = newpos.Map
                        Npclist(nIndex).Pos.X = newpos.X
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(True, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
                        Exit Sub

                    End If

                End If

            End If

        Loop
            
        'asignamos las nuevas coordenas
        Map = newpos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y

    End If
 
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    Debug.Print "Respawn de: " & Npclist(nIndex).Name; " - " & "(" & Map & "," & X & "," & Y & ")"
    
End Sub

Public Sub MakeNPCChar(ByVal toMap As Boolean, _
                       sndIndex As Integer, _
                       npcindex As Integer, _
                       ByVal Map As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    Dim CharIndex As Integer

    If Npclist(npcindex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(npcindex).Char.CharIndex = CharIndex
        CharList(CharIndex) = npcindex

    End If

    
   ' Dim showName    As Byte
   ' showName = val(GetVar(DatPath & "NPCs.dat", "NPC" & Npclist(npcindex).Numero, "ShowName"))
    
    MapData(Map, X, Y).npcindex = npcindex
    
    If Not toMap Then
        Call WriteCharacterCreate(sndIndex, Npclist(npcindex).Char.body, Npclist(npcindex).Char.Head, Npclist(npcindex).Char.heading, Npclist(npcindex).Char.CharIndex, X, Y, Npclist(npcindex).Char.WeaponAnim, Npclist(npcindex).Char.ShieldAnim, 0, 0, Npclist(npcindex).Char.CascoAnim, Npclist(npcindex).Numero, Npclist(npcindex).flags.Status, 0, 0, 0, 0, 0, 0, 0, 0)
        


        'If showName = 1 Then 'Fuerzas del Caos
        'Call WriteCharStatus(sndIndex, Npclist(npcindex).Char.CharIndex, 4) 'Fuerzas del Caos
        'End If
       ' If showName = 2 Then 'Republicano
       ' Call WriteCharStatus(sndIndex, Npclist(npcindex).Char.CharIndex, 5) 'Republicano
       ' End If
        
       '         If showName = 3 Then 'Imperial
       ' Call WriteCharStatus(sndIndex, Npclist(npcindex).Char.CharIndex, 6) 'Imperial
       ' End If
        
      '             If showName = 4 Then 'Neutral
      '  Call WriteCharStatus(sndIndex, Npclist(npcindex).Char.CharIndex, 7) 'GMG
      '  End If
        
        
      '                     If showName = 5 Then 'GM AMARIILLO
      '  Call WriteCharStatus(sndIndex, Npclist(npcindex).Char.CharIndex, 1) 'GMG
      '  End If
        
     '                      If showName = 6 Then 'GM VERDE
      '  Call WriteCharStatus(sndIndex, Npclist(npcindex).Char.CharIndex, 3) 'GMG
     '   End If
        Call FlushBuffer(sndIndex)
    Else
        Call AgregarNpc(npcindex)

    End If

End Sub
Public Sub ChangeNPCChar(ByVal npcindex As Integer, _
                         ByVal body As Integer, _
                         ByVal Head As Integer, _
                         ByVal heading As eHeading)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If npcindex > 0 Then

        With Npclist(npcindex).Char
            .body = body
            .Head = Head
            .heading = heading
            
            Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessageCharacterChange(body, Head, heading, _
                    .CharIndex, 0, 0, 0, 0, 0))

        End With

    End If

End Sub
Private Sub EraseNPCChar(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If Npclist(npcindex).Char.CharIndex <> 0 Then CharList(Npclist(npcindex).Char.CharIndex) = 0

    If Npclist(npcindex).Char.CharIndex = LastChar Then

        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1

            If LastChar <= 1 Then Exit Do
        Loop

    End If

    'Quitamos del mapa
    MapData(Npclist(npcindex).Pos.Map, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y).npcindex = 0

    'Actualizamos los clientes
    Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessageCharacterRemove(Npclist(npcindex).Char.CharIndex, True))

    'Update la lista npc
    Npclist(npcindex).Char.CharIndex = 0

    'update NumChars
    NumChars = NumChars - 1

End Sub

Public Sub MoveNPCChar(ByVal npcindex As Integer, ByVal nHeading As Byte)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2009
    '06/04/2009: ZaMa - Now npcs can force to change position with dead character
    '01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
    '***************************************************

    On Error GoTo errh

    Dim nPos      As WorldPos
    Dim UserIndex As Integer
    
    With Npclist(npcindex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalPosNPC(.Pos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Sub
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex

            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Sub

                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Sub
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(npcindex).Pos.X
                    .Pos.Y = Npclist(npcindex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList( _
                            UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))

                End With

            End If
            
            Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).npcindex = 0
            .Pos = nPos
            .Char.heading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).npcindex = npcindex
            Call CheckUpdateNeededNpc(npcindex, nHeading)
        
        ElseIf .MaestroUser = 0 Then

            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0

            End If

        End If

    End With

    Exit Sub

errh:
    LogError ("Error en move npc " & npcindex)

End Sub

Function NextOpenNPC() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim loopc As Long
      
    For loopc = 1 To MAXNPCS + 1

        If loopc > MAXNPCS Then Exit For
        If Not Npclist(loopc).flags.NPCActive Then Exit For
    Next loopc
      
    NextOpenNPC = loopc
    Exit Function

ErrHandler:
    Call LogError("Error en NextOpenNPC")

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer
    n = RandomNumber(1, 100)

    If n < 30 Then
        UserList(UserIndex).flags.Envenenado = 1
        Call WriteLocaleMsg(UserIndex, 182)
    
    End If

End Sub

Function SpawnNpc(ByVal npcindex As Integer, _
                  Pos As WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean) As Integer
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/15/2008
    '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
    '06/15/2008 -> Optimizé el codigo. (NicoNZ)
    '***************************************************
    Dim newpos         As WorldPos
    Dim altpos         As WorldPos
    Dim nIndex         As Integer
    Dim PosicionValida As Boolean
    Dim PuedeAgua      As Boolean
    Dim PuedeTierra    As Boolean

    Dim Map            As Integer
    Dim X              As Integer
    Dim Y              As Integer
    
    nIndex = OpenNPC(npcindex, Respawn)   'Conseguimos un indice

    If nIndex > MAXNPCS Then
        SpawnNpc = 0
        Exit Function

    End If

    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
    Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
    Call ClosestLegalPos(Pos, altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida

    If newpos.X <> 0 And newpos.Y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        Npclist(nIndex).Pos.Map = newpos.Map
        Npclist(nIndex).Pos.X = newpos.X
        Npclist(nIndex).Pos.Y = newpos.Y
        PosicionValida = True
    Else

        If altpos.X <> 0 And altpos.Y <> 0 Then
            Npclist(nIndex).Pos.Map = altpos.Map
            Npclist(nIndex).Pos.X = altpos.X
            Npclist(nIndex).Pos.Y = altpos.Y
            PosicionValida = True
        Else
            PosicionValida = False

        End If

    End If

    If Not PosicionValida Then
        Call QuitarNPC(nIndex)
        SpawnNpc = 0
        Exit Function

    End If

    'asignamos las nuevas coordenas
    Map = newpos.Map
    
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y

    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, _
                FXIDs.FXWARP, 0))

    End If

    SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

Private Sub NPCTirarOro(ByRef MiNPC As npc, ByVal UserIndex As Integer)
 

'SI EL NPC TIENE ORO LO TIRAMOS
 
    If UserIndex > 0 Then
        If OroLleno(UserIndex, UserList(UserIndex).Stats.GLD, MiNPC.GiveGLD) Then
            Call WriteLocaleMsg(UserIndex, 378)
            Exit Sub
        End If
        
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiNPC.GiveGLD
            Call WriteLocaleMsg(UserIndex, 29, MiNPC.GiveGLD)
            Call WriteUpdateGold(UserIndex)
            
    End If

End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, _
                        Optional ByVal Respawn = True) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
 
    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
    '
    'El que ose desafiar esta LEY, se las tendrá que ver
    'conmigo. Para leer los NPCS se deberá usar la
    'nueva clase clsIniReader.
    '
    'Alejo
    '
    '###################################################
    
    On Error GoTo OpenNPC_Err
    
    
100    Dim npcindex As Integer

102    Dim Leer     As clsIniManager

104    Dim loopc    As Long

106    Dim ln       As String
   
108    Set Leer = LeerNPCs
   
    'If requested index is invalid, abort
110    If Not Leer.KeyExists("NPC" & NpcNumber) Then
         OpenNPC = MAXNPCS + 1
112        Exit Function

114    End If
   
116    npcindex = NextOpenNPC
   
118    If npcindex > MAXNPCS Then 'Limite de npcs
120        OpenNPC = npcindex
122        Exit Function

124    End If
   
126    With Npclist(npcindex)
128        .Numero = NpcNumber
130        .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
132        .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
       
134        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
136        .Leveles = val(Leer.GetValue("NPC" & NpcNumber, "Nivel"))
            If Not .Leveles > 0 Then .Leveles = 100

138        .flags.OldMovement = .Movement
       
140        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
142        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
144        .flags.Status = val(Leer.GetValue("NPC" & NpcNumber, "Status"))
            
146        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
148        .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
       
150        .Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
152        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "ShieldAnim"))
154        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "weaponanim"))
156        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
158        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
160        .Char.heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
       
162        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
164        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
166        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
168        .flags.OldHostil = .Hostile
        
170        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * Expc
       
172        .flags.ExpCount = .GiveEXP
       
174        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
       
176        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
       
178        .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD")) * Oroc
       
180        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
182        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
       
184        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
       
186        With .Stats
188            .MaxHP = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
190            .MinHP = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
192            .MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
194            .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
196            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
198        End With
       
200        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
        
202        For loopc = 1 To Npclist(npcindex).Invent.NroItems
204            ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & loopc)
206               .Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
208               .Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
210        Next loopc
 
            
            '//Conversion de Drop
            If val(Leer.GetValue("NPC" & NpcNumber, "RandomDrop")) > 0 Then
           
            Dim DropString As String, Longitud As String
            DropString = Leer.GetValue("NPC" & NpcNumber, "RandomDrop")
     
            Dim Cantidad As Byte
            Cantidad = 0
            Dim nPos As Integer: nPos = 0
     
            Do
            nPos = InStr(nPos + 1, DropString, "-")
            
            If nPos = 0 Then Exit Do
            
                Cantidad = Cantidad + 1
                
                Loop While True

                   .Invent.NroItems = Cantidad + 1
     
                    For loopc = 1 To (.Invent.NroItems)
                            .Invent.Object(loopc).ObjIndex = val(ReadField(1, DropString, Asc(",")))
                            .Invent.Object(loopc).Amount = val(ReadField(2, DropString, Asc(",")))
                            .Invent.Object(loopc).ProbTirar = val(ReadField(3, DropString, Asc(",")))
                             Longitud = Len(.Invent.Object(loopc).ObjIndex & "," & .Invent.Object(loopc).Amount & "," & .Invent.Object(loopc).ProbTirar & "-")
                             DropString = mid(DropString, Longitud + 1)
                    Next loopc
     
           End If
           'Fin '//Conversion de Drop
           
212        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))

214        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)

216        For loopc = 1 To .flags.LanzaSpells
218            .Spells(loopc) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & loopc))
220        Next loopc
       
222        If .NPCtype = eNPCType.Entrenador Then
224            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))

226            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador

228            For loopc = 1 To .NroCriaturas
230                .Criaturas(loopc).npcindex = Leer.GetValue("NPC" & NpcNumber, "CI" & loopc)
234            Next loopc

236        End If
       
238        With .flags
240            .NPCActive = True
           
242            If Respawn Then
244                .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "RespawnTime"))
246            Else
248                .Respawn = 1

250            End If
           
                .Respawn = 0 'Mermas temporal porque los npcs n respawneana
                
252            .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
256            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "PosOrig"))
258            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "Inmunidad"))
           
260            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
262            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
264            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

266        End With
       
        'Tipo de items con los que comercia
268       .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
           
270     End With
   
    'Update contadores de NPCs
272    If npcindex > LastNPC Then LastNPC = npcindex
274    NumNPCs = NumNPCs + 1
   
    'Devuelve el nuevo Indice
276    OpenNPC = npcindex
 

278        Exit Function

OpenNPC_Err:
280     Call RegistrarError(Err.Number, Err.description, "NPCs.OpenNPC", Erl)
282     Resume Next
        
End Function


Public Sub DoFollow(ByVal npcindex As Integer, ByVal UserName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(npcindex)

        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0

        End If

    End With

End Sub

Public Sub FollowAmo(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(npcindex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0

    End With

End Sub

Public Sub ValidarPermanenciaNpc(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Chequea si el npc continua perteneciendo a algún usuario
    '***************************************************

    With Npclist(npcindex)

        If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)

    End With

End Sub

Public Sub ReloadNPCByIndex(ByVal npcindex As Integer)

    On Error GoTo ErrHandler

    Dim NpcNumber As Integer
    Dim loopc As Long
    Dim ln As String

126    With Npclist(npcindex)
128       .Numero = NpcNumber
130        .Name = LeerNPCs.GetValue("NPC" & NpcNumber, "Name")
132        .desc = LeerNPCs.GetValue("NPC" & NpcNumber, "Desc")
       
134        .Movement = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Movement"))
136        .Leveles = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Nivel"))
138        .flags.OldMovement = .Movement
       
140        .flags.AguaValida = val(LeerNPCs.GetValue("NPC" & NpcNumber, "AguaValida"))
142        .flags.TierraInvalida = val(LeerNPCs.GetValue("NPC" & NpcNumber, "TierraInValida"))
144        .flags.Status = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Status"))
            
146        .flags.AtacaDoble = val(LeerNPCs.GetValue("NPC" & NpcNumber, "AtacaDoble"))
148        .NPCtype = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NpcType"))
       
150        .Char.body = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Body"))
152        .Char.ShieldAnim = val(LeerNPCs.GetValue("NPC" & NpcNumber, "ShieldAnim"))
154        .Char.WeaponAnim = val(LeerNPCs.GetValue("NPC" & NpcNumber, "weaponanim"))
156        .Char.CascoAnim = val(LeerNPCs.GetValue("NPC" & NpcNumber, "CascoAnim"))
158        .Char.Head = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Head"))
160        .Char.heading = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Heading"))
       
162        .Attackable = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Attackable"))
164        .Comercia = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Comercia"))
166        .Hostile = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Hostile"))
168        .flags.OldHostil = .Hostile
        
170        .GiveEXP = val(LeerNPCs.GetValue("NPC" & NpcNumber, "GiveEXP")) * Expc
       
172        .flags.ExpCount = .GiveEXP
       
174        .Veneno = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Veneno"))
       
176        .flags.Domable = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Domable"))
       
178        .GiveGLD = val(LeerNPCs.GetValue("NPC" & NpcNumber, "GiveGLD")) * Oroc
       
180        .PoderAtaque = val(LeerNPCs.GetValue("NPC" & NpcNumber, "PoderAtaque"))
182        .PoderEvasion = val(LeerNPCs.GetValue("NPC" & NpcNumber, "PoderEvasion"))
       
184        .InvReSpawn = val(LeerNPCs.GetValue("NPC" & NpcNumber, "InvReSpawn"))
       
186        With .Stats
188            .MaxHP = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MaxHP"))
190            .MinHP = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MinHP"))
192            .MaxHIT = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MaxHIT"))
194            .MinHIT = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MinHIT"))
196            .def = val(LeerNPCs.GetValue("NPC" & NpcNumber, "DEF"))
198        End With
       
200        .Invent.NroItems = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NROITEMS"))
        
202        For loopc = 1 To Npclist(npcindex).Invent.NroItems
204            ln = LeerNPCs.GetValue("NPC" & NpcNumber, "Obj" & loopc)
206               .Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
208               .Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
210        Next loopc
 

         
         
212        .flags.LanzaSpells = val(LeerNPCs.GetValue("NPC" & NpcNumber, "LanzaSpells"))

214        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)

216        For loopc = 1 To .flags.LanzaSpells
218            .Spells(loopc) = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Sp" & loopc))
220        Next loopc
       
222        If .NPCtype = eNPCType.Entrenador Then
224            .NroCriaturas = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NroCriaturas"))
226            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador

228            For loopc = 1 To .NroCriaturas
230                .Criaturas(loopc).npcindex = LeerNPCs.GetValue("NPC" & NpcNumber, "CI" & loopc)
234            Next loopc

236        End If
       
238        With .flags
240            .NPCActive = True
           
242            'If Respawn Then
244            '    .Respawn = val(LeerNPCs.GetValue("NPC" & NpcNumber, "RespawnTime"))
246            'Else
248            '    .Respawn = 1

250            ' End If
           
                .Respawn = 0 'Mermas temporal porque los npcs n respawneana
                
252            .BackUp = val(LeerNPCs.GetValue("NPC" & NpcNumber, "BackUp"))
256            .RespawnOrigPos = val(LeerNPCs.GetValue("NPC" & NpcNumber, "PosOrig"))
258            .AfectaParalisis = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Inmunidad"))
           
260            .Snd1 = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Snd1"))
262            .Snd2 = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Snd2"))
264            .Snd3 = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Snd3"))

266        End With
       
        'Tipo de items con los que comercia
268       .TipoItems = val(LeerNPCs.GetValue("NPC" & NpcNumber, "TipoItems"))
           
270     End With
   

    Exit Sub

ErrHandler:
    Call LogError("Error en ReloadNPCIndexByFile - Err: " & Err.Number & " " & Err.description)

End Sub

