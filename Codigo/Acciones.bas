Attribute VB_Name = "Acciones"
Option Explicit

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    On Error GoTo Error_Err
    
    Dim tempIndex  As Integer
    
    With UserList(UserIndex)
    
        '¿Rango Visión? (ToxicWaste)
        If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then
            Exit Sub
        End If
    
        If .flags.Trabajando = True Then
            .flags.Trabajando = False
            Call WriteLocaleMsg(UserIndex, 391, vbNullString, 1, 12) 'Dejas de trabajar
        End If
        
        '¿Posicion valida?
        If InMapBounds(Map, X, Y) Then
        
             If .Invent.AnilloEqpSlot <> 0 Then          'Acciones de herramientas
                Select Case .Invent.AnilloEqpObjIndex
                    Case RED_PESCA, CAÑA_PESCA
                        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If HayAgua(Map, X, Y) Then
                            Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar...", 3)
                            .flags.Trabajando = True
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                    Case PIQUETE_MINERO
 
                        If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                            'Check distance
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
 
                            '¿Hay un yacimiento donde clickeo?
                            If ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otYacimiento Then
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar...", 2)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                    Case HACHA_LEÑADOR
 
                        If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            If MapInfo(.Pos.Map).Pk = False Then
                                Call WriteConsoleMsg(UserIndex, "No puedes talar en zona segura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            '¿Hay un arbol donde clickeo?
                            If ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otArboles Then
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar...", 2)
                            End If
                        End If
                        
                        
                    Case TIJERAS
                        
                        If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            If MapInfo(.Pos.Map).Pk = False Then
                                Call WriteConsoleMsg(UserIndex, "No puedes juntar raices en zona segura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                    
                            '¿Hay un arbol donde clickeo?
                            If ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otArboles Then
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar...", 4)
                            End If
                        End If
                        
                        
                    Case iMinerales.PlataCruda, iMinerales.HierroCrudo, iMinerales.OroCrudo
                        'Check there is a proper item there
                        If .flags.TargetObj > 0 Then
                            If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                                'Validate other items
                                If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                                    Exit Sub
                                End If
                                
                                ''chequeamos que no se zarpe duplicando oro
                                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                                    If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                        Call WriteConsoleMsg(UserIndex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub
                                    End If
                                    
                                    ''FUISTE
                                    Call WriteShowMessageBox(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                                    Call FlushBuffer(UserIndex)
                                    Call CloseSocket(UserIndex)
                                    Exit Sub
                                End If
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar...", 2)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        End If
                
                End Select
                
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otYunque Then
                        If .Invent.AnilloEqpObjIndex = MARTILLO_HERRERO Then
                             Call EnivarArmadurasConstruibles(UserIndex)
                             Call EnivarArmasConstruibles(UserIndex)
                             Call EnivarCascosConstruibles(UserIndex)
                             Call EnivarEscudosConstruibles(UserIndex)
                             Call WriteAbrirFormularios(UserIndex, 11) 'Herreria
                         End If
                    End If
                End If
                
             End If 'fin  .Invent.AnilloEqpSlot <> 0 Then
             
             If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otFragua Then
                    If .flags.Lingoteando <> 0 Then
                        .flags.Trabajando = True
                
                        Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
             
             If MapData(Map, X, Y).npcindex > 0 Then     'Acciones NPCs
             
                'Set the target NPC
                tempIndex = MapData(Map, X, Y).npcindex
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    
                    If DeadCheck(UserIndex) Then Exit Sub
                        
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                        Exit Sub
                    End If
                    
               '     If .flags.Privilegios And PlayerType.User Then
                '
                 '       'Es de una faccion en otro mapa? mermas
                      '  If (esCiuda(UserIndex) Or esArmada(UserIndex)) And (MapInfo(Map).battle_mode <> 2 And MapInfo(Map).battle_mode <> 3) Then
                  ''          Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 592, 4) 'No comercia con enemigos en la ciudad
                    '        Exit Sub
                     '   End If
                        
                        'Debug.Print MapInfo(Map).battle_mode & " " Time
                      '  If (esRepu(UserIndex) Or esMili(UserIndex)) And (MapInfo(Map).battle_mode <> 2 And MapInfo(Map).battle_mode <> 3) Then
                       '     Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 592, 4) 'No comercia con enemigos en la ciudad
                        '    Exit Sub
                       ' End If
                        
                     '   If (esRene(UserIndex) Or esCaos(UserIndex)) And (MapInfo(Map).battle_mode <> 3) Then
                      '      Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 592, 4) 'No comercia con enemigos en la ciudad
                       '     Exit Sub
                       ' End If
                    
                  '  End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)


                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Shop Then
                    
                    If DeadCheck(UserIndex) Then Exit Sub
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                        Exit Sub
                    End If
                       
                    Call WriteAbrirFormularios(UserIndex, 5) ' frmShop
                    
                 
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Convertidores Then
                
                    If DeadCheck(UserIndex) Then Exit Sub
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                        Exit Sub
                    End If

                    Select Case Npclist(tempIndex).flags.Status

                        Case 1 'Ciudadano
                            Call EntrarImperial(UserIndex)
                            
                        Case 2 'Republicano
                            Call EntrarRepublica(UserIndex)
                            
                    End Select

                ElseIf Npclist(tempIndex).NPCtype = eNPCType.facciones Then
                
                    If DeadCheck(UserIndex) Then Exit Sub
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                        Exit Sub
                    End If
                    
                    Select Case Npclist(tempIndex).flags.Status
                    
                        Case 1
                        
                            If esArmada(UserIndex) Then
                                Call ModFacciones.RecompensaArmadaReal(UserIndex)
                            Else
                                Call ModFacciones.EnlistarArmadaReal(UserIndex)
                            End If
                    
                        Case 2
                        
                            If esMili(UserIndex) Then
                                Call ModFacciones.RecompensaMilicia(UserIndex)
                            Else
                                Call ModFacciones.EnlistarMilicia(UserIndex)
                            End If
                    
                        Case 4
                            If esCaos(UserIndex) Then
                                Call ModFacciones.RecompensaCaos(UserIndex)
                            Else
                                Call ModFacciones.EnlistarCaos(UserIndex)
                            End If
                    
                    End Select
                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Entrenador Then
                
                    If DeadCheck(UserIndex) Then Exit Sub
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                       Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                       Exit Sub
                    End If

                    Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)

                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    
                    If DeadCheck(UserIndex) Then Exit Sub
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                       Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                       Exit Sub
                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(UserIndex, True)

                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Then
        
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteLocaleMsg(UserIndex, 8) ' Estas muy lejos.
                        Exit Sub
                    End If
                    
                    If .flags.Muerto = 1 And .flags.Resucitando = False Then
                        Call RevivirUsuario(UserIndex)
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(Npclist(tempIndex).Char.CharIndex, 11, 1))
                    
                    ElseIf .flags.Resucitando = False And .flags.Muerto = 0 Then
                    
                            If .Stats.MinHP < .Stats.MaxHP Then
                                .flags.Envenenado = 0
                                .flags.Incinerado = 0
                                .Stats.MinHP = .Stats.MaxHP
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SANAR, .Pos.X, .Pos.Y))
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(.Char.CharIndex, 133, 100, False, False))
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(Npclist(tempIndex).Char.CharIndex, 32, 1))
                                Call WriteUpdateHP(UserIndex)
                            End If
                    End If
                    
                    
                End If 'Fin NPCTypes
                    
             '¿Es un obj?
             ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                
                tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, UserIndex)

                    Case eOBJType.otCorreo 'Correo
                        Call AccionParaCorreo(Map, X, Y, UserIndex)
                    
                    Case eOBJType.otLeña    'Leña
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(Map, X, Y, UserIndex)
                        End If

                    Case eOBJType.otPozos   'Pozos
                        Call AccionParaPozos(Map, X, Y, UserIndex)
   
                End Select
                
                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
             ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
             
                .flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                
                Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                
                End Select
            
            
             ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
             
                .flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
        
                Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
       
                End Select
            
             ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                
                .flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                
                Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select

             End If
        End If
    
    End With

    Exit Sub

Error_Err:
    Call RegistrarError(Err.Number, Err.description, "Acciones.Accion", Erl)
    Resume Next
    
End Sub

Sub AccionParaPozos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error GoTo Error_Err

    With UserList(UserIndex)
    
        If Not (Distance(.Pos.X, .Pos.Y, X, Y) > 3) Then
            
            Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).SubTipo
            
                Case 1 'Pozo de maná
                
                    If .Stats.MinMAN < .Stats.MaxMAN Then
                        .Stats.MinMAN = .Stats.MaxMAN
                        Call WriteUpdateMana(UserIndex)
                        Call WriteLocaleMsg(UserIndex, 173) 'Has bebido bla bla...
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                    End If
                    
                Case 2 'Pozo de agua
                
                    If .Stats.MinAGU < .Stats.MaxAGU Then
                        .Stats.MinAGU = .Stats.MaxAGU
                        Call WriteUpdateSed(UserIndex)
                        Call WriteLocaleMsg(UserIndex, 175) 'Has bebido bla bla...
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                    End If
                    
            End Select
            
        Else
            Call WriteLocaleMsg(UserIndex, 8)
        End If
    
    End With

    Exit Sub

Error_Err:
    Call RegistrarError(Err.Number, Err.description, "Acciones.AccionParaPozos", Erl)
    Resume Next
    
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
    
    On Error GoTo Error_Err

    If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                'Abre la puerta
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(X, Y, MapData(Map, X, Y).ObjInfo.ObjIndex, 0))
                    
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, Map, X, Y, 0)
                    Call Bloquear(True, Map, X - 1, Y, 0)
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteLocaleMsg(UserIndex, 172)
                End If

            Else
                'Cierra puerta
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(X, Y, MapData(Map, X, Y).ObjInfo.ObjIndex, 0))
                                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                Call Bloquear(True, Map, X - 1, Y, 1)
                Call Bloquear(True, Map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

            End If
        
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        Else
            Call WriteLocaleMsg(UserIndex, 172)

        End If

    Else
        Call WriteLocaleMsg(UserIndex, 8)
        
    End If

    Exit Sub

Error_Err:
    Call RegistrarError(Err.Number, Err.description, "Acciones.AccionParaPozos", Erl)
    Resume Next
    
End Sub

Sub AccionParaRamita(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim Suerte As Byte
    Dim exito  As Byte
    Dim Obj    As Obj

    Dim Pos    As WorldPos
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    With UserList(UserIndex)

        If Distancia(Pos, .Pos) > 2 Then
             Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If
    
        If MapData(Map, X, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
             Call WriteLocaleMsg(UserIndex, 191)
            Exit Sub

        End If
    
        If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
        ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
        ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
            Suerte = 1

        End If
    
        exito = RandomNumber(1, Suerte)
    
        If exito = 1 Then
            If MapInfo(.Pos.Map).Zona <> Ciudad Then
                Obj.ObjIndex = FOGATA
                Obj.Amount = 1

                Call WriteLocaleMsg(UserIndex, 170)
                 
                Call MakeObj(Obj, Map, X, Y)
            
                'Las fogatas prendidas se deben eliminar
                Dim Fogatita As New cGarbage
                Fogatita.Map = Map
                Fogatita.X = X
                Fogatita.Y = Y
                Call TrashCollector.Add(Fogatita)
            
                Call SubirSkill(UserIndex, eSkill.Supervivencia)
            Else
                Call WriteLocaleMsg(UserIndex, 171)
                Exit Sub

            End If

        Else
            Call WriteLocaleMsg(UserIndex, 8)

        End If

    End With

End Sub

Sub AccionParaCorreo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

On Error Resume Next

Dim Pos As WorldPos

Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 3 Then

    Call WriteLocaleMsg(UserIndex, 8)
    
  Exit Sub
  
End If

If UserList(UserIndex).flags.Muerto = 1 Then

  Call WriteLocaleMsg(UserIndex, 77)
  
 Exit Sub
 
End If
Call WriteCorreoList(UserIndex)
Call WriteAbrirFormularios(UserIndex, 6)
UserList(UserIndex).flags.RecibioCorreo = 0
Call WriteMensajeSigno(UserIndex, 0)

End Sub

