Attribute VB_Name = "mMainLoop"
Option Explicit

Public prgRun As Boolean
Public LastGameTick As Long

Private Const GAME_TIMER_INTERVAL = 40
Public Sub Auditoria()

 
    On Error GoTo errhand
    
    Call PasarSegundo 'sistema de desconexion de 10 segs
    
    Static centinelSecs As Byte

    centinelSecs = centinelSecs + 1

   If centinelSecs = 5 Then
        'Every 5 seconds, we try to call the player's attention so it will report the code.
        'Call modCentinela.CallUserAttention
    
        centinelSecs = 0

    End If
    

    Exit Sub

errhand:

    Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)

    Resume Next

End Sub

Public Sub packetResend()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Attempts to resend to the user all data that may be enqueued.
    '***************************************************
    On Error GoTo ErrHandler:

    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

    Exit Sub

ErrHandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)

    Resume Next

End Sub

Public Sub TIMER_AI()
    On Error GoTo ErrorHandler

    Dim npcindex As Long
    Dim X        As Integer
    Dim Y        As Integer
    Dim UseAI    As Integer
    Dim Mapa     As Integer
    Dim e_p      As Integer

    'Barrin 29/9/03
    If Not EnPausa Then

        'Update NPCs
        For npcindex = 1 To LastNPC
        
            With Npclist(npcindex)

                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            
                    ' Chequea si contiua teniendo dueño
                    If .Owner > 0 Then Call ValidarPermanenciaNpc(npcindex)
            
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(npcindex)
                    Else

                            'Usamos AI si hay algun user en el mapa
                            If .flags.Inmovilizado = 1 Then
                                Call EfectoParalisisNpc(npcindex)

                            End If
                        
                            Mapa = .Pos.Map
                        
                            If Mapa > 0 Then
                                If MapInfo(Mapa).NumUsers > 0 Then
                                    If .Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(npcindex)

                                    End If

                                End If

                        End If

                    End If

                End If

            End With

        Next npcindex

    End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(npcindex).Name & " mapa:" & Npclist(npcindex).Pos.Map)
    Call MuereNpc(npcindex, 0)

End Sub
Public Sub GameTimer()

    Dim iUserIndex   As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean
    Dim DeltaTick    As Single
    
    DeltaTick = (GetTickCount - LastGameTick) / GAME_TIMER_INTERVAL
    LastGameTick = GetTickCount
    
    On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To MaxUsers 'LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then
                '¿User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    Call DoTileEvents(iUserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                    
                    If .flags.Muerto = 0 Then
                    
                        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then Call EfectoParalisisUser(iUserIndex, DeltaTick)
                        If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex, DeltaTick)

                        
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex, DeltaTick)
                        
                        If .flags.Meditando Then Call DoMeditar(iUserIndex, DeltaTick)
                        
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex, DeltaTick)
                        
                        If .flags.Incinerado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoIncinerado(iUserIndex, DeltaTick)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex, DeltaTick)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex, DeltaTick)
                        End If
                        
                        'check
                        If .flags.Trabajando Then DoTrabajar iUserIndex
                        
                        Call DuracionPociones(iUserIndex, DeltaTick)
      
                        Call HambreYSed(iUserIndex, DeltaTick, bEnviarAyS)
                        
                            If .flags.Hambre = 0 And .flags.Sed = 0 Then 'Si esta en 1 quiere decir que en realidad tiene el hambre en 0, es confuso xD Mermas
                            
                                If Not .flags.Descansar Then
                                
                                    Call Sanar(iUserIndex, DeltaTick, bEnviarStats, SanaIntervaloSinDescansar)
                                    
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                    Call RecStamina(iUserIndex, DeltaTick, bEnviarStats, 50)
                                    
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                Else
                                        'No esta descansando
                                        Call Sanar(iUserIndex, DeltaTick, bEnviarStats, StaminaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, DeltaTick, bEnviarStats, 50)

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False

                                        End If
                                        
                                        If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False
                                        End If
                                End If
                                       
                            End If 'fin .flags.Hambre = 0 And .flags.Sed = 0 Then
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex, DeltaTick)
                        
                    Else 'Muerto
                    
                        If .flags.Resucitando = True Then Call DoResucitar(iUserIndex)
                        
                    End If 'Muerto
                    
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + DeltaTick
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                    
                End If 'UserLogged
                 
                'Ya terminamos de procesar el paquete, sigamos recibiendo.
                '.Counters.PacketsTick = 0
                
            End If

        End With
 
    Next iUserIndex

    Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)

End Sub

