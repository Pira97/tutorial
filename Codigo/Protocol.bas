Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit
Public CantdPaquetes      As Long

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer  As New clsByteQueue


Private Enum ServerPacketID
    LoggedSuccessful        ' 1
    Logged                  ' 2
    RemoveDialogs           ' 3
    RemoveCharDialog        ' 4
    NavigateToggle          ' 5
    MontateToggle           ' 6
    Disconnect              ' 7
    CommerceEnd             ' 8
    BankEnd                 ' 9
    CommerceInit            ' 10
    BankInit                ' 11
    UpdateSta               ' 16
    UpdateMana              ' 17
    UpdateHP                ' 18
    UpdateGold              ' 19
    UpdateExp               ' 21
    ChangeMap               ' 22
    PosUpdate               ' 23
    ChatOverHead            ' 24
    ChatOverHeadLocale      ' 25
    ConsoleMsg              ' 26
    GuildChat               ' 27
    ShowMessageBox          ' 28
    UserIndexInServer       ' 29
    UserCharIndexInServer   ' 30
    CharacterCreate         ' 31
    CharacterRemove         ' 32
    CharacterMove           ' 34
    ForceCharMove           ' 35
    CharacterChange         ' 36
    CharacterChangeSlot     ' 37
    ObjectCreate            ' 38
    ObjectDelete            ' 39
    BlockPosition           ' 40
    PlayMidi                ' 41
    PlayWave                ' 42
    guildList               ' 43
    AreaChanged             ' 44
    PauseToggle             ' 45
    RainToggle              ' 46
    CreateFX                ' 47
    UpdateUserStats         ' 48
    UpdateUserStatsForLevel
    WorkRequestTarget       ' 49
    ChangeInventorySlot     ' 50
    ChangeBankSlot          ' 51
    ChangeSpellSlot         ' 52
    atributes               ' 53
    BlacksmithWeapons       ' 54
    BlacksmithArmors        ' 55
    BlacksmithHelmet        ' 56
    BlacksmithShield        ' 57
    CarpenterObjects        ' 58
    SastreObjects           ' 59
    AlquimiaObjects         ' 60
    RestOK                  ' 61
    SendMsgBox                ' 62
    Blind                   ' 63
    Dumb                    ' 64
    ChangeNPCInventorySlot  ' 66
    UpdateHungerAndThirst   ' 67
    MiniStats               ' 68
    LevelUp                 ' 69
    SetInvisible            ' 70
    MeditateToggle          ' 71
    BlindNoMore             ' 72
    DumbNoMore              ' 73
    SendSkills              ' 74
    TrainerCreatureList     ' 75
    guildNews               ' 76
    OfferDetails            ' 77
    AlianceProposalsList    ' 78
    PeaceProposalsList      ' 79
    CharacterInfo           ' 80
    GuildLeaderInfo         ' 81
    GuildMemberInfo         ' 82
    GuildDetails            ' 83
    ParalizeOK              ' 84
    ShowUserRequest         ' 85
    TradeOK                 ' 86
    BankOK                  ' 87
    Pong                    ' 89
    UpdateTagAndStatus      ' 90
    LocaleMsg               ' 91
    
    'GM messages
    ShowSOSForm             ' 93
    UserNameList            ' 94
    CorreoList               ' 95
    UpdateStrenght          ' 96
    UpdateDexterity         ' 97
    Premios                 ' 100
    EfectoCharParticula     ' 101
    AddPJ                   ' 103
    EfectoTerrenoParticula  ' 105
    EfectoTerrenoFX         ' 106
    CharStatus              ' 107
    MensajeSigno            ' 108
    MarcamosSkin            ' 112
    MostrarUbicacion        ' 1f13
    CargarSkin              ' 114
    CharMsgStatus           ' 115
    CharMsgStatusNPC        ' 116
    AbrirFormularios        ' 117
    ChangeInventorySlotUser ' 118
    AuraToChar              ' 119
    UpdateSed
    UpdateHambre
    EjecutarAccion
End Enum

Private Enum ClientPacketID
    Walk                    '5
    LoginExistingChar       '0
    LoginNewChar            '1
    Talk                    '3
    Whisper                 '4
    RequestPositionUpdate   '6
    attack                  '7
    PickUp                  '8
    CombatModeToggle        '9
    ResuscitationSafeToggle '11
    RequestGuildLeaderInfo  '12
    RequestAtributes        '14
    RequestSkills           '15
    RequestMiniStats        '16
    CommerceEnd             '17
    BankEnd                 '21
    Drop                    '24
    DropDestroy             '25
    CastSpell               '26
    LeftClick               '27
    DoubleClick             '28
    Work                    '29
    UseItem                 '30
    CraftBlacksmith         '31
    CraftCarpenter          '32
    Craftalquimia           '33
    CraftSastre             '34
    WorkLeftClick           '35
    CreateNewGuild          '36
    EquipItem               '37
    EquiparSkin             '38
    ChangeHeading           '39
    ModifySkills            '40
    Train                   '41
    CommerceBuy             '42
    BankExtractItem         '43
    CommerceSell            '44
    BankDeposit             '45
    MoveSpell               '46
    MoveBank                '47
    ClanCodexUpdate         '48
    GuildAcceptPeace        '50
    GuildRejectAlliance     '51
    GuildRejectPeace        '52
    GuildAcceptAlliance     '53
    GuildOfferPeace         '54
    GuildOfferAlliance      '55
    GuildAllianceDetails    '56
    GuildPeaceDetails       '57
    GuildRequestJoinerInfo  '58
    GuildAlliancePropList   '59
    GuildPeacePropList      '60
    GuildDeclareWar         '61
    GuildNewWebsite         '62
    GuildAcceptNewMember    '63
    GuildRejectNewMember    '64
    GuildKickMember         '65
    GuildUpdateNews         '66
    GuildMemberInfo         '67
    GuildOpenElections      '68
    GuildRequestMembership  '69
    GuildRequestDetails     '70
    Online                  '71
    Quit                    '72
    GuildLeave              '73
    Rest                    '75
    ConnectAccount          '76
    CreateNewAccount        '77
    Meditate                '78
    Resucitate              '79
    RequestStats            '80
    CommerceStart           '81
    BankStart               '82
    Enlist                  '83
    Information             '84
    Reward                  '85
    UpTime                  '86
    GuildMessage            '87
    CentinelReport          '88
    GuildOnline             '89
    GMRequest               '90
    ChangeDescription       '91
    GuildVote               '92
    Gamble                  '93
    BankExtractGold         '94
    BankDepositGold         '95
    Denounce                '96
    PidePremios             '97
    RPremios                '98
    GuildFundate            '99
    GuildFundation          '100
    Ping                    '101
    GMCommands              '102
    InitCrafting            '103
    ShowGuildNews           '104
    SwapObjects             '105
    Packets_Correo          '106
    EnviarCorreo            '107
    RetirarFaccion          '109
    RegresarHogar            '110
    ParticulaUsuario        '111
    ProcesosLogin         '112
    TransferGOLD            '113
    SeleccionarHogar        '114
    Casamiento                '115
    divorciar               '116
    HayEventos              '117
    CloseGuild              '118
    AddAmigos               '119
    DelAmigos               '120
    OnAmigos                '121
    MsgAmigos               '122
    AbrirForms              '123
    DesconectarCuenta
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 128

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_fonttt
    FONTTYPE_FIGHTNPC
    FONTTYPE_INFOBOLD2
    FONTTYPE_INFOBOLD3
    FONTTYPE_INFOBOLD4
    FONTTYPE_PALABRASMAGICAS
    FONTTYPE_fontINiCIO
    FONTTYPE_INFOBOLD5
    FONTTYPE_oro
    FONTTYPE_LETRADIOS
    FONTTYPE_LETRASEMIDIOS
    FONTTYPE_IMPERIAL
    FONTTYPE_RENEGADO
    FONTTYPE_REPUBLICANO
    FONTTYPE_MILICIANO
    FONTTYPE_FUERZASCAOS
    FONTTYPE_LETRACONSEJERO
    FONTTYPE_GRITAR
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
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean

100     On Error GoTo HandleError_Err

102    'Contamos cuantos paquetes recibimos.
104    'UserList(UserIndex).Counters.PacketsTick = UserList(UserIndex).Counters.PacketsTick + 1
    
106    'Comento esto por ahora, por que cuando hago worldsave, envia mas paquetes en 40ms
108    'y desconecta al pj, hay que reveer que hacer con esto y como solucionarlo.

109    'Si recibis 10 paquetes en 40ms (intervalo del GameTimer), cierro la conexion.
110    'If UserList(UserIndex).Counters.PacketsTick > 10 Then
112        'Debug.Print "Muchos paquetes"
114        'Call CloseSocket(UserIndex)
        
116       ' Exit Function

118    'End If
    
120     Dim packetID As Long
     
122     packetID = CLng(UserList(UserIndex).incomingData.PeekByte())

124    'Does the packet requires a logged user??
126    If Not (packetID = ClientPacketID.LoginExistingChar _
        Or packetID = ClientPacketID.LoginNewChar _
        Or packetID = ClientPacketID.ConnectAccount _
        Or packetID = ClientPacketID.CreateNewAccount _
        Or packetID = ClientPacketID.ProcesosLogin) Then

        
128        'Is the user actually logged?
130        If Not UserList(UserIndex).flags.UserLogged Then
132            Call CloseSocket(UserIndex)
134            Exit Function
        
136        'He is logged. Reset idle counter if id is valid.
138        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
140            UserList(UserIndex).Counters.IdleCount = 0

142        End If

144    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
    
146        UserList(UserIndex).Counters.IdleCount = 0
        
148        'Is the user logged?
150        If UserList(UserIndex).flags.UserLogged Then
152            Call CloseSocket(UserIndex)
154            Exit Function

156        End If

158    End If
    
160    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
162    UserList(UserIndex).flags.NoPuedeSerAtacado = False

    #If Debugger = 1 Then
        If frmRecibeDatos.Visible = True Then Call AddtoRichTextBox("El cliente mandó el paquete Nº " & packetID & " (" & UserList(UserIndex).incomingData.length & " Bytes)" & " - " & Time)
    #End If
    
    Select Case packetID
         
           Case ClientPacketID.Walk                    'M
166            Call HandleWalk(UserIndex)
            
168        Case ClientPacketID.LoginExistingChar       'OLOGIN
170            Call HandleLoginExistingChar(UserIndex)
        
172        Case ClientPacketID.LoginNewChar            'NLOGIN
174            Call HandleLoginNewChar(UserIndex)

176        Case ClientPacketID.Talk                    ';
178            Call HandleTalk(UserIndex)
        
180         Case ClientPacketID.Whisper                 '\
182            Call HandleWhisper(UserIndex)
        
184        Case ClientPacketID.RequestPositionUpdate   'RPU
186            Call HandleRequestPositionUpdate(UserIndex)
        
188        Case ClientPacketID.attack                  'AT
190            Call HandleAttack(UserIndex)
        
192        Case ClientPacketID.PickUp                  'AG
194            Call HandlePickUp(UserIndex)
            
196        Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
198            Call HanldeCombatModeToggle(UserIndex)

204        Case ClientPacketID.ResuscitationSafeToggle
206            Call HandleResuscitationToggle(UserIndex)
        
208        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
210            Call HandleRequestGuildLeaderInfo(UserIndex)

216        Case ClientPacketID.RequestAtributes        'ATR
218            Call HandleRequestAtributes(UserIndex)
     
220        Case ClientPacketID.RequestSkills           'ESKI
222            Call HandleRequestSkills(UserIndex)
        
224        Case ClientPacketID.RequestMiniStats        'FEST
226            Call HandleRequestMiniStats(UserIndex)
        
228        Case ClientPacketID.CommerceEnd             'FINCOM
230            Call HandleCommerceEnd(UserIndex)

232        Case ClientPacketID.BankEnd                 'FINBAN
234            Call HandleBankEnd(UserIndex)

236        Case ClientPacketID.Drop                    'TI
238            Call HandleDrop(UserIndex)
        
240        Case ClientPacketID.DropDestroy             'TI
242            Call HandleDropDestroy(UserIndex)
        
244        Case ClientPacketID.CastSpell               'LH
246            Call HandleCastSpell(UserIndex)
        
248        Case ClientPacketID.LeftClick               'LC
250            Call HandleLeftClick(UserIndex)
        
252        Case ClientPacketID.DoubleClick             'RC
254            Call HandleDoubleClick(UserIndex)
        
256        Case ClientPacketID.Work                    'UK
258            Call HandleWork(UserIndex)
            
260        Case ClientPacketID.UseItem                 'USA
262            Call HandleUseItem(UserIndex)
        
264        Case ClientPacketID.CraftBlacksmith         'CNS
266            Call HandleCraftBlacksmith(UserIndex)
        
268        Case ClientPacketID.CraftCarpenter          'CNC
270            Call HandleCraftCarpenter(UserIndex)
        
272        Case ClientPacketID.CraftSastre
274            Call HandleCraftSastre(UserIndex)
        
276        Case ClientPacketID.Craftalquimia
278            Call HandleCraftalquimia(UserIndex)
        
280        Case ClientPacketID.WorkLeftClick           'WLC
282            Call HandleWorkLeftClick(UserIndex)
        
284        Case ClientPacketID.CreateNewGuild          'CIG
286            Call HandleCreateNewGuild(UserIndex)

288        Case ClientPacketID.EquipItem               'EQUI
290            Call HandleEquipItem(UserIndex)
            
292        Case ClientPacketID.EquiparSkin
294            Call HandleEquiparSkin(UserIndex)
        
296        Case ClientPacketID.ChangeHeading           'CHEA
298            Call HandleChangeHeading(UserIndex)
        
300        Case ClientPacketID.ModifySkills            'SKSE
302            Call HandleModifySkills(UserIndex)
        
304        Case ClientPacketID.Train                   'ENTR
306            Call HandleTrain(UserIndex)
        
308        Case ClientPacketID.CommerceBuy             'COMP
310            Call HandleCommerceBuy(UserIndex)
        
312        Case ClientPacketID.BankExtractItem         'RETI
314            Call HandleBankExtractItem(UserIndex)
        
316        Case ClientPacketID.CommerceSell            'VEND
318            Call HandleCommerceSell(UserIndex)
        
320        Case ClientPacketID.BankDeposit             'DEPO
322            Call HandleBankDeposit(UserIndex)
            
324        Case ClientPacketID.MoveSpell               'DESPHE
326            Call HandleMoveSpell(UserIndex)
            
328        Case ClientPacketID.MoveBank
330            Call HandleMoveBank(UserIndex)
        
332        Case ClientPacketID.ClanCodexUpdate         'DESCOD
334            Call HandleClanCodexUpdate(UserIndex)

336        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
338            Call HandleGuildAcceptPeace(UserIndex)
        
340        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
342            Call HandleGuildRejectAlliance(UserIndex)
        
344        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
346            Call HandleGuildRejectPeace(UserIndex)
        
348        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
350            Call HandleGuildAcceptAlliance(UserIndex)
        
352        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
354            Call HandleGuildOfferPeace(UserIndex)
        
356        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
358            Call HandleGuildOfferAlliance(UserIndex)
        
360        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
362            Call HandleGuildAllianceDetails(UserIndex)
        
364        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
366            Call HandleGuildPeaceDetails(UserIndex)
        
368        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
370            Call HandleGuildRequestJoinerInfo(UserIndex)
        
372        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
374            Call HandleGuildAlliancePropList(UserIndex)
        
376        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
378            Call HandleGuildPeacePropList(UserIndex)
        
380        Case ClientPacketID.GuildDeclareWar         'DECGUERR
382            Call HandleGuildDeclareWar(UserIndex)
        
383        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
384            Call HandleGuildNewWebsite(UserIndex)
        
385        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
386            Call HandleGuildAcceptNewMember(UserIndex)
        
387        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
388            Call HandleGuildRejectNewMember(UserIndex)
        
389        Case ClientPacketID.GuildKickMember         'ECHARCLA
390            Call HandleGuildKickMember(UserIndex)
        
391        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
392            Call HandleGuildUpdateNews(UserIndex)
        
393        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
394            Call HandleGuildMemberInfo(UserIndex)
        
395        Case ClientPacketID.GuildOpenElections      'ABREELEC
396            Call HandleGuildOpenElections(UserIndex)
        
397        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
398            Call HandleGuildRequestMembership(UserIndex)
        
399        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
400            Call HandleGuildRequestDetails(UserIndex)
                  
401        Case ClientPacketID.Online                  '/ONLINE
402            Call HandleOnline(UserIndex)
        
403        Case ClientPacketID.Quit                    '/SALIR
404            Call HandleQuit(UserIndex)
            
405        Case ClientPacketID.GuildLeave              '/SALIRCLAN
406            Call HandleGuildLeave(UserIndex)

409         Case ClientPacketID.Rest                    '/DESCANSAR
410            Call HandleRest(UserIndex)
            
411        Case ClientPacketID.ConnectAccount          'Conectamos la cuenta
412            Call HandleLoginAccount(UserIndex)
            
413        Case ClientPacketID.CreateNewAccount        'Creamos la cuenta
414            Call HandleLoginNewAccount(UserIndex)

415        Case ClientPacketID.Meditate                '/MEDITAR
416            Call HandleMeditate(UserIndex)
 
417        Case ClientPacketID.Resucitate              '/RESUCITAR
418            Call HandleResucitate(UserIndex)

419        Case ClientPacketID.RequestStats            '/EST
420            Call HandleRequestStats(UserIndex)
        
421        Case ClientPacketID.CommerceStart           '/COMERCIAR
422            Call HandleCommerceStart(UserIndex)
        
423        Case ClientPacketID.BankStart               '/BOVEDA
424            Call HandleBankStart(UserIndex)
        
425        Case ClientPacketID.Enlist                  '/ENLISTAR
426            Call HandleEnlist(UserIndex)
        
427        Case ClientPacketID.Information             '/INFORMACION
428            Call HandleInformation(UserIndex)
        
429        Case ClientPacketID.Reward                  '/RECOMPENSA
430            Call HandleReward(UserIndex)
        
431        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)

432        Case ClientPacketID.GuildMessage            '/CMSG
433            Call HandleGuildMessage(UserIndex)

434        Case ClientPacketID.CentinelReport          '/CENTINELA
435            Call HandleCentinelReport(UserIndex)
        
436        Case ClientPacketID.GuildOnline             '/ONLINECLAN
437            Call HandleGuildOnline(UserIndex)

438        Case ClientPacketID.GMRequest               '/GM
439            Call HandleGMRequest(UserIndex)

440        Case ClientPacketID.ChangeDescription       '/DESC
442            Call HandleChangeDescription(UserIndex)
        
443        Case ClientPacketID.GuildVote               '/VOTO
444            Call HandleGuildVote(UserIndex)
            
445        Case ClientPacketID.Gamble                  '/APOSTAR
446            Call HandleGamble(UserIndex)
        
447        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
448            Call HandleBankExtractGold(UserIndex)
        
449        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
450            Call HandleBankDepositGold(UserIndex)
        
451        Case ClientPacketID.Denounce                '/DENUNCIAR
452            Call HandleDenounce(UserIndex)
        
453        Case ClientPacketID.PidePremios
454            Call HandlePremiosRequest(UserIndex)
        
455        Case ClientPacketID.RPremios
456            Call HandleRPremios(UserIndex)
        
457        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
458            Call HandleGuildFundate(UserIndex)
            
459        Case ClientPacketID.GuildFundation
460            Call HandleGuildFundation(UserIndex)
    
461        Case ClientPacketID.Ping                    '/PING
462            Call HandlePing(UserIndex)
            
463        Case ClientPacketID.GMCommands              'GM Messages
464            Call HandleGMCommands(UserIndex)
            
465        Case ClientPacketID.InitCrafting
466            Call HandleInitCrafting(UserIndex)
        
467        Case ClientPacketID.ShowGuildNews
468            Call HandleShowGuildNews(UserIndex)

469        Case ClientPacketID.SwapObjects
470            Call HandleSwapObjects(UserIndex)
            
471        Case ClientPacketID.Packets_Correo
472            Call HandlePacketsCorreo(UserIndex)
           
473        Case ClientPacketID.EnviarCorreo
474            Call HandleEnviarCorreo(UserIndex)
         
477        Case ClientPacketID.RetirarFaccion
478            Call HandleRetirarFaccion(UserIndex)
                            
479        Case ClientPacketID.RegresarHogar                    '/HOGAR
480            Call HandleRegresarHogar(UserIndex)
                   
481        Case ClientPacketID.ParticulaUsuario
482            Call HandleParticulaUsuario(UserIndex)
                    
483        Case ClientPacketID.ProcesosLogin
484            Call HandleProcesosLogin(UserIndex)

485        Case ClientPacketID.TransferGOLD
486            Call HandleTransferGOLD(UserIndex)
            
487        Case ClientPacketID.SeleccionarHogar
488            Call HandleSeleccionarHogar(UserIndex)
            
489        Case ClientPacketID.Casamiento
490            Call HandleCasamiento(UserIndex)
 
491        Case ClientPacketID.divorciar
492            Call handleDivorciar(UserIndex)

493        Case ClientPacketID.HayEventos
494            Call HandleHayEventos(UserIndex)

495        Case ClientPacketID.CloseGuild           '/CERRARCLAN
496            Call HandleCloseGuild(UserIndex)
                                                        
497        Case ClientPacketID.AddAmigos
498            Call HandleAddAmigo(UserIndex)

499        Case ClientPacketID.DelAmigos
500            Call HandleDelAmigo(UserIndex)
        
501        Case ClientPacketID.OnAmigos
502            Call HandleOnAmigo(UserIndex)
        
503        Case ClientPacketID.MsgAmigos
504            Call HandleMsgAmigo(UserIndex)

505        Case ClientPacketID.AbrirForms                    '/SALIR
506            Call HandleAbrirForms(UserIndex)
                
507        Case ClientPacketID.DesconectarCuenta
508            Call HandleDesconectarCuenta(UserIndex)
            
509        Case Else
        
510            'ERROR : Abort!
511            Call CloseSocket(UserIndex)

    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
512    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then

513        HandleIncomingData = True
  
514    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then

        'An error ocurred, log it and kick player.
515        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.Source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
                        
516        Call CloseSocket(UserIndex)
517        HandleIncomingData = False

518    Else

519        'Flush buffer - send everything that has been written

520        Call FlushBuffer(UserIndex)
521        HandleIncomingData = False

522    End If
    
    
    Exit Function
    
HandleError_Err:

523    Call RegistrarError(Err.Number, Err.description, "Protocol.HandleIncomingData", Erl)
524    Resume Next

End Function

Private Sub HandleGMCommands(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

100    Dim Command As Long


102    With UserList(UserIndex)
    
104        Call .incomingData.ReadByte
    
106        Command = CLng(.incomingData.PeekByte)
    
           Select Case Command

               Case eGMCommands.GMMessage                '/GMSG
110                Call HandleGMMessage(UserIndex)
        
112            Case eGMCommands.showName                '/SHOWNAME
114                Call HandleShowName(UserIndex)

116            Case eGMCommands.GoNearby                '/IRCERCA
118                Call HandleGoNearby(UserIndex)
        
120            Case eGMCommands.comment                 '/REM
122                Call HandleComment(UserIndex)
        
124            Case eGMCommands.serverTime              '/HORA
126                Call HandleServerTime(UserIndex)
        
128            Case eGMCommands.Where                   '/DONDE
130                Call HandleWhere(UserIndex)
        
132            Case eGMCommands.CreaturesInMap          '/NENE
134                Call HandleCreaturesInMap(UserIndex)

136            Case eGMCommands.WarpChar                '/TELEP
138                Call HandleWarpChar(UserIndex)

140            Case eGMCommands.SOSShowList             '/SHOW SOS
142                Call HandleSOSShowList(UserIndex)
        
144            Case eGMCommands.SOSRemove               'SOSDONE
146                Call HandleSOSRemove(UserIndex)
        
148            Case eGMCommands.GoToChar                '/IRA
150                Call HandleGoToChar(UserIndex)
        
152            Case eGMCommands.Invisible               '/INVISIBLE
154                Call HandleInvisible(UserIndex)
        
156            Case eGMCommands.GMPanel                 '/PANELGM
158                Call HandleGMPanel(UserIndex)
        
160            Case eGMCommands.RequestUserList         'LISTUSU
162                Call HandleRequestUserList(UserIndex)
        
164            Case eGMCommands.Working                 '/TRABAJANDO
166                Call HandleWorking(UserIndex)
        
168            Case eGMCommands.Hiding                  '/OCULTANDO
170                Call HandleHiding(UserIndex)
        
172            Case eGMCommands.Jail                    '/CARCEL
174                Call HandleJail(UserIndex)
        
176            Case eGMCommands.KillNPC                 '/RMATA
178                Call HandleKillNPC(UserIndex)
        
180            Case eGMCommands.WarnUser                '/ADVERTENCIA
182                Call HandleWarnUser(UserIndex)
        
184            Case eGMCommands.EditChar                '/MOD
186                Call HandleEditChar(UserIndex)
        
188            Case eGMCommands.RequestCharInfo         '/INFO
190                Call HandleRequestCharInfo(UserIndex)

192            Case eGMCommands.RequestCharInventory    '/INV
194                Call HandleRequestCharInventory(UserIndex)
        
196            Case eGMCommands.RequestCharBank         '/BOV
198                Call HandleRequestCharBank(UserIndex)
        
200            Case eGMCommands.RequestCharSkills       '/SKILLS
202                Call HandleRequestCharSkills(UserIndex)
        
204            Case eGMCommands.ReviveChar              '/REVIVIR
206                Call HandleReviveChar(UserIndex)
        
208            Case eGMCommands.OnlineGM                '/ONLINEGM
210                Call HandleOnlineGM(UserIndex)
        
212            Case eGMCommands.OnlineMap               '/ONLINEMAP
214                Call HandleOnlineMap(UserIndex)

216            Case eGMCommands.Kick                    '/ECHAR
218                Call HandleKick(UserIndex)
        
220            Case eGMCommands.Execute                 '/EJECUTAR
222                Call HandleExecute(UserIndex)
        
224            Case eGMCommands.BanChar                 '/BAN
226                Call HandleBanChar(UserIndex)

228            Case eGMCommands.NPCFollow               '/SEGUIR
230                Call HandleNPCFollow(UserIndex)
        
232            Case eGMCommands.SummonChar              '/SUM
234                Call HandleSummonChar(UserIndex)

244            Case eGMCommands.ResetNPCInventory       '/RESETINV
246                Call HandleResetNPCInventory(UserIndex)
        
248            Case eGMCommands.CleanWorld              '/LIMPIAR
250                Call HandleCleanWorld(UserIndex)
        
252            Case eGMCommands.ServerMessage           '/RMSG
254                Call HandleServerMessage(UserIndex)
        
256            Case eGMCommands.NickToIP                '/NICK2IP
258                Call HandleNickToIP(UserIndex)
        
260            Case eGMCommands.IPToNick                '/IP2NICK
262                Call HandleIPToNick(UserIndex)
        
264            Case eGMCommands.GuildOnlineMembers      '/ONCLAN
266                Call HandleGuildOnlineMembers(UserIndex)
        
268            Case eGMCommands.TeleportCreate ' /CT
                    Call HandleTeleportCreate(UserIndex)
        
272            Case eGMCommands.TeleportDestroy         '/DT
274                Call HandleTeleportDestroy(UserIndex)
            
276            Case eGMCommands.RainToggle              '/LLUVIA
278                Call HandleRainToggle(UserIndex)

280            Case eGMCommands.TalkAsNPC               '/TALKAS
282                Call HandleTalkAsNPC(UserIndex)
        
284            Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
286                Call HandleDestroyAllItemsInArea(UserIndex)

288            Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
290                Call HandleMakeDumbNoMore(UserIndex)

292            Case eGMCommands.SetTrigger              '/TRIGGER
294                Call HandleSetTrigger(UserIndex)
        
296            Case eGMCommands.AskTrigger              '/TRIGGER with no args
298                Call HandleAskTrigger(UserIndex)
        
300            Case eGMCommands.BannedIPList            '/BANIPLIST
302                Call HandleBannedIPList(UserIndex)
        
304            Case eGMCommands.BannedIPReload          '/BANIPRELOAD
306                Call HandleBannedIPReload(UserIndex)
        
308            Case eGMCommands.GuildMemberList         '/MIEMBROSCLAN
310                Call HandleGuildMemberList(UserIndex)
        
314            Case eGMCommands.GuildBan                '/BANCLAN
316                Call HandleGuildBan(UserIndex)
        
318            Case eGMCommands.BanIP                   '/BANIP
320                Call HandleBanIP(UserIndex)
        
322            Case eGMCommands.UnbanIP                 '/UNBANIP
324                Call HandleUnbanIP(UserIndex)
        
326            Case eGMCommands.CreateItem              '/CI
328                Call HandleCreateItem(UserIndex)
        
330            Case eGMCommands.DestroyItems            '/DEST
332                Call HandleDestroyItems(UserIndex)
  
334            Case eGMCommands.TileBlockedToggle       '/BLOQ
336                Call HandleTileBlockedToggle(UserIndex)
        
338            Case eGMCommands.KillNPCNoRespawn        '/MATA
340                Call HandleKillNPCNoRespawn(UserIndex)
        
342            Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
344                Call HandleKillAllNearbyNPCs(UserIndex)
        
346            Case eGMCommands.LastIP                  '/LASTIP
348                Call HandleLastIP(UserIndex)
        
350            Case eGMCommands.SystemMessage           '/SMSG
352                Call HandleSystemMessage(UserIndex)
        
354            Case eGMCommands.CreateNPC               '/ACC
356                Call HandleCreateNPC(UserIndex)
        
358            Case eGMCommands.CreateNPCWithRespawn    '/RACC
360                Call HandleCreateNPCWithRespawn(UserIndex)
        
362            Case eGMCommands.NavigateToggle          '/NAVE
364                Call HandleNavigateToggle(UserIndex)
           
366            Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
368                Call HandleServerOpenToUsersToggle(UserIndex)
        
370            Case eGMCommands.TurnOffServer           '/APAGAR
372                Call HandleTurnOffServer(UserIndex)

374            Case eGMCommands.RemoveCharFromGuild     '/RAJARCLAN
376                Call HandleRemoveCharFromGuild(UserIndex)

378            Case eGMCommands.AlterPassword           '/APASS
380                Call HandleAlterPassword(UserIndex)

382            Case eGMCommands.ToggleCentinelActivated '/CENTINELAACTIVADO
384                Call HandleToggleCentinelActivated(UserIndex)
 
386             Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
388                Call HandleShowGuildMessages(UserIndex)
        
390            Case eGMCommands.SaveMap                 '/GUARDAMAPA
392                Call HandleSaveMap(UserIndex)
        
394            Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
396                Call HandleChangeMapInfoPK(UserIndex)
        
398            Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
400                Call HandleChangeMapInfoBackup(UserIndex)
        
402            Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
404                Call HandleChangeMapInfoRestricted(UserIndex)
        
406            Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
408                Call HandleChangeMapInfoNoMagic(UserIndex)
        
410            Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
412                Call HandleChangeMapInfoNoInvi(UserIndex)
        
414            Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
416                Call HandleChangeMapInfoNoResu(UserIndex)
        
418            Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
420                Call HandleChangeMapInfoLand(UserIndex)
        
422            Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
424                Call HandleChangeMapInfoZone(UserIndex)
        
426            Case eGMCommands.SaveChars               '/GRABAR
428                Call HandleSaveChars(UserIndex)
        
430            Case eGMCommands.CleanSOS                '/BORRAR SOS
432                Call HandleCleanSOS(UserIndex)
        
434            Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
436                Call HandleKickAllChars(UserIndex)
        
438            Case eGMCommands.ReloadNPCs              '/RELOADNPCS
440                Call HandleReloadNPCs(UserIndex)
        
442            Case eGMCommands.ReloadServerIni         '/RELOADSINI
444                Call HandleReloadServerIni(UserIndex)
        
446            Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
448                Call HandleReloadSpells(UserIndex)
        
450            Case eGMCommands.ReloadObjects           '/RELOADOBJ
452                Call HandleReloadObjects(UserIndex)

454            Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
456                Call HandleSetIniVar(UserIndex)

458           Case eGMCommands.DARPUN                  'DARPUN
460                Call HandleDARPUN(UserIndex)

462            Case eGMCommands.ResponderGM
464                Call HandleResponderGM(UserIndex)
  
466            Case eGMCommands.Donador
468                Call HandleDonador(UserIndex)

470            Case eGMCommands.EventoOro
472                Call handleEventoOro(UserIndex)
            
474            Case eGMCommands.EventoExperiencia
476                Call handleEventoExperiencia(UserIndex)
 
490            Case eGMCommands.CuentaRegresiva
492                Call HandleCuentaRegresiva(UserIndex)

           End Select

        End With

        Exit Sub

ErrHandler:
494        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGMCommands", Erl)
           Resume Next
End Sub

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 16 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName As String, Cuenta As String, Password As String, Version As String, MacAdress As String
    
    Dim MensajeAdvertencia As Integer, HDserial As Long, CierraConexion As Boolean
    
    Cuenta = Buffer.ReadASCIIString()
    Password = Buffer.ReadASCIIString()
    Version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    UserName = Buffer.ReadASCIIString()
    MacAdress = Buffer.ReadASCIIString()
    HDserial = Buffer.ReadLong()
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
    'Chequeos generales
    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(UserIndex, 0, True, 7) 'Juego desactualizado
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        GoTo ErrHandler
        Exit Sub
    End If
 
    If Not ChequeosServerIni(UserIndex, UserName, "", MensajeAdvertencia, CierraConexion) Then
    
        Call WriteShowMessageBox(UserIndex, MensajeAdvertencia)
        
        If CierraConexion = True Then
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
 
    Else
        'Si pasó los chequeos generales hacemos los chequeos por cuenta, por si un vivo quiere logear cosas con datos invalidos :p
    
        If Not EntrarCuenta(UserIndex, Cuenta, Password, MacAdress, HDserial) Then
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
    
        If Len(UserName) <= 0 Then
            Call WriteShowMessageBox(UserIndex, 64) 'Nombre invalido
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
            
        If Not AsciiValidos(UserName) Then
            Call WriteShowMessageBox(UserIndex, 64) 'Nombre invalido
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
            
        If Not PersonajeExiste(UserName) Then
            Call WriteShowMessageBox(UserIndex, 38) 'El personaje no existe
            Call FlushBuffer(UserIndex)
             'Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
         
        '¿El personaje personaje pertenece a su cuenta?
        If Not IsPjOfAccount(UserName, Cuenta) Then
            Call WriteShowMessageBox(UserIndex, 37) 'Personaje invalido.
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
                 
        '¿Ya esta conectado el personaje?
        Dim tIndex As Integer
        tIndex = NameIndex(UserName)
        
        If tIndex > 0 And tIndex <> UserIndex Then
            If UserList(tIndex).Counters.Saliendo Then
                Call WriteShowMessageBox(UserIndex, 44)
            Else
                Call WriteShowMessageBox(UserIndex, 46)
                Call WriteShowMessageBox(tIndex, 45)
                Call Cerrar_Usuario(tIndex)
            End If
            
            GoTo ErrHandler
            Exit Sub
        End If
 
        'Se conectaron dos a la cuenta sin logear?
        If CuentaConectada(Cuenta) = 1 Then
            Call WriteShowMessageBox(UserIndex, 47) 'Ya hay un usuario conectado con esta cuenta, solo puedes logear un personaje por cuenta.
            GoTo ErrHandler
            Exit Sub
        End If
        
        If BANCheck(UserName) Then
             Call WriteShowMessageBox(UserIndex, 22) 'El personaje ha sido bloqueado y no puede ser accedido.
        Else
            
            If Not ConnectUser(UserIndex, UserName, Cuenta, CierraConexion) Then
                If CierraConexion = True Then
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    GoTo ErrHandler
                    Exit Sub
                End If
            End If
            
        End If
        
    End If
 
ErrHandler:
    Dim Error As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 22 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName As String, Password As String, MacAddress As String, Version As String, Account As String
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    Dim Cabeza As Integer, HDserial As Long, MensajeAdvertencia As Integer, CierraConexion As Boolean

    Account = Buffer.ReadASCIIString()
    Password = Buffer.ReadASCIIString()
    
    Version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    UserName = Buffer.ReadASCIIString()
 
    race = Buffer.ReadByte()
    gender = Buffer.ReadByte()
    Class = Buffer.ReadByte()
    homeland = Buffer.ReadByte()
    
    Cabeza = Buffer.ReadInteger()

    MacAddress = Buffer.ReadASCIIString()
    
    HDserial = Buffer.ReadLong()
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)

    'Chequeos generales
    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(UserIndex, 0, True, 7) 'Juego desactualizado
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        GoTo ErrHandler
        Exit Sub
    End If
 
    If Not ChequeosServerIni(UserIndex, UserName, "", MensajeAdvertencia, CierraConexion) Then
    
        Call WriteShowMessageBox(UserIndex, MensajeAdvertencia)
        
        If CierraConexion = True Then
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            GoTo ErrHandler
            Exit Sub
        End If
 
    Else
    
       If Not EntrarCuenta(UserIndex, Account, Password, MacAddress, HDserial) Then
           Call FlushBuffer(UserIndex)
           Call CloseSocket(UserIndex)
           GoTo ErrHandler
           Exit Sub
       End If
       
        If Not ConnectNewUser(UserIndex, UserName, race, gender, Class, Cabeza, Account, homeland, MensajeAdvertencia, CierraConexion) Then
            
            Call WriteShowMessageBox(UserIndex, MensajeAdvertencia)
            
            If CierraConexion = True Then
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                GoTo ErrHandler
                Exit Sub
            End If
        
        End If
        
     End If
     
ErrHandler:

    Dim Error As Long
    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)


    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        
        Dim TalkMode As Byte
        
        chat = Buffer.ReadASCIIString()
        
        TalkMode = Buffer.ReadByte()
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
 
        If Len(chat) >= 80 Then
            Call WriteLocaleMsg(UserIndex, 238)
        
        Else
 
            If Not (.flags.AdminInvisible = 1) Then
                'I see you....
                If .flags.Oculto > 0 Then
                    .flags.Oculto = 0
                    .Counters.TiempoOculto = 0
                    If .flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        Call WriteLocaleMsg(UserIndex, 307)
                    End If
                End If
        
                If .Counters.IdleCount > 0 Then .Counters.IdleCount = 0
                
                Select Case TalkMode
                
                    Case 1 'Normal
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, TalkMode))
                            
                    Case 2 ' Gritar
                    
                        If Not DeadCheck(UserIndex) Then
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessageChatOverHead(mid(chat, 2), .Char.CharIndex, TalkMode))
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessageConsoleMsg("[" & .Name & "] " & mid(chat, 2), 25))
                        End If

                    Case 3 ' Global
                        If Not DeadCheck(UserIndex) Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[" & .Name & "] " & mid(chat, 2), 23))
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageChatOverHead(mid(chat, 2), .Char.CharIndex, TalkMode))
                        End If
                    
                End Select
                
             Else
                
                If chat <> "" Then
                    Call SendData(SendTarget.ToADMINS, UserIndex, PrepareMessageConsoleMsg("Game Master> " & "[" & .Name & "] " & chat, 14))
                End If
                
             End If

         End If
 
    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Whisper" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        On Error GoTo ErrHandler

104     With UserList(UserIndex)

            'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
            Dim Buffer As New clsByteQueue
            Set Buffer = New clsByteQueue
106         Call Buffer.CopyBuffer(.incomingData)
        
            'Remove packet ID
108         Call Buffer.ReadByte
        
            Dim chat            As String
            Dim targetCharIndex As String
            
            Dim targetUserIndex As Integer
            Dim rank            As Integer
        
110         rank = .flags.Privilegios

112         targetCharIndex = Buffer.ReadASCIIString()
114         chat = Buffer.ReadASCIIString()
            
            Call .incomingData.CopyBuffer(Buffer)
            
116         targetUserIndex = NameIndex(targetCharIndex)

118         If targetUserIndex <= 0 Then 'existe el usuario destino?
120             Call WriteLocaleMsg(UserIndex, 75) 'Usuario offline
            Else
            
                If UserIndex <> targetUserIndex Then
                    
126                 If EstaPCarea(UserIndex, targetUserIndex) Then
    
128                     If LenB(chat) <> 0 Then
132                         Call SendData(SendTarget.ToADMINS, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, 4))
    
                            Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, 22)
                            Call WriteConsoleMsg(targetUserIndex, "[" & .Name & "] " & chat, 22)
                                
134                         Call WriteChatOverHead(UserIndex, chat, .Char.CharIndex, 4)
136                         Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, 4)
    
                        End If
                            
                    Else
140                     Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, 22)
142                     Call WriteConsoleMsg(targetUserIndex, "[" & .Name & "] " & chat, 22)
    
                    End If
    
                End If
            
            End If

        End With
    
ErrHandler:

        Dim Error As Long

148     Error = Err.Number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
150     Set Buffer = Nothing
    
152     If Error <> 0 Then Err.Raise Error

End Sub


''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
    
    On Error GoTo HandleWalk_Err
    
    Dim demora      As Long

    Dim demorafinal As Long
    
    demora = GetTickCount And &H7FFFFFFF
    
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy    As Long
    Dim TempTick As Long
    
    Dim heading  As eHeading
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte
        
        If .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then

            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.Loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteLocaleMsg(UserIndex, 123)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), 0, True, True))
            End If
                
            'Move user
            Call MoveUserChar(UserIndex, heading)

            'Stop resting if needed
            If .flags.Descansar Then
                .flags.Descansar = False
                
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            If UserList(UserIndex).flags.Trabajando = True Then
               UserList(UserIndex).flags.Trabajando = False
               Call WriteLocaleMsg(UserIndex, 391, vbNullString, 1)
            End If
            
            Dim TiempoDeWalk As Byte
            
            TiempoDeWalk = 34
            
            'Prevent SpeedHack
            If .flags.TimesWalk >= TiempoDeWalk Then
            
                TempTick = GetTickCount And &H7FFFFFFF
                dummy = (TempTick - .flags.StartWalk)
                
                    ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
                    '(it's about 193 ms per step against the over 200 needed in perfect conditions)
                    If dummy < 5200 Then
                        If TempTick - .flags.CountSH > 30000 Then
                            .flags.CountSH = 0
                        End If
                    
                        If Not .flags.CountSH = 0 Then
                            If dummy <> 0 Then dummy = 126000 \ dummy
        
                            Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH. (DUFTIN/MAYCO SACAR ESTE SISTEM FEO", FontTypeNames.FONTTYPE_SERVER))
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        Else
                            .flags.CountSH = TempTick
                        End If
                    End If
                    
                .flags.StartWalk = TempTick
                .flags.TimesWalk = 0
               
            End If
            
            
            .flags.TimesWalk = .flags.TimesWalk + 1
            
            Call CancelExit(UserIndex)
                    

        Else    'paralized
        
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                Call WriteLocaleMsg(UserIndex, 54)
            End If
            
            .flags.CountSH = 0
                
        End If
        
    End With
    
    demorafinal = (GetTickCount And &H7FFFFFFF) - demora
    
    Exit Sub

HandleWalk_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWalk", Erl)
     Resume Next
        
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)

    On Error GoTo HandleRequestPositionUpdate_Err
    
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)

    Exit Sub

HandleRequestPositionUpdate_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestPositionUpdate", Erl)
     Resume Next
        
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
    
    On Error GoTo HandleAttack_Err
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If DeadCheck(UserIndex) Then Exit Sub
        
        If SeguroCheck(UserIndex, 1) Then Exit Sub
        
        'If user meditates, can't attack
        If .flags.Meditando Then Exit Sub


        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil > 0 Then
                Call WriteLocaleMsg(UserIndex, 127)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.Invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteLocaleMsg(UserIndex, 307)
            End If
        End If

    End With
    
    Exit Sub

HandleAttack_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleAttack", Erl)
     Resume Next
        
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
    
    On Error GoTo HandlePickUp_Err

    With UserList(UserIndex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If DeadCheck(UserIndex) Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteLocaleMsg(UserIndex, 25)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)

    End With
        
    Exit Sub

HandlePickUp_Err:
    Call RegistrarError(Err.Number, Err.description, "Protocol.HandlePickUp", Erl)
    Resume Next
        
End Sub

''
' Handles the "CombatModeToggle" message.
'
' @param    userIndex The index of the user sending the message.
 
Private Sub HanldeCombatModeToggle(ByVal UserIndex As Integer)

On Error GoTo HanldeCombatModeToggle_Err
        
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.ModoCombate Then
            Call WriteLocaleMsg(UserIndex, 13)
        Else
            Call WriteLocaleMsg(UserIndex, 12)
        End If
        
        .flags.ModoCombate = Not .flags.ModoCombate
        
    End With
    
    Exit Sub

HanldeCombatModeToggle_Err:
112     Call RegistrarError(Err.Number, Err.description, "Protocol.HanldeCombatModeToggle", Erl)
114     Resume Next
        
End Sub


''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
    
    On Error GoTo HandleResuscitationToggle_Err

    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If .flags.SeguroResu Then
            Call WriteLocaleMsg(UserIndex, 15)
        Else
            Call WriteLocaleMsg(UserIndex, 14)
        End If
        
        .flags.SeguroResu = Not .flags.SeguroResu

    End With
    
    Exit Sub

HandleResuscitationToggle_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleResuscitationToggle", Erl)
     Resume Next
        
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)

    On Error GoTo HandleRequestGuildLeaderInfo_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)

    Exit Sub

HandleRequestGuildLeaderInfo_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
     Resume Next
        
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)

    On Error GoTo HandleRequestAtributes_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAttributes(UserIndex)

    Exit Sub

HandleRequestAtributes_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestAtributes", Erl)
     Resume Next
        
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)

    On Error GoTo HandleRequestSkills_Err
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)
    
    Exit Sub

HandleRequestSkills_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestSkills", Erl)
     Resume Next
        
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)

    On Error GoTo HandleRequestMiniStats_Err

    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMiniStats(UserIndex)

    Exit Sub

HandleRequestMiniStats_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestMiniStats", Erl)
     Resume Next
        
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)

    On Error GoTo HandleCommerceEnd_Err
    

    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
    
    Exit Sub

HandleCommerceEnd_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceEnd", Erl)
     Resume Next
        
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
    
    On Error GoTo HandleBankEnd_Err
    

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)

    End With
        
    Exit Sub

HandleBankEnd_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankEnd", Erl)
     Resume Next
        
End Sub
 
''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
    
    On Error GoTo HandleDrop_Err

    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim slot   As Byte
    Dim Amount As Long
    
    Dim ObjDestroy As ObjData
    Dim ObjIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
  
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not _
                .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If slot = FLAGORO Then
            If Amount > 100000 Then Exit Sub  'Don't drop too much gold

            Call TirarOro(Amount, UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ORO2, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            Call WriteUpdateGold(UserIndex)
        Else
           
            If slot <= MAX_INVENTORY_SLOTS And slot > 0 Then
                If .Invent.Object(slot).ObjIndex = 0 Then
                    Exit Sub

                End If
             
             ' $ Shermie80
             ObjIndex = .Invent.Object(slot).ObjIndex
             ObjDestroy = ObjData(ObjIndex)


            If .flags.Montando = 1 Then
                If ObjData(ObjIndex).OBJType = otMonturas Then
                    Call DoEquita(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
                End If
            End If
            
            If .flags.Navegando = 1 Then
                If ObjData(ObjIndex).OBJType = otBarcos Then
                    Call WriteLocaleMsg(UserIndex, 20)
                    Exit Sub
                End If
            End If

             If ObjDestroy.Destruir = 1 Then
                Call WriteShowMessageBox(UserIndex, "", True, 1)
                Exit Sub
             End If
             
             
             Call DropObj(UserIndex, slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)

            End If

        End If

    End With

    Exit Sub

HandleDrop_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDrop", Erl)
     Resume Next
        
End Sub

''
' Handles the "DropDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDropDestroy(ByVal UserIndex As Integer)

    On Error GoTo HandleDropDestroy_Err
    
    '***************************************************
    'Author: Shermie80
    'Last Modification: 01/08/20
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim slot   As Byte
    Dim Amount As Long

     With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()

        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or .flags.Muerto = 1 Or .flags.Montando = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not _
           .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub
            
            'Only drop valid slots
            If slot <= MAX_INVENTORY_SLOTS And slot > 0 Then
                If .Invent.Object(slot).ObjIndex = 0 Then
                    Exit Sub

            End If

           Call QuitarUserInvItem(UserIndex, slot, Amount)
           Call UpdateUserInv(False, UserIndex, slot)

         End If

    End With

    Exit Sub

HandleDropDestroy_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDropDestroy", Erl)
     Resume Next
        
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
    
    On Error GoTo HandleCastSpell_Err
    

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
100    With UserList(UserIndex)
    
          'Remove packet ID
102        Call .incomingData.ReadByte
        
104        Dim Spell As Byte
        
106        Spell = .incomingData.ReadByte()
        
108        If DeadCheck(UserIndex) Then Exit Sub
        
110        'Now you can be atacked
112        .flags.NoPuedeSerAtacado = False
        
114        If Spell < 1 Then
116            .flags.Hechizo = 0
118            Exit Sub
120        ElseIf Spell > MAXUSERHECHIZOS Then
122            .flags.Hechizo = 0
124            Exit Sub
126        End If
           
           .flags.Hechizo = Spell

130        If Spell <> 0 Then
        
132            Dim uh As Integer
134            uh = .Stats.UserHechizos(Spell)

135            If uh <> 0 Then
136                If Hechizos(uh).AutoLanzar > 0 Then
138                    .flags.TargetUser = UserIndex
140                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
142                Else
144                    Call WriteWorkRequestTarget(UserIndex, eSkill.magia)
                
146                End If

147             End If

148         End If
        
150    End With

152    Exit Sub

HandleCastSpell_Err:
156     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCastSpell", Erl)
158     Resume Next
        
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
    
    On Error GoTo HandleLeftClick_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

    End With

    Exit Sub

HandleLeftClick_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleLeftClick", Erl)
     Resume Next
        
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
    
    On Error GoTo HandleDoubleClick_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

    End With

    Exit Sub

HandleDoubleClick_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDoubleClick", Erl)
     Resume Next
        
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
    
    On Error GoTo HandleWork_Err
    
    
        If UserList(UserIndex).incomingData.length < 2 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
        
        With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If DeadCheck(UserIndex) Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
        
            Case pesca, robar, talar, mineria, FundirMetal, domar
            
            
                Call WriteWorkRequestTarget(UserIndex, Skill)
 
 
            Case Ocultarse
            
                If .flags.Oculto = 1 Then Exit Sub
                 
                If .flags.Navegando = 1 Then
                    Call WriteLocaleMsg(UserIndex, 56) 'No podes ocultarte si estas navegando.
                    Exit Sub
                End If
                
                If .flags.Montando = 1 Then
                    Call WriteLocaleMsg(UserIndex, 67) 'No podes ocultarte si estas montando.
                    Exit Sub
                End If

                If MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto > 0 Then
                    Call WriteLocaleMsg(UserIndex, 447) '¡No puedes ocultarte aquí!
                    Exit Sub
                End If

                Call DoOcultarse(UserIndex)

        End Select

    End With

    Exit Sub

HandleWork_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWork", Erl)
     Resume Next
        
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '
    '***************************************************
    Dim TotalItems    As Long
    Dim ItemsPorCiclo As Integer
      If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger
        
        If TotalItems > 0 Then
            
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)
            
        End If

    End With

End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
    
    On Error GoTo HandleUseItem_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        
        slot = .incomingData.ReadByte()

   
        If slot <= MAX_INVENTORY_SLOTS And slot > 0 Then
            If .Invent.Object(slot).ObjIndex = 0 Then Exit Sub
        End If

        Call UseInvItem(UserIndex, slot)
        
    End With

    Exit Sub

HandleUseItem_Err:
    Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUseItem", Erl)
    Resume Next
        
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)

    On Error GoTo HandleCraftBlacksmith_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim cant As Integer
        
        Item = .ReadInteger()
        cant = .ReadInteger()
        
        If Item < 1 Or cant < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub

        Call HerreroConstruirItem(UserIndex, Item, cant)

    End With

    Exit Sub

HandleCraftBlacksmith_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftBlacksmith", Erl)
     Resume Next
        
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)

    On Error GoTo HandleCraftCarpenter_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim cant As Integer
        
        Item = .ReadInteger()
        cant = .ReadInteger()
        
        If Item < 1 Or cant < 1 Or cant > 1000 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(UserIndex, Item, cant)
    End With
    
    Exit Sub

HandleCraftCarpenter_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftCarpenter", Erl)
     Resume Next
        
End Sub

Private Sub HandleCraftalquimia(ByVal UserIndex As Integer)

    On Error GoTo HandleCraftAlquimia_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim cant As Integer
        
        Item = .ReadInteger()
        cant = .ReadInteger()
        
        If Item < 1 Or cant < 1 Or cant > 1000 Then Exit Sub
        
        If ObjData(Item).SkPociones = 0 Then Exit Sub
        
        Call druidaConstruirItem(UserIndex, Item, cant)
    End With
    
    Exit Sub

HandleCraftAlquimia_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftAlquimia", Erl)
     Resume Next
        
End Sub

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)

    On Error GoTo HandleCraftSastre_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim cant As Integer
        
        Item = .ReadInteger()
        cant = .ReadInteger()
        
        If Item < 1 Or cant < 1 Or cant > 1000 Then Exit Sub
        
        If ObjData(Item).SkSastreria = 0 Then Exit Sub
        Call SastreConstruirItem(UserIndex, Item, cant)
    End With
    
    Exit Sub

HandleCraftSastre_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftSastre", Erl)
     Resume Next
        
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
    
    On Error GoTo HandleWorkLeftClick_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 14/01/2010 (ZaMa)
    '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
    '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
    '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X        As Byte
        Dim Y        As Byte
        Dim Skill    As eSkill
        Dim DummyInt As Integer
        Dim tU       As Integer   'Target user
        Dim tN       As Integer   'Target NPC
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
          Case eSkill.ArmasArrojadizas
            
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    If .WeaponEqpObjIndex = 0 Then
                         DummyInt = 1
                '    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                '        DummyInt = 1
                   ElseIf ObjData(.WeaponEqpObjIndex).SubTipo <> 5 Then
                        DummyInt = 2
                    ElseIf .Object(.WeaponEqpSlot).Amount < 1 Then
                        DummyInt = 1
                   End If

                 If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteLocaleMsg(UserIndex, 125)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If

                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    Call WriteLocaleMsg(UserIndex, 93)
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
 
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteLocaleMsg(UserIndex, 8)
                        Exit Sub
                    End If

                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteLocaleMsg(UserIndex, 298)
                        
                        Exit Sub
                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    Call UsuarioAtacaUsuario(UserIndex, tU)
                    
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteLocaleMsg(UserIndex, 8)
                        Exit Sub
                    End If

                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                 
                With .Invent
                    DummyInt = .WeaponEqpSlot
                    
                    'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                    
                    'If .Object(DummyInt).amount > 0 Then
                    '    'QuitarUserInvItem unequipps the ammo, so we equip it again
                    '   .WeaponEqpSlot = DummyInt
                    '    .WeaponEqpObjIndex = .Object(DummyInt).ObjIndex
                    '    .Object(DummyInt).Equipped = 1
                    'Else
                    '    .WeaponEqpSlot = 0
                    '    .WeaponEqpObjIndex = 0
                    'End If
                    Call UpdateUserInv(False, UserIndex, DummyInt)
                End With
                '-----------------------------------
            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub

                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub

                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                Dim Atacked As Boolean
                Atacked = True
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent

                    ' Tiene arma equipada?
                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                        ' En un slot válido?
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                        ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
                    ElseIf ObjData(.WeaponEqpObjIndex).Municion = 1 Then

                        ' La municion esta equipada en un slot valido?
                        If .MunicionEqpSlot < 1 Or .MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                            DummyInt = 1
                            ' Tiene munición?
                        ElseIf .MunicionEqpObjIndex = 0 Then
                            DummyInt = 1
                            ' Son flechas?
                        ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                            DummyInt = 1
                            ' Tiene suficientes?
                        ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                            DummyInt = 1

                        End If

                        ' Es un arma de proyectiles?
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2

                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteLocaleMsg(UserIndex, 127)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)

                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub

                    End If

                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else

                    Call WriteLocaleMsg(UserIndex, 93)

                    Exit Sub

                End If
                
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then

                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteLocaleMsg(UserIndex, 8)
                        Exit Sub

                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteLocaleMsg(UserIndex, 298)
                        Exit Sub

                    End If
                    
                    'Attack!
                    Atacked = UsuarioAtacaUsuario(UserIndex, tU)
                    
                ElseIf tN > 0 Then

                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > _
                            RANGO_VISION_X Then
                        Call WriteLocaleMsg(UserIndex, 8)
                        Exit Sub

                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Atacked = UsuarioAtacaNpc(UserIndex, tN)

                    End If

                End If
                
                ' Solo pierde la munición si pudo atacar al target, o tiro al aire
                If Atacked Then

                    With .Invent

                        ' Tiene equipado arco y flecha?
                        If ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                            DummyInt = .MunicionEqpSlot
                            
                            'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                            Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            
                            If .Object(DummyInt).Amount > 0 Then
                                'QuitarUserInvItem unequips the ammo, so we equip it again
                                .MunicionEqpSlot = DummyInt
                                .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .MunicionEqpSlot = 0
                                .MunicionEqpObjIndex = 0

                            End If

                            ' Tiene equipado un arma arrojadiza
                        Else
                            DummyInt = .WeaponEqpSlot
                            
                            'Take 1 knife away
                            Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            
                            If .Object(DummyInt).Amount > 0 Then
                                'QuitarUserInvItem unequips the weapon, so we equip it again
                                .WeaponEqpSlot = DummyInt
                                .WeaponEqpObjIndex = .Object(DummyInt).ObjIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .WeaponEqpSlot = 0
                                .WeaponEqpObjIndex = 0

                            End If
                            
                        End If
                        
                        Call UpdateUserInv(False, UserIndex, DummyInt)

                    End With

                End If
            
            Case eSkill.magia
            
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteLocaleMsg(UserIndex, 385)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posición (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then

                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If
                End If
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                End If
            
            Case eSkill.robar

                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                    Call WriteLocaleMsg(UserIndex, 8)
                                    Exit Sub

                                End If
                                 
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, X, Y).Trigger = eTrigger.ZONASEGURA Then
                                    Call WriteLocaleMsg(UserIndex, 310)
                                    Exit Sub

                                End If
                                 
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                                    Call WriteLocaleMsg(UserIndex, 310)
                                    Exit Sub

                                End If
                                 
                                Call DoRobar(UserIndex, tU)

                            End If

                        End If

                    End If

                Else
                
                    Call WriteLocaleMsg(UserIndex, 246, vbNullString, 1)

                End If

            Case eSkill.domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteLocaleMsg(UserIndex, 8)
                            Exit Sub

                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteLocaleMsg(UserIndex, 300)
                            Exit Sub

                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteLocaleMsg(UserIndex, 300)

                    End If

                Else
                    Call WriteLocaleMsg(UserIndex, 300)

                End If
                
            'Public Const FundirMetal                As Integer = 88
            Case FundirMetal
                Call FundirMineral(UserIndex, X, Y)

            

 
                
        End Select

    End With

    Exit Sub

HandleWorkLeftClick_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWorkLeftClick", Erl)
     Resume Next
        
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/11/09
    '05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
    '***************************************************
    If UserList(UserIndex).incomingData.length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim desc      As String
        Dim GuildName As String
        Dim site      As String
        Dim codex()   As String
        Dim errorStr  As String
        
        desc = Buffer.ReadASCIIString()
        GuildName = Trim$(Buffer.ReadASCIIString())
        site = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & " fundó el clan " & GuildName _
                    & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            Call QuitarObjetos(406, 1, UserIndex)
            Call QuitarObjetos(408, 1, UserIndex)
            Call QuitarObjetos(409, 1, UserIndex)
            Call QuitarObjetos(410, 1, UserIndex)
            
            'Update tag
            Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
    
    On Error GoTo HandleEquipItem_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
      
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If DeadCheck(UserIndex) Then Exit Sub
        
        'Validate item slot
        If itemSlot > MAX_INVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
 
        Call EquiparInvItem(UserIndex, itemSlot)
 
    End With
    
    Exit Sub

HandleEquipItem_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleEquipItem", Erl)
     Resume Next
        
End Sub

Private Sub HandleEquiparSkin(ByVal UserIndex As Integer)

    On Error GoTo HandleEquiparItem_Err

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Equipo As Byte, BackOrNext As Byte
        
        Equipo = .incomingData.ReadByte()
        BackOrNext = .incomingData.ReadByte()
        
        
       Call General.EquipaSkin(UserIndex, Equipo, BackOrNext)
       Call ChangeUserCharTodo(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    
    End With

    Exit Sub

HandleEquiparItem_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleEquiparItem", Erl)
     Resume Next
        
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

    On Error GoTo HandleChangeHeading_Err
    

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
     
        Dim heading As eHeading
 
        heading = .incomingData.ReadByte()
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
                
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.heading, 3))
        
        End If
        
    End With
    Exit Sub

HandleChangeHeading_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeHeading", Erl)
     Resume Next
        
End Sub


''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)

    On Error GoTo HandleModifySkills_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 1 + NUMSKILLS Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            count = count + points(i)
        Next i
        
        If count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100
                End If
            Next i
        End With
    End With
    Exit Sub

HandleModifySkills_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleModifySkills", Erl)
     Resume Next
        
End Sub


''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
    
    On Error GoTo HandleTrain_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim PetIndex   As Byte
        
        PetIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).npcindex, Npclist( _
                        .flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1

                End If

            End If

        Else
        
            Call WriteLocaleMsg(UserIndex, 593, vbNullString, 1)

        End If

    End With

        
    Exit Sub

HandleTrain_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleTrain", Erl)
     Resume Next
        
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
    
    On Error GoTo HandleCommerceBuy_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot   As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 594, 4) 'Npc no tengo interes en comerciar
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteLocaleMsg(UserIndex, 388)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, slot, Amount)

        
    End With
    
    Exit Sub

HandleCommerceBuy_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceBuy", Erl)
     Resume Next
        
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot   As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
 
          Call UserRetiraItem(UserIndex, slot, Amount)

    End With

End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot   As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 594, 4)  'Npc no tengo interes en comerciar
            Exit Sub

        End If
        
        'User compra el item del slot
        'Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)
         

          Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, slot, Amount)

    End With

End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot   As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
        
        'User deposita el item del slot rdata
        'Call UserDepositaItem(UserIndex, Slot, Amount)

           Call UserDepositaItem(UserIndex, slot, Amount)

    End With

End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1

        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())

    End With

End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir      As Integer
        Dim slot     As Byte
        Dim TempItem As Obj
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1

        End If
        
        slot = .ReadByte()

    End With
        
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .BancoInvent.Object(slot) = .BancoInvent.Object(slot - 1)
            .BancoInvent.Object(slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .BancoInvent.Object(slot) = .BancoInvent.Object(slot + 1)
            .BancoInvent.Object(slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(slot + 1).Amount = TempItem.Amount

        End If

    End With
    
    Call UpdateBanUserInv(True, UserIndex, 0)
    Call UpdateVentanaBanco(UserIndex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim desc    As String
        Dim codex() As String
        
        desc = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg( _
                    "Tu clan ha firmado la paz con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg( _
                    "Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", _
                    FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg( _
                    "Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName( _
                    .GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", _
                    FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg( _
                    "Tu clan rechazado la propuesta de paz de " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName( _
                    .GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", _
                    FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        
        guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg( _
                    "Tu clan ha firmado la alianza con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg( _
                    "Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", _
                    FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild    As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje224)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild    As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje224)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild    As String
        Dim errorStr As String
        Dim details  As String
        
        guild = Buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild    As String
        Dim errorStr As String
        Dim details  As String
        
        guild = Buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim User    As String
        Dim details As String
        
        User = Buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje225)
        Else
            Call WriteShowUserRequest(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))

End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))

End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild           As String
        Dim errorStr        As String
        Dim otherGuildIndex As Integer
        
        guild = Buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg( _
                    "TU CLAN HA ENTRADO EN GUERRA CON " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName( _
                    .GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, _
                    NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)

            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)

            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & _
                    " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim reason   As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, reason)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName   As String
        Dim GuildIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & _
                    " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje226)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg( _
                    "¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " _
                    & .Name, FontTypeNames.FONTTYPE_GUILD))

        End If

    End With

End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild       As String
        Dim application As String
        Dim errorStr    As String
        
        guild = Buffer.ReadASCIIString()
        application = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & _
                    guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub
Private Sub HandleOnline(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
 
        Dim i     As Long
        Dim count As Long
        Dim str   As String
   
        With UserList(UserIndex)
                'Remove packet ID
                Call .incomingData.ReadByte
                       
                For i = 1 To LastUser
 
                        If LenB(UserList(i).Name) <> 0 Then

                           If LastUser = i Then
                             str = str + UserList(i).Name
                           Else
                             str = str + UserList(i).Name & ", "
                           End If
                                   
                           count = count + 1
                                   
                        End If
 
                Next i
                
                Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(count) & " - Record de usuarios: " & CStr(recordusuarios), FontTypeNames.FONTTYPE_INFO)
                
                If count > 0 Then
                    Call WriteConsoleMsg(UserIndex, str, FontTypeNames.FONTTYPE_INFO)
                End If
                
        End With
 
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    'Dim tuser As Integer
    'Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteLocaleMsg(UserIndex, 119)
            Exit Sub
        End If

19       Call Cerrar_Usuario(UserIndex)
         
    End With
    
End Sub

 

 

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)
        
        If GuildIndex > 0 Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje229)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", _
                    FontTypeNames.FONTTYPE_GUILD))
        Else
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje230)

        End If

    End With

End Sub
 
''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/08 (NicoNZ)
    'Arreglé un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            
            Exit Sub

        End If
 
        If .flags.Montando = 1 Then
 
        .flags.Montando = 0
      
          .Char.Head = UserList(UserIndex).OrigChar.Head
           If .Invent.ArmourEqpObjIndex > 0 Then
              .Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
           Else
               Call DarCuerpoDesnudo(UserIndex)
           End If
           
         If .Invent.NudiEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.NudiEqpObjIndex).WeaponAnim
         If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
         If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
         If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
 
 Call ChangeUserCharTodo(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
 Call WriteMontateToggle(UserIndex)
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then Call WriteLocaleMsg(UserIndex, 123)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
      
        Else
            '.Counters.bPuedeMeditar = False
            
            '.Char.FX = 0
            '.Char.loops = 0
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), 0, True, True))
            
        End If

    End With

End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteLocaleMsg(UserIndex, 22)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If
        
        Call RevivirUsuario(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(Npclist(.flags.TargetNPC).Char.CharIndex, 11, 1))
    End With

End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(UserIndex, UserIndex)

End Sub
 
''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteLocaleMsg(UserIndex, 387)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then

            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
            
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 594, 4) 'Npc no tengo interes en comerciar
                End If
                
                Exit Sub

            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteLocaleMsg(UserIndex, 8)
                Exit Sub

            End If
            
                   'Es de una faccion en otro mapa? mermas /COMERCIAR (el otro es el click derecho)
 
                    'If .flags.Privilegios And PlayerType.User Then
                    
                        'Es de una faccion en otro mapa? mermas
                     '   If (esCiuda(UserIndex) Or esArmada(UserIndex)) And (MapInfo(.Pos.Map).battle_mode <> 2 And MapInfo(.Pos.Map).battle_mode <> 3) Then
                      '      Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 592, 4) 'No comercia con enemigos en la ciudad
                       '     Exit Sub
                       ' End If
                        
                    '    If (esRepu(UserIndex) Or esMili(UserIndex)) And (MapInfo(.Pos.Map).battle_mode <> 2 And MapInfo(.Pos.Map).battle_mode <> 3) Then
                     '       Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 592, 4) 'No comercia con enemigos en la ciudad
                      '      Exit Sub
                       ' End If
                        
                       ' If (esRene(UserIndex) Or esCaos(UserIndex)) And (MapInfo(.Pos.Map).battle_mode <> 3) Then
                        '    Call WriteChatOverHeadLocale(UserIndex, Npclist(.flags.TargetNPC).Char.CharIndex, 592, 4) 'No comercia con enemigos en la ciudad
                         '   Exit Sub
                       ' End If
                    
                   ' End If
                    
                    
                    
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
            '[Alejo]
        End If

    End With

End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        If .flags.Comerciando Then
            Call WriteLocaleMsg(UserIndex, 387)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteLocaleMsg(UserIndex, 8)
                Exit Sub

            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex, False)

            End If

        Else
            Call WriteLocaleMsg(UserIndex, 22)

        End If

    End With

End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteLocaleMsg(UserIndex, 22)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.facciones Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If
        
        Select Case Npclist(.flags.TargetNPC).flags.Status
        
        Case 1
        Call EnlistarArmadaReal(UserIndex)
        
        Case 2
        Call EnlistarMilicia(UserIndex)
        
        Case 4
        Call EnlistarCaos(UserIndex)
        
        End Select
        

    End With

End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteLocaleMsg(UserIndex, 22)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.facciones Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If

        Select Case Npclist(.flags.TargetNPC).flags.Status
        
        Case 1 'Armada
        
            If Not esArmada(UserIndex) Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex)
            Exit Sub
            End If
            
            If UserList(UserIndex).Faccion.Rango = 10 Then
            Call WriteChatOverHead(UserIndex, "Ya no tengo trabajo para darte, has alcanzado el rango más alto aquí.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
            Else
            Call WriteChatOverHead(UserIndex, "¡¡Tu deber es combatir criminales, cada 5 criminales te daré una recompensa!!", Npclist(.flags.TargetNPC).Char.CharIndex)
            Exit Sub
            End If
            
            
        Case 2 'mili
        
            If Not esMili(UserIndex) Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la Milicia Real!!", Npclist(.flags.TargetNPC).Char.CharIndex)
            Exit Sub
            End If
            
            If UserList(UserIndex).Faccion.Rango = 7 Then
            Call WriteChatOverHead(UserIndex, "Ya no tengo trabajo para darte, has alcanzado el rango más alto aquí.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
            Else
            Call WriteChatOverHead(UserIndex, "¡¡Tu deber es combatir criminales, cada 5 criminales te daré una recompensa!!", Npclist(.flags.TargetNPC).Char.CharIndex)
            Exit Sub
            End If
            
        Case 4 'Caos

            If Not esCaos(UserIndex) Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las Fuerzas del Caos, largo de aquí!!", Npclist(.flags.TargetNPC).Char.CharIndex)
            Exit Sub
            End If
            
            If UserList(UserIndex).Faccion.Rango = 10 Then
            Call WriteChatOverHead(UserIndex, "Ya no tengo trabajo para darte, has alcanzado el rango más alto aquí.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
            Else
            Call WriteChatOverHead(UserIndex, "¡¡Tu deber es eliminar a todo tipo de personas, cada 10 personas muertas te daré una recompensa!!", Npclist(.flags.TargetNPC).Char.CharIndex)
            Exit Sub
            End If
            
          End Select

    End With

End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteLocaleMsg(UserIndex, 22)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.facciones Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If
        
        
        Select Case Npclist(.flags.TargetNPC).flags.Status
        
        Case 1 'Armada
        
            If Not esArmada(UserIndex) Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las tropas reales!!", Npclist( _
                        .flags.TargetNPC).Char.CharIndex)
                Exit Sub

            End If

            Call RecompensaArmadaReal(UserIndex)
            
            
        Case 2 'mili
        
            If Not esMili(UserIndex) Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la Milicia republicana!!", Npclist( _
                        .flags.TargetNPC).Char.CharIndex)
                Exit Sub

            End If

            Call RecompensaMilicia(UserIndex)
            
        Case 4 'Caos
        
            If Not esCaos(UserIndex) Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la legión oscura!!", Npclist( _
                        .flags.TargetNPC).Char.CharIndex)
                Exit Sub

            End If

            Call RecompensaCaos(UserIndex)
        
        End Select
        

    End With

End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/08
    '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim Time      As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    Time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    
    If Time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr

    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)

End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 15/07/2009
    '02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
    '15/07/2009: ZaMa - Now invisible admins only speak by console
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim chat As String
        
        chat = Buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
 
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat))
                
                If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToClanArea, UserIndex, _
                        PrepareMessageChatOverHead("< " & chat & " >", .Char.CharIndex))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger())

    End With

End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Compañeros de tu clan conectados: " & onlineList, _
                    FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje099)

        End If

    End With

End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)


    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue
        Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Consulta As String
        Dim Tipo As Byte
        
        Tipo = Buffer.ReadByte()
        Consulta = Buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

        If Not Ayuda.Existe(.Name) Then
 
          Call GuardarConsultaIni(.Name, Consulta, Tipo)
          Call WriteLocaleMsg(UserIndex, 85) 'Mandaste consulta
 
        
          Call Ayuda.Push(.Name, Tipo)
          Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(500, .Name)) 'Mando una consulta
          
        Else
          Call WriteLocaleMsg(UserIndex, 427) 'Volviste a mandar, espera...
          
        End If
        
        
    End With
    
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim description As String
        
        description = Buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
        Else

            If Not AsciiValidos(description) Then
                Call WriteLocaleMsg(UserIndex, 392, vbNullString, 1)
            Else
                .desc = Trim$(description)
                Call WriteLocaleMsg(UserIndex, 111)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim vote     As String
        Dim errorStr As String
        
        vote = Buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMA
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call modGuilds.SendGuildNews(UserIndex)

    End With

End Sub
 
''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Integer
        
        Amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteLocaleMsg(UserIndex, 22)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist( _
                    .flags.TargetNPC).Char.CharIndex)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist( _
                    .flags.TargetNPC).Char.CharIndex)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 5000 monedas.", Npclist( _
                    .flags.TargetNPC).Char.CharIndex)
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex)
        Else

            If RandomNumber(1, 100) <= 47 Then
                        
            
            
                If OroLleno(UserIndex, UserList(UserIndex).Stats.GLD, CLng(Amount)) Then
                Call WriteConsoleMsg(UserIndex, "Tienes la cantidad máxima de oro que puedes tener. No has obtenido oro", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If

                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", _
                        Npclist(.flags.TargetNPC).Char.CharIndex)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", _
                        Npclist(.flags.TargetNPC).Char.CharIndex)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)

        End If

    End With

End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteLocaleMsg(UserIndex, 22)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then

            If OroLleno(UserIndex, UserList(UserIndex).Stats.GLD, Amount) Then
            Call WriteConsoleMsg(UserIndex, "Tienes la cantidad máxima de oro que puedes tener. No has obtenido oro", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            End If
        
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.GLD = .Stats.GLD + Amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist( _
                    .flags.TargetNPC).Char.CharIndex)
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ORO2, .Pos.X, _
                                    .Pos.Y))
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex)

        End If
        
        Call WriteUpdateGold(UserIndex)

    End With

End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
        Call WriteLocaleMsg(UserIndex, 22)
            Exit Sub

        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
            
 
 
        If OroLleno(UserIndex, UserList(UserIndex).Stats.Banco, Amount) Then
                Call WriteConsoleMsg(UserIndex, "Tienes tu billetera llena, por lo tanto no has obtenido oro.", FontTypeNames.FONTTYPE_INFO)
                
        ElseIf Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist( _
                    .flags.TargetNPC).Char.CharIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ORO2, .Pos.X, _
                                    .Pos.Y))
            Call WriteUpdateGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex)
        End If

    End With

End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Text As String
        
        Text = Buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
    
            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " DENUNCIA: " & Text, _
                    FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteLocaleMsg(UserIndex, 85)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 1 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim Error As String
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If HasFound(.Name) Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje271)
            Exit Sub

        End If

        If esCiuda(UserIndex) Or esArmada(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_IMPERIAL
        ElseIf esRepu(UserIndex) Or esMili(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_REPUBLICANO
        ElseIf esCaos(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_CAOTICO
        ElseIf esRene(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_RENEGADO
        Else
            Call WriteConsoleMsg(UserIndex, "Hay un error en su faccion, comuniquese con algun GameMaster", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteAbrirFormularios(UserIndex, 7)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub
    
''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType
        Dim Error    As String
        
        clanType = .incomingData.ReadByte()
        
        If HasFound(.Name) Then
            'Call WriteMensajes(UserIndex, eMensajes.Mensaje271)
            Call LogCheating("El usuario " & .Name & _
                    " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .ip)
            Exit Sub

        End If
        
        Select Case UCase$(Trim(clanType))

            Case eClanType.ct_RoyalArmy
           '     .FundandoGuildAlineacion = ALINEACION_ARMADA
            
            Case eClanType.ct_Milicia
             '   .FundandoGuildAlineacion = ALINEACION_MILICIA
            
            Case eClanType.ct_Evil
               ' .FundandoGuildAlineacion = ALINEACION_LEGION

            Case eClanType.ct_Neutral
               ' .FundandoGuildAlineacion = ALINEACION_NEUTRO

            Case eClanType.ct_GM
              '  .FundandoGuildAlineacion = ALINEACION_MASTER

            Case eClanType.ct_Legal
              '  .FundandoGuildAlineacion = ALINEACION_CIUDA

            Case eClanType.ct_Criminal
              '  .FundandoGuildAlineacion = ALINEACION_CRIMINAL

            Case Else
              '  Call WriteConsoleMsg( userindex, "Alineación inválida.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

        End Select
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteAbrirFormularios(UserIndex, 7)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild       As String
        Dim memberCount As Integer
        Dim i           As Long
        Dim UserName    As String
        
        guild = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")

            End If

            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")

            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        
        message = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
 
                Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg("[Chat GM][" & .Name & "] " & message, _
                        FontTypeNames.FONTTYPE_fonttt))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)

        End If

    End With

End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/07
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        
        UserName = Buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X      As Long
        Dim Y      As Long
        Dim i      As Long
        Dim found  As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or _
                PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or _
                    PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteLocaleMsg(UserIndex, 75)
                Else

                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList( _
                                                tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList( _
                                                tIndex).Pos.Y)
                                        found = True
                                        Exit For

                                    End If

                                End If

                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteLocaleMsg(UserIndex, 403)
                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim comment As String
        comment = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteLocaleMsg(UserIndex, 404)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")

    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))

End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteLocaleMsg(UserIndex, 75)
            Else

                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or _
                        PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or _
                        PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> _
                        0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.Map & ", " _
                            & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/Donde " & UserName)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 30/07/06
    'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1()    As String
        Dim List2()    As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then

            For i = 1 To LastNPC

                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then

                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else

                            For j = 0 To NPCcount1 - 1

                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1

                            End If

                        End If

                    Else

                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else

                            For j = 0 To NPCcount2 - 1

                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1

                            End If

                        End If

                    End If

                End If

            Next i
            
            Call WriteLocaleMsg(UserIndex, 406)

            If NPCcount1 = 0 Then
                Call WriteLocaleMsg(UserIndex, 405)
            Else

                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call WriteLocaleMsg(UserIndex, 407)

            If NPCcount2 = 0 Then
                Call WriteLocaleMsg(UserIndex, 408)
            Else

                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call LogGM(.Name, "Numero enemigos en mapa " & Map)

        End If

    End With

End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        Dim TargetUserName As String, UserName As String
        Dim Map As Integer, X As Integer, Y As Integer, tUser As Integer
        Dim HayTarget As Boolean
        
        TargetUserName = .flags.TargetUser
        
        UserName = Buffer.ReadASCIIString()
        Map = Buffer.ReadInteger()
        
        X = Buffer.ReadByte()
        Y = Buffer.ReadByte()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
        
        If Not PlayerType.User And .flags.Privilegios Then
        
            If TargetUserName > 0 Then
                UserName = UserList(TargetUserName).Name
            End If
            
            If MapaValido(Map) And LenB(UserName) <> 0 Then
            
                If UCase$(UserName) <> "YO" Then
                
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If

                Else
                
                    tUser = UserIndex

                End If
                
                If tUser <= 0 Then
                
                    Call WriteLocaleMsg(UserIndex, 77)
                Else
                
                    If InMapBounds(Map, X, Y) Then
                        Call FindLegalPos(tUser, Map, X, Y)
                        Call WarpUserChar(tUser, Map, X, Y, True, True)
                        Call WriteLocaleMsg(UserIndex, 487, .Name)
                        Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                    End If
                    
                End If
                
            End If
        End If

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)

    End With

End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        tIndex = NameIndex(UserName)
        If Not .flags.Privilegios And PlayerType.User Then Call Ayuda.Quitar(UserName)
         'UserList(tIndex).flags.EnvioGM = 0 'Reseteamos flag
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim X        As Integer
        Dim Y        As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or _
                PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or _
                    PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteLocaleMsg(UserIndex, 75)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", _
                                FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)

                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList( _
                            tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
             
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteAbrirFormularios(UserIndex, 8)

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    'I haven`t found a solution to split, so i make an array of names
    '***************************************************
    Dim i       As Long
    Dim names() As String
    Dim count   As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        count = 1
        
        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) Then
               ' If UserList(i).flags.Privilegios And PlayerType.User And PlayerType.User Then
                    If UserList(i).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

                    names(count) = UserList(i).Name
                    Else
                    names(count) = UserList(i).Name
                    End If
                    
                    count = count + 1

                'End If

            End If

        Next i
        
        If count > 1 Then Call WriteUserNameList(UserIndex, names(), count - 1)

    End With

End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                users = users & ", " & UserList(i).Name
                
                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then users = users & " (*)"

            End If

        Next i
        
        If LenB(users) <> 0 Then
            users = Right$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteLocaleMsg(UserIndex, 409)

        End If

    End With

End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                users = users & UserList(i).Name & ", "

            End If

        Next i
        
        If LenB(users) <> 0 Then
            users = Left$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteLocaleMsg(UserIndex, 410)

        End If

    End With

End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason   As String
        Dim jailTime As Byte
        Dim count    As Byte
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        jailTime = Buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")

        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) _
                <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteLocaleMsg(UserIndex, 411)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteLocaleMsg(UserIndex, 75)
                Else

                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteLocaleMsg(UserIndex, 298)
                    ElseIf jailTime > 60 Then
                        Call WriteLocaleMsg(UserIndex, 412)
                    Else

                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")

                        End If

                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")

                        End If
                        
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                            count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & count + 1, LCase$(.Name) & _
                                    ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(reason) & " " & Date & " " & Time)

                        End If
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarceló a " & UserName)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/22/08 (NicoNZ)
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC   As Integer
        Dim auxNPC As npc

        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, _
                    FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteLocaleMsg(UserIndex, 22)

        End If

    End With

End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason   As String
        Dim privs    As PlayerType
        Dim count    As Byte
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) _
                <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteLocaleMsg(UserIndex, 413)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteLocaleMsg(UserIndex, 298)
                Else

                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")

                    End If

                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")

                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & count + 1, LCase$(.Name) & _
                                ": ADVERTENCIA por: " & LCase$(reason) & " " & Date & " " & Time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", _
                                FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/06/2009
    '02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
    '11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
    '***************************************************
    If UserList(UserIndex).incomingData.length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName      As String
        Dim tUser         As Integer
        Dim Opcion        As Byte
        Dim arg1          As String
        Dim arg2          As String
        Dim valido        As Boolean
        Dim loopc         As Byte
        Dim CommandString As String
        Dim n             As Byte
        Dim UserCharPath  As String
        Dim Var           As Long
        
        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)

        End If
        
        Opcion = Buffer.ReadByte()
        arg1 = Buffer.ReadASCIIString()
        arg2 = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then

            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)

                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, body y level
                    valido = tUser = UserIndex And (Opcion = eEditOptions.eo_Body Or Opcion = eEditOptions.eo_Head Or _
                            Opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    valido = (Opcion = eEditOptions.eo_Level And tUser = UserIndex) Or Opcion = eEditOptions.eo_Body _
                            Or Opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    valido = (Opcion = eEditOptions.eo_Level And tUser = UserIndex) Or Opcion = eEditOptions.eo_Body _
                            Or Opcion = eEditOptions.eo_Head Or Opcion = eEditOptions.eo_CiticensKilled Or Opcion = _
                            eEditOptions.eo_CriminalsKilled Or Opcion = eEditOptions.eo_Class Or Opcion = _
                            eEditOptions.eo_Skills Or Opcion = eEditOptions.eo_addGold

            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True

        End If

        If valido Then
            UserCharPath = CharPath & UserName & ".chr"

            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteLocaleMsg(UserIndex, 80)
                Call LogGM(.Name, "Intentó editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case Opcion

                    Case eEditOptions.eo_Gold

                        If val(arg1) <= MAX_ORO_EDIT Then
                            If tUser <= 0 Then ' Esta offline?
                                Call WriteVar(UserCharPath, "STATS", "GLD", val(arg1))
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                        FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.GLD = val(arg1)
                                Call WriteUpdateGold(tUser)

                            End If

                        Else
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & _
                                    MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", _
                                    FontTypeNames.FONTTYPE_INFO)

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience

                        If val(arg1) > 20000000 Then
                            arg1 = 20000000

                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Var = GetVar(UserCharPath, "STATS", "EXP")
                            Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                        
                            If UserList(tUser).Stats.ELV >= STAT_MAXELV Then
                              
                                Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV _
                                        & ".", FONTTYPE_INFO)
                                'If we got here then packet is complete, copy data back to original queue
                                Call .incomingData.CopyBuffer(Buffer)
                                Exit Sub

                            End If
                        
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body

                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Body", arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserCharTodo(tUser, val(arg1), UserList(tUser).Char.Head, UserList( _
                                    tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList( _
                                    tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head

                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Head", arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserCharTodo(tUser, UserList(tUser).Char.body, val(arg1), UserList( _
                                    tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList( _
                                    tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.RenegadosMatados = Var

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level

                        If val(arg1) > STAT_MAXELV Then
                            arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & _
                                    ".", FONTTYPE_INFO)

                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                        If val(arg1) >= 25 Then
                            
                            Dim Gi As Integer

                            If tUser <= 0 Then
                                Gi = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
                            Else
                                Gi = UserList(tUser).GuildIndex

                            End If
                            
                            If Gi > 0 Then
                                If modGuilds.GuildAlignment(Gi) = "Del Mal" Or modGuilds.GuildAlignment(Gi) = "Real" _
                                        Then
                                    'We get here, so guild has factionary alignment, we have to expulse the user
                                    Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                                    
                                    Call SendData(SendTarget.ToGuildMembers, Gi, PrepareMessageConsoleMsg(UserName & _
                                            " deja el clan.", FontTypeNames.FONTTYPE_GUILD))

                                    ' Si esta online le avisamos
                                    If tUser > 0 Then Call WriteConsoleMsg(tUser, _
                                            "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", _
                                            FontTypeNames.FONTTYPE_GUILD)

                                End If

                            End If

                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "ELV", val(arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.ELV = val(arg1)
                            Call WriteUpdateExp(tUser)

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class

                        For loopc = 1 To NUMCLASES

                            If UCase$(ListaClases(loopc)) = UCase$(arg1) Then Exit For
                        Next loopc
                            
                        If loopc > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else

                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "INIT", "Clase", loopc)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                        FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Clase = loopc

                            End If

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills

                        For loopc = 1 To NUMSKILLS

                            If UCase$(Replace$(SkillsNames(loopc), " ", "+")) = UCase$(arg1) Then Exit For
                        Next loopc
                        
                        If loopc > NUMSKILLS Then
                            Call WriteLocaleMsg(UserIndex, 414)
                        Else

                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "Skills", "SK" & loopc, arg2)
                                Call WriteVar(UserCharPath, "Skills", "EXPSK" & loopc, 0)
                                
                                If arg2 < MAXSKILLPOINTS Then
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & loopc, ELU_SKILL_INICIAL * 1.05 ^ _
                                            arg2)
                                Else
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & loopc, 0)

                                End If
    
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                        FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(loopc) = val(arg2)
                                Call CheckEluSkill(tUser, loopc, True)

                            End If

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft

                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "SkillPtsLibres", arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                    FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.SkillPts = val(arg1)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLSLIBRES "
 
                      Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase(arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                        Sex = IIf(UCase(arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                        If Sex <> 0 Then ' Es Hombre o mujer?
                            If tUser <= 0 Then ' OffLine
                                Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                        FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Genero = Sex

                            End If

                        Else
                            Call WriteLocaleMsg(UserIndex, 414)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        
                        arg1 = UCase$(arg1)

                        Select Case arg1

                            Case "HUMANO"
                                raza = eRaza.Humano

                            Case "ELFO"
                                raza = eRaza.Elfo

                            Case "DROW"
                                raza = eRaza.Drow

                            Case "ENANO"
                                raza = eRaza.enano

                            Case "GNOMO"
                                raza = eRaza.gnomo

                            Case "ORCO"
                                raza = eRaza.Orco
                                
                            Case Else
                                raza = 0

                        End Select
                            
                        If raza = 0 Then
                            Call WriteLocaleMsg(UserIndex, 414)
                        Else

                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Raza", raza)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, _
                                        FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza

                            End If

                        End If
                            
                        ' Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim bankGold As Long
                        
                        If Abs(arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & _
                                    MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If tUser <= 0 Then
                                bankGold = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                                Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(arg1) <= 0, 0, _
                                        bankGold + val(arg1)))
                                Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & arg1 & " monedas de oro a " & _
                                        UserName & ".", FONTTYPE_TALK)
                            Else
                                UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(arg1) <= 0, 0, _
                                        UserList(tUser).Stats.Banco + val(arg1))
                                Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)

                            End If

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "AGREGAR "
                        
                    Case Else
                        Call WriteLocaleMsg(UserIndex, 394, vbNullString, 1)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & arg1 & " " & arg2
                Call LogGM(.Name, CommandString & " " & UserName)
                
            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

    Exit Sub

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid).. alto bug zapallo..
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
                
        Dim targetName  As String
        Dim targetIndex As Integer
        
        targetName = Replace$(Buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            'is the player offline?
            If targetIndex <= 0 Then
                    Call SendUserStatsTxtOFF(UserIndex, targetName)
 

            Else

                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or _
                        PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            If tUser <= 0 Then
                
                Call SendUserInvTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserInvTxt(UserIndex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim loopc    As Long
        Dim message  As String
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")

                End If
                
                For loopc = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(loopc) & " = " & GetVar(CharPath & UserName & ".chr", _
                            "SKILLS", "SK" & loopc) & vbCrLf
                Next loopc
                
                Call WriteConsoleMsg(UserIndex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", _
                        "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim loopc    As Byte
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex

            End If
            
            If tUser <= 0 Then
                Call WriteLocaleMsg(UserIndex, 75)
            Else

                With UserList(tUser)

                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        Call DarCuerpoDesnudo(tUser)

                        Call ChangeUserCharTodo(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, _
                                .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", _
                                FontTypeNames.FONTTYPE_INFO)
                    Else
                    
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", _
                                FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    .Stats.MinHP = .Stats.MaxHP
                    
                
                End With
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.Name, "Resucito a " & UserName)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 12/28/06
    '
    '***************************************************
    Dim i    As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or _
                PlayerType.Admin
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).Name & ", "

            End If

        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteLocaleMsg(UserIndex, 410)

        End If

    End With

End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 23/03/2009
    '23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim loopc As Long
        Dim list  As String
        Dim priv  As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or _
                PlayerType.Admin)
        
        For loopc = 1 To LastUser

            If LenB(UserList(loopc).Name) <> 0 And UserList(loopc).Pos.Map = Map Then
                If UserList(loopc).flags.Privilegios And priv Then list = list & UserList(loopc).Name & ", "

            End If

        Next loopc
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim rank     As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteLocaleMsg(UserIndex, 75)
            Else

                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteLocaleMsg(UserIndex, 298)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echó a " & UserName & ".", _
                            FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echó a " & UserName)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
 
        If EsGm(UserIndex) Then
            If OnlineCheck(UserIndex, tUser) Then
                Call UserDie(tUser)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(458, UserName & "%" & .Name, 5))
                Call LogGM(.Name, " ejecuto a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
        
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim Accion   As Byte
        
        UserName = Buffer.ReadASCIIString()
        Accion = Buffer.ReadByte()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
        
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
        
            Select Case Accion
            
                Case 1  'Unban
                
                    Call BanCharacter(UserIndex, UserName)
                    
                
                Case 0  ' Banea
                
                    Call UnBanCharacter(UserIndex, UserName)
                
            End Select
 
        End If

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub
 
''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0

        End If

    End With

End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim X        As Integer
        Dim Y        As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteLocaleMsg(UserIndex, 75)
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList( _
                        tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteLocaleMsg(UserIndex, 298)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)

    End With

End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
 
    End With

End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & "> " & _
                        message, FontTypeNames.FONTTYPE_TALK))
                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).Name & " > " & message

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 24/07/07
    'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        Dim priv     As PlayerType
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or _
                        PlayerType.Admin
            Else
                priv = PlayerType.User

            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, _
                            FontTypeNames.FONTTYPE_INFO)
                    Dim ip    As String
                    Dim lista As String
                    Dim loopc As Long
                    ip = UserList(tUser).ip

                    For loopc = 1 To LastUser

                        If UserList(loopc).ip = ip Then
                            If LenB(UserList(loopc).Name) <> 0 And UserList(loopc).flags.UserLogged Then
                                If UserList(loopc).flags.Privilegios And priv Then
                                    lista = lista & UserList(loopc).Name & ", "

                                End If

                            End If

                        End If

                    Next loopc

                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, _
                            FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteLocaleMsg(UserIndex, 80)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip    As String
        Dim loopc As Long
        Dim lista As String
        Dim priv  As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User

        End If

        For loopc = 1 To LastUser

            If UserList(loopc).ip = ip Then
                If LenB(UserList(loopc).Name) <> 0 And UserList(loopc).flags.UserLogged Then
                    If UserList(loopc).flags.Privilegios And priv Then
                        lista = lista & UserList(loopc).Name & ", "

                    End If

                End If

            End If

        Next loopc
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild    As Integer
        
        GuildName = Buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")

        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline( _
                        UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub
Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler

    ' Verifica que haya suficientes datos en la entrada
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)
        ' Elimina el ID del paquete
        Call .incomingData.ReadByte

        Dim Mapa As Integer
        Dim X As Byte
        Dim Y As Byte

        ' Lee los datos del paquete
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()

        ' Verifica privilegios
        If Not (.flags.Privilegios And PlayerType.Admin) Then
            Call WriteConsoleMsg(UserIndex, "No tienes permisos para usar este comando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Log de comando
        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y)

        ' Verifica mapa y coordenadas
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Mapa o coordenadas inválidas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Verifica si la casilla de destino está ocupada
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en esa ubicación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Crea el objeto de teletransporte
        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 378 ' ID del objeto teletransporte

        ' Crea el objeto en la posición del GM
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)

        ' Establece el destino del teleport
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With

        Call WriteConsoleMsg(UserIndex, "Teletransporte creado correctamente.", FontTypeNames.FONTTYPE_INFO)
    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.description, "HandleTeleportCreate", Erl)
    Resume Next
End Sub

End Sub


''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        Dim Mapa As Integer
        Dim X    As Byte
        Dim Y    As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)

            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 378 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)

                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0

            End If

        End With

End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
       If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    With UserList(UserIndex)
       
        
        'Remove packet ID
        Call .incomingData.ReadByte
        Queclima = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If Queclima = 1 Then ' Lanzamos lluvia común
            Call LogGM(.Name, "/CLIMA 0")
            Lloviendo = Not Lloviendo
            
            'Call SendData(ToAll, 0, PrepareMessagePlayWave(105, RandomNumber(1, 100), RandomNumber(1, 100)))
               Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(1))
            Exit Sub
        ElseIf Queclima = 2 Then 'Activamos Lluvia elecitrca
            Call LogGM(.Name, "/CLIMA 1")
            Lloviendo = Not Lloviendo
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(2))
          '  If Lloviendo Then
           '   '  Call SendData(ToAll, 0, PrepareMessagePlayWave(105, RandomNumber(1, 100), RandomNumber(1, 100)))
            '    Exit Sub
           ' End If
         ElseIf Queclima = 3 Then ' Nieve
            Call LogGM(.Name, "/CLIMA 3")
            Lloviendo = Not Lloviendo
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(3))
            Exit Sub
        Else
        Lloviendo = Not Lloviendo
        Queclima = 0
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(0))
        Exit Sub
        End If
        
        
    End With
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex))
            Else
                Call WriteLocaleMsg(UserIndex, 22)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X       As Long
        Dim Y       As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)

                        End If

                    End If

                End If

            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")

    End With

End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser    As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And ( _
                PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteLocaleMsg(UserIndex, 75)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog     As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    '
    '***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & _
                .Pos.Y, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim loopc As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For loopc = 1 To BanIps.count
            lista = lista & BanIps.Item(loopc) & ", "
        Next loopc
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar

    End With

End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim GuildName   As String
        Dim cantMembers As Integer
        Dim loopc       As Long
        Dim member      As String
        Dim count       As Byte
        Dim tIndex      As Integer
        Dim tFile       As String
        
        GuildName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneó al clan " & UCase$( _
                        GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For loopc = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & loopc)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & _
                            "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)

                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)

                    End If
                    
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & count + 1, LCase$(.Name) & _
                            ": BAN AL CLAN: " & GuildName & " " & Date & " " & Time)
                Next loopc

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/02/09
    'Agregado un CopyBuffer porque se producia un bucle
    'inifito al intentar banear una ip ya baneada. (NicoNZ)
    '07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
    '***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser    As Integer
        Dim reason   As String
        Dim i        As Long
        
        ' Is it by ip??
        If Buffer.ReadBoolean() Then
            bannedIP = Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte()
        Else
            tUser = NameIndex(Buffer.ReadASCIIString())
            
            If tUser > 0 Then bannedIP = UserList(tUser).ip

        End If
        
        reason = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", _
                            FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser

                        If UserList(i).ConnIDValida Then
                            If UserList(i).ip = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name)

                            End If

                        End If

                    Next i

                End If

            ElseIf tUser <= 0 Then
                Call WriteLocaleMsg(UserIndex, 75)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", _
                    FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", _
                    FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub
Private Sub HandleCreateItem(ByVal UserIndex As Integer)

       On Error GoTo ErrHandler
       
        If UserList(UserIndex).incomingData.length < 5 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
   
        With UserList(UserIndex)
    
        Call .incomingData.ReadByte

        
        Dim tObj   As Integer
        Dim Objeto As Obj

        tObj = .incomingData.ReadInteger
        Objeto.Amount = .incomingData.ReadInteger()
        Objeto.ObjIndex = tObj
            
        If .flags.Privilegios And PlayerType.Dios Then
  
            If tObj >= 1 And tObj <= NumObjDatas Then
                If LenB(ObjData(tObj).Name) > 0 Then
                    If MeterItemEnInventario(UserIndex, Objeto) Then
                        Call WriteLocaleMsg(UserIndex, 446, ObjData(tObj).Name & "%" & Objeto.Amount)
                        Call LogGM(.Name, "/CI: " & tObj)
                    End If
                End If
            End If
        
        End If

   End With

   Exit Sub

ErrHandler:

    Call LogError("Error en HandleCreateItem en " & Erl & ". Err " & Err.Number & " " & Err.description)

End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST")
        
        If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And MapData( _
                .Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then

            Exit Sub

        End If
        
        Call EraseObj(10000, .Pos.Map, .Pos.X, .Pos.Y)

    End With

End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0

        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)

    End With

End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)

    End With

End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).npcindex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).npcindex)

                End If

            Next X
        Next Y

        Call LogGM(.Name, "/MASSKILL")

    End With

End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName   As String
        Dim lista      As String
        Dim loopc      As Byte
        Dim priv       As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then

            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")

            End If

            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")

            End If

            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")

            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios _
                        And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And ( _
                        PlayerType.Admin Or PlayerType.Dios)) <> 0

            End If
            
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"

                    For loopc = 1 To 5
                      '  lista = lista & vbCrLf & loopc & " - " & GetVar(CharPath & UserName & ".chr", "INIT", _
                                "LastIP" & loopc)
                    Next loopc

                   ' Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                  '  Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", _
                            FontTypeNames.FONTTYPE_INFO)

                End If

            Else
               ' Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", _
                        FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the objects
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        
        Call LoadOBJData

    End With

End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the spells
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos

    End With

End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s INI
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini

    End With

End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s NPC
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
        Call WriteLocaleMsg(UserIndex, 416)
    End With

End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Kick all the chars that are online
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados

    End With

End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Clean the SOS
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        
        Call Ayuda.Reset

    End With

End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)

        With UserList(UserIndex)
        
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsGm(UserIndex) Then
            Call GuardarUsuarios
            Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        
        End If
        
        End With

End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the backup`s info of the map
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0

        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo( _
                .Pos.Map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, _
                FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Change the pk`s info of the  map
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, _
                FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la información sobre si es restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList( _
                        UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & MapInfo(.Pos.Map).Restringir, _
                        FontTypeNames.FONTTYPE_INFO)
            'Else
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje329)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'MagiaSinEfecto -> Options: "1" , "0".
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList( _
                    UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo( _
                    .Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'InviSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & _
                    " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList( _
                    UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo( _
                    .Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'ResuSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & _
                    " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList( _
                    UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo( _
                    .Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = _
                    "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información del terreno del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList( _
                        UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).Terreno, _
                        FontTypeNames.FONTTYPE_INFO)
            Else
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje330)
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje331)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = _
                    "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList( _
                        UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, _
                        FontTypeNames.FONTTYPE_INFO)
            Else
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje330)
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje332)
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Saves the map
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteLocaleMsg(UserIndex, 417)

    End With

End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Allows admins to read guild messages
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim guild As String
        
        guild = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/26/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Activate or desactivate the Centinel
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        With Centinela
            .RevisandoUserIndex = 0
            .clave = 0
            .TiempoRestante = 0

        End With
    
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0

        End If
        
        If centinelaActivado Then
            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", _
                    FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", _
                    FontTypeNames.FONTTYPE_SERVER))

        End If

    End With

End Sub
 
''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contraseña de " & UserName & " cuya cuenta es: " & .Account)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje339)
            Else

                If Not FileExist(AccountPath & .Account & ".cnt") Or Not FileExist(AccountPath & copyFrom & ".cnt") Then
                    Call WriteConsoleMsg(UserIndex, "La cuenta no existe " & UserName & "@" & copyFrom, _
                            FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(AccountPath & copyFrom & ".cnt", .Account, "Password")
                    Call WriteVar(AccountPath & .Account & ".cnt", .Account, "Password", Password)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " cuya cuenta es: " & .Account & " ha cambiado por la de " & copyFrom, _
                            FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim npcindex As Integer
        
        npcindex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        npcindex = SpawnNpc(npcindex, .Pos, True, False)
        
        If npcindex <> 0 Then
            Call LogGM(.Name, "Sumoneó a " & Npclist(npcindex).Name & " en mapa " & .Pos.Map)

        End If

    End With

End Sub

''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim npcindex As Integer
        
        npcindex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        npcindex = SpawnNpc(npcindex, .Pos, True, True)
        
        If npcindex <> 0 Then
            Call LogGM(.Name, "Sumoneó con respawn " & Npclist(npcindex).Name & " en mapa " & .Pos.Map)

        End If

    End With

End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/12/07
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1

        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)

    End With

End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteLocaleMsg(UserIndex, 418)
            ServerSoloGMs = 0
        Else
            Call WriteLocaleMsg(UserIndex, 419)
            ServerSoloGMs = 1

        End If

    End With

End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    'Turns off the server
    '***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or _
                PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡¡" & .Name & " VA A APAGAR EL SERVIDOR!!!", _
                FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & Time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain

    End With

End Sub


''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName   As String
        Dim GuildIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje342)
            Else
                'Call WriteMensajes(UserIndex, eMensajes.Mensaje343)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & _
                        " ha sido expulsado del clan por los administradores del servidor.", _
                        FontTypeNames.FONTTYPE_GUILD))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/29/06
    'Send a message to all the users
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or _
                PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
        If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
       Dim Time As Long
        
       Time = .incomingData.ReadLong()
        Call WritePong(UserIndex, Time)

    End With

End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Brian Chaia (BrianPr)
    'Last Modification: 01/23/10 (Marco)
    'Modify server.ini
    '***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String

        'Obtengo los parámetros
        sLlave = Buffer.ReadASCIIString()
        sClave = Buffer.ReadASCIIString()
        sValor = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteLocaleMsg(UserIndex, 10)
            Else
                'Obtengo el valor según llave y clave
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & _
                            " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modificó " & sLlave & " " & sClave & " a " & sValor & _
                            ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteLocaleMsg(UserIndex, 414)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

ErrHandler:
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Logged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Logged)
    UserList(UserIndex).Redundance = RandomNumber(15, 250)
    Call UserList(UserIndex).outgoingData.WriteByte(UserList(UserIndex).Redundance)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MontateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMontateToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WriteMontateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MontateToggle)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer, ByVal goliath As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(UserIndex).outgoingData.WriteByte(goliath)
    
    If goliath = 1 Then
        Call UserList(UserIndex).outgoingData.WriteLong(UserList(UserIndex).Stats.Banco)
        Call UserList(UserIndex).outgoingData.WriteByte(UserList(UserIndex).BancoInvent.NroItems)
    Else
        Call UserList(UserIndex).outgoingData.WriteLong(0)
        Call UserList(UserIndex).outgoingData.WriteByte(0)
    End If
Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
 
''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
Public Sub WriteUpdateGold(ByVal UserIndex As Integer)


   On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
 
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budsi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
        
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
    End With
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
        
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal ModeChat As Byte = 0)

    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, ModeChat))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteChatOverHeadLocale(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal index As Long, ByVal Modo As Byte)

    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHeadLocale(CharIndex, index, Modo))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As Byte)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat, FontIndex))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
            
''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String, Optional ByVal EsPregunta As Boolean = False, Optional ByVal Accion As Byte = 0)

    On Error GoTo ErrHandler
    
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageShowMessageBox(message, EsPregunta, Accion))

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
                                ByVal body As Integer, _
                                ByVal Head As Integer, _
                                ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal X As Byte, _
                                ByVal Y As Byte, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal Name As String, _
                                ByVal Privileges As Byte, _
                                ByVal Donador As Byte, ByVal ParticulaFx As Byte, ByVal Arma_Aura As Byte, ByVal Body_Aura As Byte, ByVal Escudo_Aura As Byte, ByVal Head_Aura As Byte, ByVal Otra_Aura As Byte, ByVal Anillo_Aura As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, _
            CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, Name, Privileges, Donador, ParticulaFx, Arma_Aura, Body_Aura, Escudo_Aura, Head_Aura, Otra_Aura, Anillo_Aura))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Desvanecido As Boolean)


    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex, Desvanecido))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal CharIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
                                ByVal body As Integer, _
                                ByVal Head As Integer, _
                                ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer)

    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
Public Sub WriteCharacterChangeSlot(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal SlotIndex As Integer, ByVal index As Byte)


    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChangeSlot(CharIndex, SlotIndex, index))
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub


Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(X, Y, ObjIndex, Amount))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, _
                         ByVal midi As Byte, _
                         Optional ByVal Loops As Integer = -1)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, Loops))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, _
                         ByVal wave As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)


    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim Tmp As String
    Dim i   As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer, ByVal clima As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle(clima))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
                         ByVal CharIndex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)


    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteInteger(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
Public Sub WriteUpdateUserStatsForLevel(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStatsForLevel)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal slot As Byte)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        
        Dim ObjIndex As Integer
        Dim obData   As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex

        If ObjIndex > 0 Then obData = ObjData(ObjIndex)
        
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(slot)
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(slot).Equipped)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteByte(IIf(obData.MinELV < UserList(UserIndex).Stats.ELV And SexoPuedeUsarItem(UserIndex, ObjIndex) = True And FaccionPuedeUsarItem(UserIndex, ObjIndex) = True And ClasePuedeUsarItem(UserIndex, ObjIndex) = True, 1, 0))
        
        End With
    
     Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal slot As Byte)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(slot).Amount)
        
        Call .WriteLong(obData.Valor)
        
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal slot As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(slot))
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData

        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria)) Then
                If Not (ObjData(ArmasHerrero(i)).LingH = 0 And ObjData(ArmasHerrero(i)).LingO = 0 And ObjData(ArmasHerrero(i)).LingP = 0) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
    
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria)) Then
                If Not (ObjData(ArmadurasHerrero(i)).LingH = 0 And ObjData(ArmadurasHerrero(i)).LingO = 0 And ObjData(ArmadurasHerrero(i)).LingP = 0) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithHelmet(ByVal UserIndex As Integer)
'***************************************************
'Author: $ Shermie80
'
'***************************************************
On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(CascosHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithHelmet)
        
        For i = 1 To UBound(CascosHerrero())
        
            If ObjData(CascosHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria)) Then
                If Not (ObjData(CascosHerrero(i)).LingH = 0 And ObjData(CascosHerrero(i)).LingO = 0 And ObjData(CascosHerrero(i)).LingP = 0) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
            
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(CascosHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(CascosHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBlacksmithShield(ByVal UserIndex As Integer)
'***************************************************
'Author: $ Shermie80
'
'***************************************************
On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(EscudosHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithShield)
        
        For i = 1 To UBound(EscudosHerrero())
        
            If ObjData(EscudosHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria)) Then
                If Not (ObjData(EscudosHerrero(i)).LingH = 0 And ObjData(EscudosHerrero(i)).LingO = 0 And ObjData(EscudosHerrero(i)).LingP = 0) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
            
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(EscudosHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(EscudosHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjCarpintero(i) <> 0 Then
                If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).Clase) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAlquimiaObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjDruida()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlquimiaObjects)
        
        For i = 1 To UBound(ObjDruida())
            ' Can the user create this object? If so add it to the list....
            If ObjDruida(i) <> 0 Then
                If ObjData(ObjDruida(i)).SkPociones <= UserList(UserIndex).Stats.UserSkills(eSkill.alquimia) \ Modalquimia(UserList(UserIndex).Clase) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            Obj = ObjData(ObjDruida(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.Raies)
            Call .WriteInteger(ObjDruida(validIndexes(i)))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteTejiblesObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjSastre()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SastreObjects)
        
        For i = 1 To UBound(ObjSastre())
   
            ' Can the user create this object? If so add it to the list....
            If ObjSastre(i) <> 0 Then

                If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) \ ModSastreria(UserList(UserIndex).Clase) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
  
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
              
        For i = 1 To count

            Obj = ObjData(ObjSastre(validIndexes(i)))
            Call .WriteInteger(Obj.Numero)
            Call .WriteInteger(Obj.PielLobo)
            Call .WriteInteger(Obj.PielOsoPardo)
            Call .WriteInteger(Obj.PielOsoPolar)
            Call .WriteInteger(ObjSastre(validIndexes(i)))
        Next i
        
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSendMsgBox(ByVal UserIndex As Integer, ByVal message As String, Optional ByVal Modo As Byte = 0)


    On Error GoTo ErrHandler
    
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSendMsgBox(message, Modo))
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
                                       ByVal slot As Byte, _
                                       ByRef Obj As Obj, _
                                       ByVal price As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Last Modified by: Budi
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
    '***************************************************
    On Error GoTo ErrHandler

    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(slot)
        Call .WriteInteger(Obj.Amount)
        Call .WriteSingle(price)
        Call .WriteInteger(Obj.ObjIndex)
     End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
Public Sub WriteMiniStats(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.RenegadosMatados)
        Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        Call .WriteByte(UserList(UserIndex).Clase)
        Call .WriteByte(UserList(UserIndex).raza)
        Call .WriteByte(UserList(UserIndex).Genero)
        
        Call .WriteLong(UserList(UserIndex).flags.MuertesUsuario)
        
        Call .WriteByte(UserList(UserIndex).Faccion.Status)
        Call .WriteLong(UserList(UserIndex).Faccion.RepublicanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CaosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.ArmadaMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.MilicianosMatados)
        
        
    End With
Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
                             ByVal CharIndex As Integer, _
                             ByVal Invisible As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, Invisible))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/03/2020; Shermie80, _
                                querido diario hoy codie como un negro.
'Writes the "SendSkills" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal npcindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(npcindex).NroCriaturas
            str = str & Npclist(npcindex).Criaturas(i).npcindex & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, _
                          ByVal guildNews As String, _
                          ByRef enemies() As String, _
                          ByRef allies() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Tmp = vbNullString

        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
 ' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class _
        As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As _
        String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal _
        criminalsKilled As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(Gold)
        Call .WriteLong(bank)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, _
                                ByRef guildList() As String, _
                                ByRef MemberList() As String, _
                                ByVal guildNews As String, _
                                ByRef joinRequests() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString

        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserIndex As Integer, _
                                ByRef guildList() As String, _
                                ByRef MemberList() As String)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, _
                             ByVal GuildName As String, _
                             ByVal founder As String, _
                             ByVal foundationDate As String, _
                             ByVal leader As String, _
                             ByVal URL As String, _
                             ByVal memberCount As Integer, _
                             ByVal electionsOpen As Boolean, _
                             ByVal alignment As String, _
                             ByVal enemiesCount As Integer, _
                             ByVal AlliesCount As Integer, _
                             ByVal antifactionPoints As String, _
                             ByRef codex() As String, _
                             ByVal guildDesc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i    As Long
    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(antifactionPoints)
        
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo ErrHandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    'Call WritePosUpdate(UserIndex)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        Call .WriteASCIIString(details)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TradeOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

        Dim i As Long
        Dim Tmp As String
    
100     With UserList(UserIndex).outgoingData

102         Call .WriteByte(ServerPacketID.ShowSOSForm)
        
104         For i = 1 To Ayuda.Longitud
106             Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
108         Next i
        
110         If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
 
112         Call .WriteASCIIString(Tmp)
        
            Dim tInicio As Integer
            Dim tFin As Integer
            Dim j As Integer
            Dim tmpStr As String
            Dim tmpsplit() As String
            
            tmpsplit = Split(Tmp, "=")
            
            tInicio = 1
            tFin = 1
            '[MERMAS]-19:00:19-5= [MERMAS]-19:00:35-5= [GUSANIO]-19:04:26-5=
            
            For j = 0 To (UBound(tmpsplit) - 1)

                tInicio = InStr(tFin, tmpsplit(j), "[")
                tFin = InStr(tInicio, tmpsplit(j), "]")
                
                tmpStr = mid(tmpsplit(j), tInicio, tFin)
                Debug.Print tmpStr
                
            Next j
            

        End With
        
        Exit Sub

ErrHandler:

114     If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
116         Resume
        End If

End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
                             ByRef userNamesList() As String, _
                             ByVal cant As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal Time As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
    Call UserList(UserIndex).outgoingData.WriteLong(Time)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************
    Dim sndData As String
    
    With UserList(UserIndex).outgoingData

        If .length = 0 Then Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(UserIndex, sndData)

    End With

End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal Invisible As Boolean) As String
    
    On Error GoTo PrepareMessageSetInvisible_Err

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageSetInvisible_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageSetInvisible", Erl)
     Resume Next
        
End Function

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As Integer) As String

On Error GoTo PrepareMessageConsoleMsg_Err

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
    
    Exit Function

PrepareMessageConsoleMsg_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageConsoleMsg", Erl)
     Resume Next
        
End Function
Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal ModeChat As Byte = 0) As String

    On Error GoTo PrepareMessageChatOverHead_Err
   
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(ModeChat)
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageChatOverHead_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageChatOverHead", Erl)
     Resume Next
        
End Function

Public Function PrepareMessageChatOverHeadLocale(ByVal CharIndex As Integer, ByVal index As Long, ByVal Modo As Byte) As String

    On Error GoTo PrepareMessageChatOverHeadLocale_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PrepareMessageChatOverHead" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHeadLocale)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(index)
        Call .WriteByte(Modo)
        PrepareMessageChatOverHeadLocale = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageChatOverHeadLocale_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageChatOverHeadLocale", Erl)
     Resume Next
        
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, _
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer) As String

    On Error GoTo PrepareMessageCreateFX_Err
    

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageCreateFX_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCreateFX", Erl)
     Resume Next
        
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Integer, _
                                       ByVal X As Byte, _
                                       ByVal Y As Byte) As String

    On Error GoTo PrepareMessagePlayWave_Err


    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteInteger(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessagePlayWave_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePlayWave", Erl)
     Resume Next
        
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As String

    On Error GoTo PrepareMessageGuildChat_Err

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)

    End With
    
    Exit Function

PrepareMessageGuildChat_Err:
108     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageGuildChat", Erl)
110     Resume Next
        
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal message As String, Optional ByVal EsPregunta As Boolean = False, Optional ByVal Accion As Byte = 0) As String
    
    On Error GoTo PrepareMessageShowMessageBox_Err

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
        Call .WriteBoolean(EsPregunta)
        Call .WriteByte(Accion)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)

    End With

    Exit Function
  
PrepareMessageShowMessageBox_Err:
108     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageShowMessageBox", Erl)
110     Resume Next
        
End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, _
                                       Optional ByVal Loops As Integer = -1) As String

    On Error GoTo PrepareMessagePlayMidi_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "GuildChat" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteByte(midi)
        Call .WriteInteger(Loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)

    End With
        
    Exit Function

PrepareMessagePlayMidi_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePlayMidi", Erl)
     Resume Next
        
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String

    On Error GoTo PrepareMessagePauseToggle_Err
        
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PauseToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)

    End With
    
    Exit Function

PrepareMessagePauseToggle_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePauseToggle", Erl)
     Resume Next
        
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle(ByVal clima As Byte) As String

On Error GoTo PrepareMessageRainToggle_Err

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        Call .WriteByte(clima)
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)
    End With
    
    Exit Function

PrepareMessageRainToggle_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageRainToggle", Erl)
     Resume Next
        
End Function


''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String

    On Error GoTo PrepareMessageObjectDelete_Err

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageObjectDelete_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageObjectDelete", Erl)
     Resume Next
        
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, _
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Boolean) As String
                                            
    On Error GoTo PrepareMessageBlockPosition_Err


    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)

    End With
    
    Exit Function

PrepareMessageBlockPosition_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageBlockPosition", Erl)
     Resume Next
        
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal X As Byte, ByVal Y As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer) As String

    On Error GoTo PrepareMessageObjectCreate_Err
    
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(Amount)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageObjectCreate_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageObjectCreate", Erl)
     Resume Next
        
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer, ByVal Desvanecido As Boolean) As String
    
    On Error GoTo PrepareMessageCharacterRemove_Err
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Desvanecido)
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageCharacterRemove_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterRemove", Erl)
     Resume Next
        
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
    
    On Error GoTo PrepareMessageRemoveCharDialog_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)

    End With
    
    Exit Function

PrepareMessageRemoveCharDialog_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageRemoveCharDialog", Erl)
     Resume Next
        
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal X As Byte, _
                                              ByVal Y As Byte, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Name As String, _
                                              ByVal Privileges As Byte, _
                                              ByVal Donador As Byte, ByVal ParticulaFx As Byte, ByVal Arma_Aura As Byte, ByVal Body_Aura As Byte, ByVal Escudo_Aura As Byte, ByVal Head_Aura As Byte, ByVal Otra_Aura As Byte, ByVal Anillo_Aura As Byte) As String

    On Error GoTo PrepareMessageCharacterCreate_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(Privileges)
        Call .WriteByte(Donador)
        Call .WriteByte(ParticulaFx)
        Call .WriteByte(Arma_Aura)
        Call .WriteByte(Body_Aura)
        Call .WriteByte(Escudo_Aura)
        Call .WriteByte(Head_Aura)
        Call .WriteByte(Otra_Aura)
        Call .WriteByte(Anillo_Aura)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageCharacterCreate_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterCreate", Erl)
     Resume Next
        
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer) As String

    On Error GoTo PrepareMessageCharacterChange_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterChange" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageCharacterChange_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterChange", Erl)
     Resume Next
        
End Function

Public Function PrepareMessageCharacterChangeSlot(ByVal CharIndex As Integer, ByVal SlotIndex As Integer, ByVal index As Byte) As String

    On Error GoTo PrepareMessageCharacterChangeSlot_Err
    
    
    With auxiliarBuffer
        
        Call .WriteByte(ServerPacketID.CharacterChangeSlot)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(SlotIndex)
        Call .WriteByte(index)
        
        PrepareMessageCharacterChangeSlot = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageCharacterChangeSlot_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterChangeSlot", Erl)
     Resume Next
        
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, _
                                            ByVal X As Byte, _
                                            ByVal Y As Byte) As String

    On Error GoTo PrepareMessageCharacterMove_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageCharacterMove_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterMove", Erl)
     Resume Next
        
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String

    On Error GoTo PrepareMessageForceCharMove_Err
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Prepares the "ForceCharMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageForceCharMove_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageForceCharMove", Erl)

        
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
                                                 ByRef Tag As String, _
                                                 ByVal Status As Byte, Optional ByVal Donador As Byte = 0) As String

    On Error GoTo PrepareMessageUpdateTagAndStatus_Err
    
    '***************************************************
    'Author: Alejandro Salvo (Salvito)
    'Last Modification: 04/07/07
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    'Prepares the "UpdateTagAndStatus" message and returns it
    '15/01/2010: ZaMa - Now sends the nick color instead of the status.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteASCIIString(Tag)
        Call .WriteByte(Status)
        Call .WriteByte(Donador)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageUpdateTagAndStatus_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageUpdateTagAndStatus", Erl)
     Resume Next
        
End Function

Public Function PrepareMessageSendMsgBox(ByVal message As String, Optional ByVal Modo As Byte = 0) As String
    
    On Error GoTo PrepareMessageMsgBox_Err
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SendMsgBox)
        Call .WriteASCIIString(message)
        Call .WriteByte(Modo)

        
        PrepareMessageSendMsgBox = .ReadASCIIStringFixed(.length)

    End With

    Exit Function

PrepareMessageMsgBox_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageMsgBox", Erl)
     Resume Next
        
End Function
Public Function PrepareMessageEfectoCharParticula(ByVal CharIndex As Integer, ByVal Particle As Integer, ByVal Time As Single, ByVal Remove As Boolean, ByVal EsUsuario As Boolean) As String

    On Error GoTo PrepareMessageEfectoCharParticula_Err
    
1    If EsUsuario = True Then Call AgregarParticula(CharIndex, Particle, Time, Remove, True)
    
    With auxiliarBuffer
2        Call .WriteByte(ServerPacketID.EfectoCharParticula)
3        Call .WriteInteger(CharIndex)
4        Call .WriteInteger(Particle)
5        Call .WriteSingle(Time)
6        Call .WriteBoolean(Remove)
         
        PrepareMessageEfectoCharParticula = .ReadASCIIStringFixed(.length)

    End With
    
    Exit Function

PrepareMessageEfectoCharParticula_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageEfectoCharParticula", Erl)
     Resume Next
        
End Function
Public Sub HandleLoginNewAccount(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 10 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
1    On Error GoTo ErrHandler

2    Dim Buffer As New clsByteQueue
3    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
5    'Remove packet ID
4    Call Buffer.ReadByte

6    Dim UserName As String, UserPassword As String, UserCode As String, Version As String
7    Dim MensajeAdvertencia As Integer
    
8    UserName = Buffer.ReadASCIIString()
9    UserPassword = Buffer.ReadASCIIString()
10   UserCode = Buffer.ReadASCIIString()
11   Version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
114  If Not VersionOK(Version) Then
116     Call WriteShowMessageBox(UserIndex, 0, True, 7) 'Juego desactualizado
117     Call FlushBuffer(UserIndex)
118     Call CloseSocket(UserIndex)
119     Exit Sub
120  End If
        
121  If Not CheckDataNewAccount(UserName, SDesencriptar(UserPassword), SDesencriptar(UserCode), MensajeAdvertencia) Then 'Mermas//, si llego acá es porque edito los chqueos para bugear el juego, ej: manda un dato con valor en 0 pasando de largo las condiciones en el cliente
            
122     If MensajeAdvertencia > 0 Then
123         Call WriteSendMsgBox(UserIndex, MensajeAdvertencia, 1)
124     End If
            
125     Call FlushBuffer(UserIndex)
126     Call CloseSocket(UserIndex)
127     Exit Sub

128  End If
    
129  If Not CuentaExiste(UserName) Then
130     Call SaveNewAccount(UserIndex, UserName, SDesencriptar(UserPassword), SDesencriptar(UserCode))
131   Else
132      Call WriteShowMessageBox(UserIndex, 41) 'Ya existe la cuenta.

133   End If
        
134    'If we got here then packet is complete, copy data back to original queue
135    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
ErrHandler:

    Dim Error As Long
    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub
Public Sub HandleLoginAccount(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 14 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName        As String
    Dim UserPassword    As String
    Dim Version         As String
    Dim MacAddress     As String
    Dim HDserial       As Long
    Dim MensajeAdvertencia As Integer
    
    UserName = Buffer.ReadASCIIString()
    UserPassword = Buffer.ReadASCIIString()
    Version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
120 MacAddress = Buffer.ReadASCIIString()
122 HDserial = Buffer.ReadLong()

114     If Not VersionOK(Version) Then
116         Call WriteShowMessageBox(UserIndex, 0, True, 7) 'Juego desactualizado
            Call FlushBuffer(UserIndex)
118         Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        If Not CheckDataLoginAccount(UserName, SDesencriptar(UserPassword), MensajeAdvertencia) Then 'Mermas//, si llego acá es porque edito los chqueos para bugear el juego, ej: manda un dato con valor en 0 pasando de largo las condiciones en el cliente
            
            If MensajeAdvertencia > 0 Then
                Call WriteSendMsgBox(UserIndex, MensajeAdvertencia, 1)
            End If
            
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
    
152     If Not EntrarCuenta(UserIndex, UCase$(LTrim(RTrim(UserName))), UserPassword, MacAddress, HDserial) Then
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        Else
            Call LoginAccountCharfile(UserIndex)
        
        End If
        
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)

 
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error
        
End Sub
Public Sub WriteAddPj(ByVal UserIndex As Integer, ByVal NameUser As String, ByVal NumberOfCharacters As Byte, ByRef Characters() As AccountUser)

    On Error GoTo ErrHandler
 
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddPJ)
        Call .WriteASCIIString(NameUser)
        Call .WriteByte(NumberOfCharacters)
        
        If NumberOfCharacters > 0 Then

            For i = 1 To NumberOfCharacters
                Call .WriteASCIIString(Characters(i).Name)
                Call .WriteInteger(Characters(i).Head)
                Call .WriteInteger(Characters(i).body)
                Call .WriteInteger(Characters(i).casco)
                Call .WriteInteger(Characters(i).weapon)
                Call .WriteInteger(Characters(i).shield)
                Call .WriteByte(Characters(i).nivel)
                Call .WriteByte(Characters(i).Clase)
                Call .WriteInteger(Characters(i).Mapa)
                Call .WriteByte(Characters(i).color)
                Call .WriteBoolean(Characters(i).gameMaster)
            Next i

        End If
    End With
    
ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleSwapObjects(ByVal UserIndex As Integer)
Dim ObjSlot1 As Byte
Dim ObjSlot2 As Byte
Dim tmpUserObj As UserObj
 
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
    With UserList(UserIndex)
        'Leemos el paquete
        Call .incomingData.ReadByte
       
        ObjSlot1 = .incomingData.ReadByte
        ObjSlot2 = .incomingData.ReadByte
       
        'Cambiamos si alguno es un anillo
        If .Invent.AnilloEqpSlot = ObjSlot1 Then
            .Invent.AnilloEqpSlot = ObjSlot2
        ElseIf .Invent.AnilloEqpSlot = ObjSlot2 Then
            .Invent.AnilloEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un armor
        If .Invent.ArmourEqpSlot = ObjSlot1 Then
            .Invent.ArmourEqpSlot = ObjSlot2
        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
            .Invent.ArmourEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un barco
        If .Invent.BarcoSlot = ObjSlot1 Then
            .Invent.BarcoSlot = ObjSlot2
        ElseIf .Invent.BarcoSlot = ObjSlot2 Then
            .Invent.BarcoSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un casco
        If .Invent.CascoEqpSlot = ObjSlot1 Then
            .Invent.CascoEqpSlot = ObjSlot2
        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
            .Invent.CascoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un escudo
        If .Invent.EscudoEqpSlot = ObjSlot1 Then
            .Invent.EscudoEqpSlot = ObjSlot2
        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
            .Invent.EscudoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es munición
        If .Invent.MunicionEqpSlot = ObjSlot1 Then
            .Invent.MunicionEqpSlot = ObjSlot2
        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
            .Invent.MunicionEqpSlot = ObjSlot1
        End If
        
              'Cambiamos si alguno es un arma
        If .Invent.NudiEqpSlot = ObjSlot1 Then
            .Invent.NudiEqpSlot = ObjSlot2
        ElseIf .Invent.NudiEqpSlot = ObjSlot2 Then
            .Invent.NudiEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un arma
        If .Invent.WeaponEqpSlot = ObjSlot1 Then
            .Invent.WeaponEqpSlot = ObjSlot2
        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
            .Invent.WeaponEqpSlot = ObjSlot1
        End If
        
                'Cambiamos si alguno es un montura
        If .Invent.MonturaSlot = ObjSlot1 Then
            .Invent.MonturaSlot = ObjSlot2
        ElseIf .Invent.MonturaSlot = ObjSlot2 Then
            .Invent.MonturaSlot = ObjSlot1
        End If

        
        'Cambiamos si alguno es un item magico
        If .Invent.MagicSlot = ObjSlot1 Then
            .Invent.MagicSlot = ObjSlot2
        ElseIf .Invent.MagicSlot = ObjSlot2 Then
            .Invent.MagicSlot = ObjSlot1
        End If
 
 
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, UserIndex, ObjSlot1)
        Call UpdateUserInv(False, UserIndex, ObjSlot2)
    End With
End Sub
 
Public Sub HandleResponderGM(ByVal UserIndex As Integer)
     If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
 
        Dim UserName As String
        Dim tIndex As Integer
        Dim MensajeUser As String
        Dim todos As String
        
        UserName = Buffer.ReadASCIIString()
        MensajeUser = Buffer.ReadASCIIString()
        todos = Buffer.ReadASCIIString()
        tIndex = NameIndex(UserName)
 

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
        
                    
        If todos = "Usuario" Then
           If Not tIndex > 0 Then
              Call WriteLocaleMsg(UserIndex, 75)
               If CantSendCorreo(UserIndex, UserName) Then
                      If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
                    Call mod_Correos.EnviarCorreo(UserIndex, UserName, UserList(UserIndex).Name & ":" & MensajeUser, 0, 0, 1)
                    Call WriteConsoleMsg(UserIndex, "Se le ha enviado por correo la respuesta enviada.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                Call WriteConsoleMsg(UserIndex, "El usuario tiene su casilla de correos llena, por favor comunicate con el usuario a la brevedad.", FontTypeNames.FONTTYPE_INFO)
                End If
           Else

           If MapInfo(UserList(tIndex).Pos.Map).Pk = True Then
              Call WriteLocaleMsg(UserIndex, 426)
                If CantSendCorreo(UserIndex, UserName) Then
                Call mod_Correos.EnviarCorreo(UserIndex, UserName, UserList(UserIndex).Name & ":" & MensajeUser, 0, 0, 1)
                Else
                Call WriteConsoleMsg(UserIndex, "El usuario tiene su casilla de correos llena, por favor comunicate con el usuario a la brevedad.", FontTypeNames.FONTTYPE_INFO)
                End If
           Else
               Call WriteSendMsgBox(tIndex, UserList(UserIndex).Name & ": " & "" & MensajeUser & "") 'responde frm
               Call FlushBuffer(tIndex)
               Call WriteLocaleMsg(UserIndex, 251)
               'UserList(tIndex).flags.EnvioGM = 0 'Reseteamos flag
            End If
          End If
       End If
 
             End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

 
Private Sub HandleRetirarFaccion(ByVal UserIndex As Integer)
        
    On Error GoTo HandleRetirar_Err
        
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.User Then
        
            If DeadCheck(UserIndex) Then Exit Sub
            
            If .Stats.ELV <= 14 Then
                Call WriteLocaleMsg(UserIndex, 425, LimiteNewbie)
                Exit Sub
            End If
        
            If esRene(UserIndex) Or UserList(UserIndex).GuildIndex > 0 Then Exit Sub
            
            If .flags.Meditando Then
                Dim Estado As Byte
                Estado = ParticleToLevel(UserIndex)
            End If
            
            Select Case .Faccion.Status
            
                Case 6
                    ExpulsarFaccionMilicia (UserIndex)
            
                Case 4
                    ExpulsarFaccionCaos (UserIndex)
            
                Case 5
                    ExpulsarFaccionReal (UserIndex)
                    
                Case Else
                
                    .Faccion.Status = 1
                    .Hogar = cRinkel
                    
            End Select
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(.Char.CharIndex, .Faccion.Status))
            
            If UserList(UserIndex).flags.Meditando Then
                If Estado <> ParticleToLevel(UserIndex) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(.Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(.Char.CharIndex, Estado, 0, True, True))
                End If
            End If
    
        End If

    End With
    
    Exit Sub

HandleRetirar_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRetirar", Erl)
     Resume Next
End Sub
Private Sub HandleRegresarHogar(ByVal UserIndex As Integer)

On Error GoTo HandleRegresarHogar_Err

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Muerto = 1 Then
        
            Call WarpUserChar(UserIndex, Ciudades(.Hogar).Dead_Map, Ciudades(.Hogar).Dead_X, Ciudades(.Hogar).Dead_Y, True)
            
        End If

    End With
    
    Exit Sub

HandleRegresarHogar_Err:
112     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRegresarHogar", Erl)
114     Resume Next
        
End Sub
 

Private Sub HandleParticulaUsuario(ByVal UserIndex As Integer)


    On Error GoTo ErrorHandler
    
     If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
   
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
 
    Dim UserName As String
    Dim tUser As Integer, Particula As Integer
       
    UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
    tUser = NameIndex(UserName)
    Particula = Buffer.ReadInteger()
    
    If PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Then
        
        If PersonajeExiste(UserName) Then
            
            If tUser <= 0 Then
                
                Call WriteLocaleMsg(UserIndex, 75)
            
            Else
            
                If UserList(tUser).Char.ParticulaFx > 0 Then
                
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(tUser).Char.CharIndex, UserList(tUser).Char.ParticulaFx, 0, True, True))
                Else
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(tUser).Char.CharIndex, Particula, -1, False, True))
                     Call WriteLocaleMsg(UserIndex, 461, UserList(tUser).Name)
                End If
            
            End If
            
        Else
        
            Call WriteLocaleMsg(UserIndex, 80)
            
        End If
        
    End If
            
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)

ErrorHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
154     Set Buffer = Nothing

156     If Error <> 0 Then Err.Raise Error

End Sub
 
Public Sub HandleProcesosLogin(ByVal UserIndex As Integer)


    If UserList(UserIndex).incomingData.length < 11 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
    On Error GoTo ErrHandler
    
    With UserList(UserIndex)
    
        Dim Buffer As New clsByteQueue
    
        Call Buffer.CopyBuffer(.incomingData)
            
        Dim UserCuenta As String
        Dim UserCode As String
        Dim UserContraseña As String
        Dim Version As String
        Dim PinHash As String
        Dim Salt As String

        Dim oSHA256 As CSHA256
        Set oSHA256 = New CSHA256
        Dim PasswordHash As String * 64
        
        Dim Tipo As Byte
        
        'Remove PacketID
        Call Buffer.ReadByte
            
        UserCuenta = Buffer.ReadASCIIString()
        UserCode = Buffer.ReadASCIIString()
        UserContraseña = Buffer.ReadASCIIString()
        Version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
        Tipo = Buffer.ReadByte()
        
114     If Not VersionOK(Version) Then
116         Call WriteShowMessageBox(UserIndex, 0, True, 7) 'Juego desactualizado
            Call FlushBuffer(UserIndex)
118         Call CloseSocket(UserIndex)
            Exit Sub
        End If

        If Not CuentaExiste(UserCuenta) Then
            Call WriteSendMsgBox(UserIndex, 39, 1) 'La cuenta no existe
            Call WriteEjecutarAccion(UserIndex, 5, "0")
        Else
 
            If LenB(SDesencriptar(UserContraseña)) = 0 Then
                Call WriteSendMsgBox(UserIndex, 49, 1) 'Ingrese una contraseña valida
                Call WriteEjecutarAccion(UserIndex, 5, "0")
            Else
            
                Salt = GetVar(AccountPath & UCase$(UserCuenta) & ".cnt", UserCuenta, "Salt")
                PasswordHash = oSHA256.SHA256(SDesencriptar(UserContraseña) & Salt)
                  
                If Tipo = 2 Then 'Borrar cuenta
                
                Call WriteEjecutarAccion(UserIndex, 5, "0")
                Else
                
                    If Tipo = 1 Then 'Cambiar contraseña
                        PinHash = GetVar(AccountPath & UserCuenta & ".cnt", UserCuenta, "Password")
                    Else 'Recuperar cuenta
                        PinHash = GetVar(AccountPath & UserCuenta & ".cnt", UserCuenta, "UserCodigo")
                    End If
                
                    
                    If PinValido(SDesencriptar(UserCode), PinHash, Salt) Then
                        Call WriteVar(AccountPath & UserCuenta & ".cnt", UserCuenta, "Password", PasswordHash)
                        Call WriteSendMsgBox(UserIndex, 50, 1) 'Contraseña restablecida
                        Call WriteEjecutarAccion(UserIndex, 5, "1")
                        
                    Else
                        If Tipo = 1 Then
                            Call WriteSendMsgBox(UserIndex, 52, 1) 'contraseña incorrecta.
                        Else
                            Call WriteSendMsgBox(UserIndex, 51, 1) 'Código incorrecto.
                        End If
                        Call WriteEjecutarAccion(UserIndex, 5, "0")
                    End If
                    
                End If

            End If
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
ErrHandler:

     Dim Error As Long

     Error = Err.Number

     On Error GoTo 0
    
    'Destroy auxiliar buffer
     Set Buffer = Nothing
     Set oSHA256 = Nothing
     
     If Error <> 0 Then Err.Raise Error

End Sub
 

 Public Sub HandleTransferGOLD(ByVal UserIndex As Integer)
'***************************************************
'Transfer gold to user
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Amount As Long
        Dim tIndex As Integer
        Dim GLDBANCO As Long
        Dim UserPath As String
                
        UserName = Buffer.ReadASCIIString()
        Amount = Buffer.ReadLong()
        
        tIndex = NameIndex(UserName)
        UserPath = CharPath & UserName & ".chr"
        If PersonajeExiste(UserName) Then
            If tIndex <> UserIndex Then
                If Amount > 0 And Amount <= MAXORO Then
                    If .Stats.Banco >= Amount Then
                        If tIndex <= 0 Then 'Personaje offline: se deposita en el banco
                            GLDBANCO = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                            
                            Call WriteVar(UserPath, "STATS", "BANCO", GLDBANCO + val(Amount))
                            .Stats.Banco = .Stats.Banco - Amount
                            Call WriteUpdateGold(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "¡Se ha realizado la transferencia correctamente! Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", FontTypeNames.FONTTYPE_INFO)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ORO2, .Pos.X, _
                                    .Pos.Y))
                        Else
                       
                        UserList(tIndex).Stats.Banco = UserList(tIndex).Stats.Banco + Amount
                           Call WriteUpdateGold(tIndex)
                           
                            .Stats.Banco = .Stats.Banco - Amount
                           Call WriteUpdateGold(UserIndex)
                           Call WriteConsoleMsg(UserIndex, "¡Se ha realizado la transferencia correctamente! Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", FontTypeNames.FONTTYPE_INFO)
                           Call WriteConsoleMsg(tIndex, "¡" & .Name & "ha transferido " & Amount & " monedas de oro a tu cuenta Goliath!.", FontTypeNames.FONTTYPE_INFO)
                           Call WriteConsoleMsg(tIndex, "¡Has recibido un nuevo mensaje de Finanzas Goliath, ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFO)  'corroe que no esta xD
                                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ORO2, .Pos.X, _
                                    .Pos.Y))
                        End If
                        
                    Else
                    Call WriteConsoleMsg(UserIndex, "No tienes suficientes monedas de oro depositadas en tu cuenta.", FontTypeNames.FONTTYPE_INFO)
                    End If
                 End If
                 Else
                 Call WriteConsoleMsg(UserIndex, "Objetivo inválido.", FontTypeNames.FONTTYPE_INFO)
            End If
            
        Else
            Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub
Public Sub WriteEfectoTerrenoParticula(ByVal UserIndex As Integer, ByVal ParticulaFx As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Time As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WriteEfectoTerrenoParticula" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageEfectoTerrenoParticula(ParticulaFx, X, Y, Time))
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
Public Function PrepareMessageEfectoTerrenoParticula(ByVal ParticulaFx As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Time As Long) As String
    
    On Error GoTo EfectoTerrenoParticula_Err
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.EfectoTerrenoParticula)
        Call .WriteInteger(ParticulaFx)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteLong(Time)
        
        PrepareMessageEfectoTerrenoParticula = .ReadASCIIStringFixed(.length)
    End With
        
    Exit Function

EfectoTerrenoParticula_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.EfectoTerrenoParticula", Erl)
     Resume Next
        
End Function
Public Sub WriteEfectoTerrenoFX(ByVal UserIndex As Integer, ByVal FX As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Loops As Integer)
 
 
On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageEfectoTerrenoFX(FX, X, Y, Loops))
    
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
Public Function PrepareMessageEfectoTerrenoFX(ByVal FX As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Loops As Integer) As String
 
    On Error GoTo PrepareMessageFXTerreno_Err
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.EfectoTerrenoFX)
        Call .WriteInteger(FX)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(Loops)
        
        PrepareMessageEfectoTerrenoFX = .ReadASCIIStringFixed(.length)
    End With
    
    Exit Function

PrepareMessageFXTerreno_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageFXTerreno", Erl)
     Resume Next
        
End Function
Public Sub WriteCorreoList(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    Dim i As Integer
    Dim slot As Byte
    
    slot = UserList(UserIndex).flags.CantidadCorreos
    
    With UserList(UserIndex).outgoingData
        
        Call .WriteByte(ServerPacketID.CorreoList)
        Call .WriteByte(slot)
        
        For i = 1 To (slot)
            Call .WriteASCIIString(UserList(UserIndex).Correos(i).Emisor)
            Call .WriteASCIIString(UserList(UserIndex).Correos(i).Carta)
            Call .WriteByte(UserList(UserIndex).Correos(i).Leida)
            Call .WriteInteger(UserList(UserIndex).Correos(i).ObjetoCantidad)
            Call .WriteInteger(UserList(UserIndex).Correos(i).ObjetoIndex)
        Next i
        
    End With
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub HandlePacketsCorreo(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
       
        Dim index As Byte
        Dim SlotCorreo As Byte
       
            index = .incomingData.ReadByte
            SlotCorreo = .incomingData.ReadByte
         Select Case index
       
            Case 1
                Call mod_Correos.ResetCorreos(UserIndex, SlotCorreo)
               
            Case 2
                If SlotCorreo <> 0 Then
                UserList(UserIndex).Correos(SlotCorreo).Leida = 1
                End If
            Case 3
                Call mod_Correos.RetirarItemCorreo(UserIndex, SlotCorreo)
            Case 4
            
        End Select
       
    End With
End Sub

Public Sub HandleEnviarCorreo(ByVal UserIndex As Integer)
'***************************************************
'Author: Shak
'Last Modification: 21/05/2015
'Enviamos el correo (Transferir objeto)
'***************************************************
    If UserList(UserIndex).incomingData.length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call Buffer.ReadByte

        Dim Destinatario As String
        Dim Mensaje As String
        Dim Objeto As Obj
        Dim tIndex As Integer
        Dim pjexiste As Boolean
        Destinatario = Buffer.ReadASCIIString
        Mensaje = Buffer.ReadASCIIString
        
        Objeto.ObjIndex = Buffer.ReadInteger
        Objeto.Amount = Buffer.ReadInteger

        pjexiste = FileExist(CharPath & Destinatario & ".chr")

         'Si no existe el personaje.
        If Not pjexiste Then
          Call WriteConsoleMsg(UserIndex, "No existe el personaje [" & Destinatario & "]", FontTypeNames.FONTTYPE_INFO)
       Else
       
        If Objeto.ObjIndex = 0 And Objeto.Amount = 0 Then 'es un msj de txto
                    If CantSendCorreo(UserIndex, Destinatario) Then
                      Call mod_Correos.EnviarCorreo(UserIndex, Destinatario, Mensaje, 0, 0, 0)
                    End If
        
        Else
                    If CantSendCorreo(UserIndex, Destinatario) Then
                      Call mod_Correos.EnviarCorreo(UserIndex, Destinatario, Mensaje, Objeto.ObjIndex, Objeto.Amount, 0)
                    End If
        End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
 
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set Buffer = Nothing
   
    If Error <> 0 Then _
        Err.Raise Error
End Sub
Public Sub HandleDonador(ByVal UserIndex As Integer)

        If UserList(UserIndex).incomingData.length < 3 Then
           Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
           Exit Sub
       End If
    
       On Error GoTo ErrHandler
       
        With UserList(UserIndex)
        
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        
        Dim UserName As String
        Dim tIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        tIndex = NameIndex(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
        
        If .flags.Privilegios And PlayerType.Dios Then
        
            If tIndex < 1 Then
                Call WriteLocaleMsg(UserIndex, 75) 'Usuario offline
            Else
        
                If UserList(tIndex).Donador.activo = 0 Then
                    Call WriteLocaleMsg(UserIndex, 478, UserName) 'Has hecho donador a la cuenta de #1.|12|1
                    Call WriteLocaleMsg(tIndex, 424) '¡Acabas de pasar al modo donador!
                    UserList(tIndex).Donador.activo = 1
                Else
                    UserList(tIndex).Donador.activo = 0
                    Call WriteLocaleMsg(UserIndex, 479, UserName) ' Has quitado el donador a la cuenta de #1.|12|1
                    Call WriteLocaleMsg(tIndex, 480) ' Ya no eres donador
                End If
            
                Call WriteVar(AccountPath & UserList(tIndex).Account & ".cnt", UserList(tIndex).Account, "Donador", .Donador.activo)
                
                Call RefreshCharStatus(tIndex)
            
                Call LogGM(UserList(UserIndex).Name, "Utilizó el comando /DONADOR con el usuario: " & UserName)
            End If
        
        End If
        
        
        End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error
End Sub
 
Public Sub handleEventoOro(ByVal UserIndex As Integer)
     If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
  Dim Temporal As Byte
  Dim tiempoevento As Byte
  Dim multiplicadorOroTemporal As Byte
Temporal = multiplicadorOro
With UserList(UserIndex)
Call .incomingData.ReadByte
 multiplicadorOroTemporal = .incomingData.ReadByte()
  tiempoevento = .incomingData.ReadByte()
If UserList(UserIndex).Name = "Shermie" Or UserList(UserIndex).Name = "Mermas" Then
multiplicadorOro = multiplicadorOroTemporal
Else
multiplicadorOro = 2

End If
 

If UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
Call LogGM(.Name, "El Game Master " & UserList(UserIndex).Name & "utilizó el comando /ORO " & multiplicadorOro)
   If OroModificada = False Then
      OroModificada = True
      Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg("¡El evento 'Oro x" & multiplicadorOro & "' ha dado comienzo! Duración: " & tiempoevento & " minutos.", FontTypeNames.FONTTYPE_INFO))
      'frmMain.EventoOro.Enabled = True
    '  MINUTOSSPAM(5) = tiempoevento
   Else
      OroModificada = False
      Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg("¡El evento 'Oro x" & Temporal & "' ha finalizado!", FontTypeNames.FONTTYPE_INFO))
   End If
   
Else

 Exit Sub
End If

End With

End Sub
Public Sub handleEventoExperiencia(ByVal UserIndex As Integer)
     If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
  Dim Temporal As Byte
  Dim GuardamosAca As Byte
  Dim tiempoevento As Byte
Temporal = multiplicadorExp
With UserList(UserIndex)
Call .incomingData.ReadByte
 GuardamosAca = .incomingData.ReadByte()
  tiempoevento = .incomingData.ReadByte()
If UserList(UserIndex).Name = "Shermie" Or UserList(UserIndex).Name = "Mermas" Then
multiplicadorExp = GuardamosAca
Else
multiplicadorExp = 2
End If

If UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then

Call LogGM(.Name, "El Game Master " & UserList(UserIndex).Name & " utilizó el comando /EXP " & multiplicadorExp)
  If ExpModificada = False Then
     ExpModificada = True
     Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg("¡El evento 'Experiencia x" & multiplicadorExp & "' ha dado comienzo! Duración: " & tiempoevento & " minutos.", FontTypeNames.FONTTYPE_INFO))
     'frmMain.EventoExp.Enabled = True
     EventosOroandExp = tiempoevento
       Else
     ExpModificada = False
     Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg("¡El evento 'Experiencia x" & Temporal & "' ha finalizado!", FontTypeNames.FONTTYPE_INFO))
  End If
  
Else

   Exit Sub
End If

End With
End Sub
Public Sub WriteCharStatus(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Status As Byte)

    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharStatus(CharIndex, Status))
   
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Function PrepareMessageCharStatus(ByVal CharIndex As Integer, ByVal priv As Byte) As String
    
    On Error GoTo PrepareMessageCharStatus_Err
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharStatus)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(priv)
        
        PrepareMessageCharStatus = .ReadASCIIStringFixed(.length)
    End With
    
    Exit Function

PrepareMessageCharStatus_Err:
     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharStatus", Erl)
     Resume Next
        
End Function
Public Sub HandleSeleccionarHogar(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
   Dim Hogar As Byte: Dim Equidad As Byte
   
    With UserList(UserIndex)
        Call .incomingData.ReadByte
                
    Select Case .incomingData.ReadByte()
    
    Case 0
    
            'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, _
                    "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", _
                    FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor) Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "Estás muy lejos del sacerdote.", _
                    FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
     Call WriteShowMessageBox(UserIndex, 0, True, 5)
      
    Case 1
    
        Select Case .Hogar
        Case 1: Equidad = 34
        Case 2: Equidad = 194
        Case 3: Equidad = 1
        Case 4: Equidad = 59
        Case 5: Equidad = 20
        Case 6: Equidad = 37
        Case 7: Equidad = 62
        Case 8: Equidad = 151
        Case 9: Equidad = 218
        Case 10: Equidad = 180
        Case 11: Equidad = 185
        Case 12: Equidad = 111
        End Select
        
        If .Pos.Map = Equidad Then
        Call WriteConsoleMsg(UserIndex, Trim$(MapInfo(UserList(UserIndex).Pos.Map).Name) & " es tu hogar.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
        End If
                
        Select Case .Pos.Map
        Case 20: UserList(UserIndex).Hogar = cRinkel
        Case 151: UserList(UserIndex).Hogar = cARGHAL
        Case 218: UserList(UserIndex).Hogar = cTIAMA
        Case 180: UserList(UserIndex).Hogar = cORAC
        Case 112: UserList(UserIndex).Hogar = cNueva
        Case Else
            If (esArmada(UserIndex)) Or (esCiuda(UserIndex)) = 1 Then
                Select Case .Pos.Map
                Case 1: UserList(UserIndex).Hogar = cUllathorpe
                Case 34: UserList(UserIndex).Hogar = cNix
                Case 59: UserList(UserIndex).Hogar = cBanderbill
                End Select
            ElseIf (esRepu(UserIndex) Or esMili(UserIndex)) = 1 Then
                Select Case .Pos.Map
                Case 194: UserList(UserIndex).Hogar = cIlliandor
                Case 63: UserList(UserIndex).Hogar = cLindos
                Case 184: UserList(UserIndex).Hogar = cSURAMEI
                End Select
            Else
            Call WriteConsoleMsg(UserIndex, "Ciudad inválida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            End If
        End Select
            Call WriteConsoleMsg(UserIndex, "Tu nuevo hogar ahora es " & Trim$(MapInfo(UserList(UserIndex).Pos.Map).Name) & ".", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
    End Select
    
    End With
End Sub
Public Sub HandleCasamiento(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
                
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String, Pareja As String
        Dim tUser As Integer, tUserPareja As Integer
        
        Dim Modo As Byte
        
        UserName = Replace(Buffer.ReadASCIIString, "+", " ")
        tUser = NameIndex(UserName)
        
        Modo = Buffer.ReadByte()

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
        
        If tUser <= 0 Then
            Call WriteLocaleMsg(UserIndex, 75) 'Usuario offline
        
        Else
            Call PuedeCasarse(UserIndex, tUser)
            
        End If
 
        
    End With

ErrHandler:
    Dim Error As Long
    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Public Sub handleDivorciar(ByVal UserIndex As Integer)

 
    With UserList(UserIndex)
    Call .incomingData.ReadByte


        Dim parejaactual As String
        Dim parejaindex As Integer
        parejaactual = UserList(UserIndex).Casamiento.Pareja
        parejaindex = NameIndex(parejaactual)
        
If .flags.Muerto = 1 Then
 Call WriteLocaleMsg(UserIndex, 77)
 Exit Sub
 
End If

    
    If UserList(UserIndex).Casamiento.Casado = 0 And UserList(UserIndex).Casamiento.Pareja = "" Then
      Call WriteConsoleMsg(UserIndex, "No puedes divorciarte porque no estás casado.", FONTTYPE_INFO)
       Exit Sub
       
    End If
 
 UserList(UserIndex).Casamiento.Casado = 0
UserList(UserIndex).Casamiento.Pareja = ""
 Call WriteConsoleMsg(UserIndex, parejaactual & " ya no es más tu pareja.", FontTypeNames.FONTTYPE_INFO)

UserList(parejaindex).Casamiento.Casado = 0
UserList(parejaindex).Casamiento.Pareja = ""
 Call WriteConsoleMsg(parejaindex, .Name & " ha decidido divorciarse de ti.", FontTypeNames.FONTTYPE_INFO)
.Casamiento.Candidato = 0
End With
 
End Sub
 
  

Public Sub HandleHayEventos(ByVal UserIndex As Integer)
With UserList(UserIndex)
Call .incomingData.ReadByte
 
 
If ExpModificada = True Then
' WriteConsoleMsg Userindex, "¡El evento 'Experiencia x" & multiplicadorExp & "' se encuentra en curso!" & " Duración: " & (EventosOroandExp(0) - EventosOroandExp(1)) & " minutos.", FontTypeNames.FONTTYPE_INFO
End If
If OroModificada = True Then
 ' WriteConsoleMsg Userindex, "¡El evento 'Oro x" & multiplicadorOro & "' se encuentra en curso!" & " Duración: " & (EventosOroandExp(2) - EventosOroandExp(3)) & " minutos.", FontTypeNames.FONTTYPE_INFO
End If
If OroModificada = False And ExpModificada = False Then
 'WriteConsoleMsg Userindex, "No hay eventos en curso.", FontTypeNames.FONTTYPE_INFO
End If


 
End With
End Sub

Public Sub WritePremios(ByVal UserIndex As Integer)

    If CantPremios = 0 Then Exit Sub
    
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Premios)
        Call .WriteInteger(CantPremios)
        Call .WriteLong(UserList(UserIndex).Donador.CreditoDonador)
        
    For i = 1 To CantPremios
    
            Call .WriteASCIIString(ObjData(PremiosInfo(i).ObjIndex).Name)
            Call .WriteInteger(PremiosInfo(i).puntos)
    Next i
    End With
        Exit Sub

ErrHandler:

102     If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
104         Resume
        End If

End Sub

Public Sub HandlePremiosRequest(ByVal UserIndex As Integer)
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePremios(UserIndex)
End Sub
 
 
 
Public Sub HandleRPremios(ByVal UserIndex As Integer)
 
Dim Premio As Obj
Dim index As Integer
 
 
With UserList(UserIndex).incomingData
    .ReadByte
    index = .ReadInteger
    
End With
 
'Set the object
Premio.ObjIndex = PremiosInfo(index).ObjIndex
Premio.Amount = 1
 
If Premio.ObjIndex <= 0 Then Exit Sub
 
 
If PremiosInfo(index).puntos <= UserList(UserIndex).Donador.CreditoDonador Then
  If Not MeterItemEnInventario(UserIndex, Premio) Then
  Call WriteLocaleMsg(UserIndex, 328)
  Exit Sub
  Else
  UserList(UserIndex).Donador.CreditoDonador = UserList(UserIndex).Donador.CreditoDonador - PremiosInfo(index).puntos
  Call UpdateUserInv(True, UserIndex, 0)
  End If
Else
    Call WriteLocaleMsg(UserIndex, 423)
End If
 
End Sub
 
Private Sub HandleDARPUN(ByVal UserIndex As Integer)
 
If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call Buffer.ReadByte
       
        Dim UserName As String
        Dim DARPUN As Long
        Dim tUser As Integer
       
        UserName = Buffer.ReadASCIIString()
        DARPUN = Buffer.ReadLong()
       
        If InStr(1, UserName, " ") Then
            UserName = Replace(UserName, " ", " ")
        End If
       
       If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
       
        If LenB(UserName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Comando Incorrecto. Utilice /darpun NICKNAME@PUNTOS.", FontTypeNames.FONTTYPE_INFO)
        Else
        tUser = NameIndex(UserName)
        End If
       
        If tUser <= 0 Then
         Call WriteLocaleMsg(UserIndex, 75)
        End If
       
        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
        UserList(tUser).Donador.CreditoDonador = UserList(tUser).Donador.CreditoDonador + DARPUN
        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te otorgó " & DARPUN & " créditos.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Has enviado " & DARPUN & " créditos a " & UserList(tUser).Name, FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg("Servidor> " & UserList(UserIndex).Name & " otorgó " & DARPUN & " créditos a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_TALK))

        Else
        Call WriteLocaleMsg(UserIndex, 80)
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
   
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set Buffer = Nothing
   
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub WriteMensajeSigno(ByVal UserIndex As Integer, ByVal Recibio As Byte)

    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MensajeSigno)
    Call UserList(UserIndex).outgoingData.WriteByte(Recibio)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub
 
Public Sub HandleCloseGuild(ByVal UserIndex As Integer)
With UserList(UserIndex)
Call .incomingData.ReadByte
 
Dim NombreClan As String
Dim LiderClan As String
Dim CantidadClanes As Byte
NombreClan = modGuilds.GuildName(.GuildIndex)
LiderClan = modGuilds.GuildLeader(.GuildIndex)
 
 
If .flags.Muerto = 1 Then
 Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If

If MapInfo(.Pos.Map).Pk = True Then
  Call WriteConsoleMsg(UserIndex, "¡No puedes cerrar el clan en zona insegura!", FontTypeNames.FONTTYPE_INFO)
 Exit Sub
End If
 
 

If .GuildIndex = 0 Then
WriteConsoleMsg UserIndex, "¡No perteneces a ningún clan!", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If
 
If LiderClan <> .Name Then
WriteConsoleMsg UserIndex, "¡No eres el lider del clan!", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If
 
If GetVar(App.Path & "\guilds\" & NombreClan & "-members.mem", "INIT", "NroMembers") > 1 Then
WriteConsoleMsg UserIndex, "¡Debes hechar a todos los miembros del clan para cerrarlo!", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If
 
Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha cerrado el clan llamado: " & NombreClan, FontTypeNames.FONTTYPE_INFO))
 
'Call Kill(App.Path & "\guilds\" & NombreClan & "-members.mem")
'Call Kill(App.Path & "\guilds\" & NombreClan & "-solicitudes.sol")

Call Kill(App.Path & "\guilds\" & NombreClan & "-members.mem")
Call Kill(App.Path & "\guilds\" & NombreClan & "-solicitudes.sol")
  
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Founder", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "GuildName", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Date", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Antifaccion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Alineacion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex1", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex2", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex3", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex4", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex5", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex6", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex7", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex8", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Desc", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "GuildNews", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Leader", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "URL", vbNullString)
 
Call WriteVar(CharPath & LiderClan & ".chr", "GUILD", "GUILDINDEX", vbNullString)
Call WriteVar(CharPath & LiderClan & ".chr", "GUILD", "AspiranteA", vbNullString)
Call WriteVar(CharPath & LiderClan & ".chr", "GUILD", "Miembro", vbNullString)

'Call modGuilds.m_EcharMiembroDeClan(-1, UserList(userindex).name)
'Call modGuilds.m_DesconectarMiembroDelClan(.name, .GuildIndex)
 

.GuildIndex = 0

  'Upda

WriteConsoleMsg UserIndex, "¡Clan eliminado!", FontTypeNames.FONTTYPE_INFO
  'Update tag
 Call RefreshCharStatus(UserIndex)
' Call WarpUserChar(userindex, .Pos.Map, .Pos.x, .Pos.Y, False)
 
End With
End Sub
 Private Sub HandleCuentaRegresiva(ByVal UserIndex As Integer)
 
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call Buffer.ReadByte
       
        Dim Seconds As Byte
         Dim lugar As Byte
        Seconds = Buffer.ReadByte()
        lugar = Buffer.ReadByte()
 
        
        CuentaRegresivaTimer = Seconds
        mapasegundos = lugar
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
   
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set Buffer = Nothing
   
    If Error <> 0 Then _
        Err.Raise Error
End Sub
Public Sub WriteMarcamosSkin(ByVal UserIndex As Integer, ByVal index As Byte)

        On Error GoTo ErrHandler
        With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MarcamosSkin)
        Call .WriteByte(index)
        End With
        Exit Sub

ErrHandler:

        If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
            Resume
        End If

End Sub

Public Sub HandleMsgAmigo(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
  If UserList(UserIndex).incomingData.length < 3 Then
  Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
  Exit Sub
  End If

  On Error GoTo ErrHandler
  With UserList(UserIndex)
  'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
  Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
  Call Buffer.CopyBuffer(.incomingData)

  'Remove packet ID
  Call Buffer.ReadByte

  Dim Mensaje As String
  Dim i As Long

  Mensaje = Buffer.ReadASCIIString()
    
  For i = 1 To .flags.CantidadAmigos
  If .Amigos(i).index > 0 Then _
  Call WriteConsoleMsg(.Amigos(i).index, "[" & .Name & "] " & Mensaje, FontTypeNames.FONTTYPE_INFOBOLD4)
  Next i
  Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & Mensaje, FontTypeNames.FONTTYPE_INFOBOLD4)


  'If we got here then packet is complete, copy data back to original queue
  Call .incomingData.CopyBuffer(Buffer)
  End With

ErrHandler:
  Dim Error As Long
  Error = Err.Number
  On Error GoTo 0

  'Destroy auxiliar buffer
  Set Buffer = Nothing

  If Error <> 0 Then _
  Err.Raise Error
End Sub

Public Sub HandleOnAmigo(ByVal UserIndex As Integer)

    With UserList(UserIndex)
  'Remove packet ID
  Call .incomingData.ReadByte
 
  Dim list As String
  Dim i As Long
  Dim tUser As Long
  Dim tUser2 As Long
  Dim Cantidad(5) As String
 
  
  For i = 1 To .flags.CantidadAmigos
 
   Cantidad(i) = UserList(UserIndex).Amigos(i).Nombre
   tUser2 = NameIndex(Cantidad(i))
    
        If tUser2 <= 0 Then
          If i = .flags.CantidadAmigos Then
          list = list & Cantidad(i) & "(Offline)" & "."
          Else
          list = list & Cantidad(i) & "(Offline)" & ","
          End If
        Else
        If i = .flags.CantidadAmigos Then
        list = list & Cantidad(i) & "(Online)(" & Trim$(MapInfo(UserList(tUser2).Pos.Map).Name) & ")."
        Else
        list = list & Cantidad(i) & "(Online)(" & Trim$(MapInfo(UserList(tUser2).Pos.Map).Name) & "), "
        End If
        End If
     Next i
  
  If LenB(list) > 0 Then
  Call WriteConsoleMsg(UserIndex, "Amigos conectados: " & list, FontTypeNames.FONTTYPE_INFO)
  Else
  Call WriteConsoleMsg(UserIndex, "Tu lista de amigos está vacía.", FontTypeNames.FONTTYPE_INFO)
  End If
 
  End With
 
End Sub
Public Sub HandleAddAmigo(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
  If UserList(UserIndex).incomingData.length < 4 Then
  Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
  Exit Sub
  End If

  On Error GoTo ErrHandler
  With UserList(UserIndex)
  'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
  Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
  Call Buffer.CopyBuffer(.incomingData)

  'Remove packet ID
  Call Buffer.ReadByte

  Dim UserName As String
  Dim caso As Byte
  Dim razon As String
  Dim tUser As Integer
  Dim slot As Byte

  UserName = Buffer.ReadASCIIString()
  caso = Buffer.ReadByte
  tUser = NameIndex(UserName)

   
  Select Case caso
  
  Case 1 'Mandar solicitudad de amistad
  If IntentarAgregarAmigo(UserIndex, tUser, razon) = True Then
  Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " fue agregado a tu lista de amigos, espere confirmación.", FontTypeNames.FONTTYPE_INFO)
  Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te agregó a tu lista de amigos, para aceptarlo ingrese /FACCEPT " & .Name & ".", FontTypeNames.FONTTYPE_INFO)
  UserList(tUser).QuienAmigo = .Name
  Else
  Call WriteConsoleMsg(UserIndex, razon, FontTypeNames.FONTTYPE_INFO)
  End If
  
  Case 2 'Confirmar solicitud de amistad
  If IntentarAgregarAmigo(UserIndex, tUser, razon) = True Then
  If LenB(.QuienAmigo) >= 3 Then
  If UCase$(.QuienAmigo) = UCase$(UserList(tUser).Name) Then
  slot = BuscarSlotAmigoVacio(UserIndex)
  .Amigos(slot).Nombre = UserList(tUser).Name
  
  slot = BuscarSlotAmigoVacio(tUser)
  UserList(tUser).Amigos(slot).Nombre = .Name
  
  Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " está jugando en Mohurall (Argentina).", FontTypeNames.FONTTYPE_INFO)
  Call WriteConsoleMsg(tUser, .Name & " está jugando en Mohurall (Argentina).", FontTypeNames.FONTTYPE_INFO)
  UserList(tUser).flags.CantidadAmigos = UserList(tUser).flags.CantidadAmigos + 1
  UserList(UserIndex).flags.CantidadAmigos = UserList(UserIndex).flags.CantidadAmigos + 1
  If UserList(UserIndex).flags.CantidadAmigos = 1 Then UserList(UserIndex).flags.CheckAmigos = 1
  If UserList(tUser).flags.CantidadAmigos = 1 Then UserList(tUser).flags.CheckAmigos = 1
   slot = ObtenerIndexLibre(UserIndex)
  If slot > 0 Then .Amigos(slot).index = tUser
  slot = ObtenerIndexLibre(tUser)
  If slot > 0 Then UserList(tUser).Amigos(slot).index = UserIndex
  .QuienAmigo = vbNullString
  Else
  Call WriteConsoleMsg(UserIndex, "Acción invalida", FontTypeNames.FONTTYPE_INFO)
  End If
  End If
  Else
  Call WriteConsoleMsg(UserIndex, razon, FontTypeNames.FONTTYPE_INFO)
  End If
  End Select
  'If we got here then packet is complete, copy data back to original queue
  Call .incomingData.CopyBuffer(Buffer)
  End With

ErrHandler:
  Dim Error As Long
  Error = Err.Number
  On Error GoTo 0

  'Destroy auxiliar buffer
  Set Buffer = Nothing

  If Error <> 0 Then _
  Err.Raise Error
End Sub


Private Sub HandleDelAmigo(ByVal UserIndex As Integer)
'***************************************************
'Author:Bateman
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
  Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
  Exit Sub
  End If

  With UserList(UserIndex)
  'Remove packet ID
  Call .incomingData.ReadByte
  Dim nick As String
  Dim nick2 As String
  Dim i As Long
  Dim Cantidad As Byte
  Dim slot As Byte
  Dim tUser As Integer
  Dim UserName As String
  Dim Looper As Byte
  
      nick = .incomingData.ReadASCIIString()
      
      If Not PersonajeExiste(nick) Then Exit Sub
      
      For i = 1 To UserList(UserIndex).flags.CantidadAmigos
      nick2 = UserList(UserIndex).Amigos(i).Nombre
      If UCase$(nick) = UCase$(nick2) Then
      slot = i
      Exit For
      End If
      Next i
      
  If slot <= 0 Or slot > UserList(UserIndex).flags.CantidadAmigos Then Exit Sub
  If .Amigos(slot).Nombre = "Vacío" Then Exit Sub


  tUser = NameIndex(.Amigos(slot).Nombre)
  UserName = .Amigos(slot).Nombre
    Call WriteConsoleMsg(UserIndex, .Amigos(slot).Nombre & " fue quitado de tu lista de amigos.", FontTypeNames.FONTTYPE_INFO)
   'reseteamos el slot
    Cantidad = UserList(UserIndex).flags.CantidadAmigos
    If slot = Cantidad Then
   .Amigos(slot).Nombre = "Vacío"
   .QuienAmigo = 0
    Else
    For Looper = slot To Cantidad
    If Looper = Cantidad Then Exit For
    UserList(UserIndex).Amigos(Looper).Nombre = UserList(UserIndex).Amigos(Looper + 1).Nombre
    UserList(UserIndex).Amigos(Looper + 1).Nombre = "Vacío"
    
    UserList(UserIndex).Amigos(Looper).index = UserList(UserIndex).Amigos(Looper + 1).index
    UserList(UserIndex).Amigos(Looper + 1).index = 0
    Next Looper
    .QuienAmigo = 0
    End If
    
    UserList(UserIndex).flags.CantidadAmigos = UserList(UserIndex).flags.CantidadAmigos - 1
    If .flags.CantidadAmigos = 0 Then .flags.CheckAmigos = 0
  If tUser > 0 Then
  'Puede pasar....
  If BuscarSlotAmigoName(tUser, .Name) Then
  Call WriteConsoleMsg(tUser, .Name & " te ha quitado de su lista de amigos.", FontTypeNames.FONTTYPE_INFO)
  slot = BuscarSlotAmigoNameSlot(tUser, .Name)
  Cantidad = UserList(tUser).flags.CantidadAmigos
  UserList(tUser).flags.CantidadAmigos = UserList(tUser).flags.CantidadAmigos - 1
  If UserList(tUser).flags.CantidadAmigos = 0 Then UserList(tUser).flags.CheckAmigos = 0
    If slot = Cantidad Then
    UserList(tUser).Amigos(slot).Nombre = "Vacío"
    UserList(tUser).QuienAmigo = 0
    slot = ObtenerIndexUsuado(UserIndex, tUser)
   ' If slot > 0 Then .Amigos(slot).index = 0
    slot = ObtenerIndexUsuado(tUser, UserIndex)
   ' If slot > 0 Then UserList(tuser).Amigos(slot).index = 0
    Else
    For Looper = slot To Cantidad
    If Looper = Cantidad Then Exit For
    UserList(tUser).Amigos(Looper).Nombre = UserList(tUser).Amigos(Looper + 1).Nombre
    UserList(tUser).Amigos(Looper + 1).Nombre = "Vacío"
    
    UserList(tUser).Amigos(Looper).index = UserList(tUser).Amigos(Looper + 1).index
    UserList(tUser).Amigos(Looper + 1).index = 0
    Next Looper
    End If
    End If
    
  Else
  'verificamos desde el char
  Call delAmigoOfli(UserName, .Name)
  End If

  End With
End Sub
Public Sub Writemostrarubicacion(ByVal UserIndex As Integer, ByVal CharIndex As String, ByVal NumAmigo As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

        On Error GoTo ErrHandler
        With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MostrarUbicacion)
        Call .WriteASCIIString(CharIndex)
        Call .WriteByte(NumAmigo)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        End With
        Exit Sub

ErrHandler:

        If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
            Resume
        End If

End Sub
Public Sub WriteCargarSkin(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    Dim Head As Integer, body As Integer, casco As Byte, weapon As Byte, shield As Byte
    
    Head = UserList(UserIndex).Char.Head
    body = UserList(UserIndex).Char.body
    casco = UserList(UserIndex).Char.CascoAnim
    weapon = UserList(UserIndex).Char.WeaponAnim
    shield = UserList(UserIndex).Char.ShieldAnim
     
    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.AddPJ)
        Call .WriteInteger(Head)
        Call .WriteInteger(body)
        Call .WriteByte(casco)
        Call .WriteByte(weapon)
        Call .WriteByte(shield)
            
    End With
     
ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
Public Sub WriteCharMsgStatus(ByVal UserIndex As Integer, ByVal tI As Integer)

    On Error GoTo ErrHandler

    Dim St1 As Integer, St2 As Byte
    
    With UserList(UserIndex).outgoingData
         
        Call .WriteByte(ServerPacketID.CharMsgStatus)
    
        Call .WriteInteger(UserList(tI).Char.CharIndex)
           
        Select Case UserList(tI).flags.Privilegios
            
            Case PlayerType.Consejero
                Call .WriteByte(10)
            Case PlayerType.RoleMaster
                Call .WriteByte(11)
            Case PlayerType.SemiDios
                Call .WriteByte(12)
            Case PlayerType.Dios
                Call .WriteByte(13)
            
            Case Else
                
                Select Case UserList(tI).Faccion.Status
                    Case 1
                        Call .WriteByte(1)
                    Case 2
                        Call .WriteByte(2)
                    Case 3
                        Call .WriteByte(3)
                    Case 4
                        Call .WriteByte(5)
                    Case 5
                        Call .WriteByte(6)
                    Case 6
                        Call .WriteByte(7)
                    Case Else
                        Call .WriteByte(1)
                End Select
                
       End Select
        
       Call .WriteLong(CLng(((UserList(tI).Stats.MinHP / 100) / (UserList(tI).Stats.MaxHP / 100)) * 100))
     
       St1 = Generate_Char_Stat(tI)
       St2 = Generate_Char_StatEx(tI)
    
       Call .WriteInteger(St1)
       Call .WriteByte(St2)
       
    
       Call .WriteByte(UserList(tI).Clase)
       
       If UserList(UserIndex).Stats.ELV >= UserList(tI).Stats.ELV + 5 Or UserList(UserIndex).Stats.ELV >= UserList(tI).Stats.ELV Or _
           UserList(UserIndex).Stats.ELV + 5 >= UserList(tI).Stats.ELV Then
           Call .WriteInteger(UserList(tI).Stats.ELV)
       Else
           Call .WriteInteger(255)
       End If
       
       Call .WriteByte(UserList(tI).raza)
       
       Call .WriteByte(UserList(tI).Donador.activo)
       
       Call .WriteByte(UserList(tI).Faccion.Rango)
       
       Call .WriteASCIIString(UserList(tI).Casamiento.Pareja)
       Call .WriteASCIIString(UserList(tI).desc)
       
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub






''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteMsg(userindex)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub WriteRestOK(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Sub WriteCharMsgStatusNPC(ByVal UserIndex As Integer, ByVal tI As Integer)

 
    On Error GoTo ErrHandler

    Dim St As Byte
        
    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.CharMsgStatusNPC)
        Call .WriteInteger(Npclist(tI).Numero)

        Select Case Npclist(tI).flags.Status
        
            Case 0
                Call .WriteByte(0)
            Case 1
                Call .WriteByte(1)
            Case 2
                Call .WriteByte(2)
            Case 3
                Call .WriteByte(3)
            Case 4
                Call .WriteByte(4)
            Case Else
                Call .WriteByte(1)
                
        End Select
        
        If EsGm(UserIndex) Or UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
            Call .WriteByte(1)
            Call .WriteLong(CLng(Npclist(tI).Stats.MinHP))
        Else
            Call .WriteByte(0)
            Call .WriteLong(CLng(((Npclist(tI).Stats.MinHP / 100) / (Npclist(tI).Stats.MaxHP / 100)) * 100))
        End If
        
        St = Generate_Char_StatExNpcs(tI)
        Call .WriteByte(St)
        
        If UserList(UserIndex).Stats.ELV >= Npclist(tI).Leveles + 10 Or UserList(UserIndex).Stats.ELV >= Npclist(tI).Leveles Or UserList(UserIndex).Stats.ELV + 10 >= Npclist(tI).Leveles Then
            Call .WriteInteger(Npclist(tI).Leveles)
        Else
            Call .WriteInteger(255)
        End If
        
        If tI = CentinelaNPCIndex Then
            Call modCentinela.CentinelaSendClave(UserIndex)
        End If
        
        If Npclist(tI).MaestroUser > 0 Then
            Call .WriteByte(1)
        Else
            Call .WriteByte(0)
        End If
        
        If Npclist(tI).Owner > 0 Then
            Call .WriteByte(1)
        Else
            Call .WriteByte(0)
        End If
        
    End With
    
    Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub


 
Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, ByVal id As Integer, Optional ByVal strExtra As String = vbNullString, Optional ByVal Modo As Byte = 0, Optional ByVal fuente As Byte = 0)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "WriteLocaleMsg" message to the given user's outgoing data buffer
        '***************************************************
        On Error GoTo ErrHandler
        
        If UserIndex = 0 Then Exit Sub
        
100     Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageLocaleMsg(id, strExtra, Modo, fuente))

        Exit Sub

ErrHandler:

102     If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
104         Resume
        End If

End Sub

Public Function PrepareMessageLocaleMsg(ByVal id As Integer, Optional ByVal chat As String = vbNullString, Optional ByVal Modo As Byte = 0, Optional ByVal fuente As Byte = 0) As String
        
        On Error GoTo PrepareMessageLocaleMsg_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "PrepareMessageLocaleMsg" message and returns it.
        '***************************************************
        
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.LocaleMsg)
104         Call .WriteInteger(id)
106         Call .WriteASCIIString(chat)
108         Call .WriteByte(Modo)
109         Call .WriteByte(fuente)

110         PrepareMessageLocaleMsg = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageLocaleMsg_Err:
112     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageLocaleMsg", Erl)
114     Resume Next
        
End Function

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedSuccessful(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Logged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.LoggedSuccessful)
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub WriteAbrirFormularios(ByVal UserIndex As Integer, ByVal Formulario As Byte)


    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.AbrirFormularios)
    Call UserList(UserIndex).outgoingData.WriteByte(Formulario)
    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Private Sub HandleAbrirForms(ByVal UserIndex As Integer)
    
    On Error GoTo ErrorHandler
    
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        
        Dim Formulario As Integer
                
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Formulario = .incomingData.ReadByte
 
    
        Select Case Formulario
        
        Case 1
        
            If .Donador.activo = 1 Then
                Call WriteAbrirFormularios(UserIndex, 10)
            Else
                Call WriteAbrirFormularios(UserIndex, 9)
            End If

        End Select

    End With

    Exit Sub

ErrorHandler:
     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleAbrirForms", Erl)
     Resume Next
End Sub


Public Sub writeChangeInventorySlotUser(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Accion As Byte, ByVal Valor As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlotUser)
        Call .WriteByte(slot)
        Call .WriteByte(Accion)
        Call .WriteInteger(Valor)
    End With
    
     Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Function PrepareMessageAuraToChar(ByVal CharIndex As Integer, ByVal Aura As Byte, ByVal Tipo As Byte) As String
        
        On Error GoTo PrepareMessageAuraToChar_Err
        
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.AuraToChar)
104         Call .WriteInteger(CharIndex)
106         Call .WriteByte(Aura)
110         Call .WriteByte(Tipo)
112         PrepareMessageAuraToChar = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageAuraToChar_Err:
114     Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageAuraToChar", Erl)
116     Resume Next
        
End Function

Public Sub WriteUpdateSed(ByVal UserIndex As Integer)


    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSed)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteUpdateHambre(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHambre)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
 

Private Sub HandleDesconectarCuenta(ByVal UserIndex As Integer)

100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
    
        On Error GoTo ErrHandler

104     With UserList(UserIndex)

            'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
            Dim Buffer As New clsByteQueue
106         Call Buffer.CopyBuffer(.incomingData)
        
            'Remove packet ID
108         Call Buffer.ReadByte
        
            Dim Personaje As String, Cuenta As String, Nombre As String
            
            Dim PersonajeIndex As Integer
            
112         Personaje = Trim$(Buffer.ReadASCIIString())

            If InStr(1, Personaje, "+") Then
                Personaje = Replace(Personaje, "+", " ")
            End If
        
             
            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
 
                 PersonajeIndex = NameIndex(Personaje)
                 
                 If PersonajeIndex > 0 Then 'existe el usuario destino?
                    
                    Cuenta = UserList(PersonajeIndex).Account
                    Nombre = UserList(PersonajeIndex).Name
                    
                    If UserList(UserIndex).Name <> Nombre Then
                        Call WriteVar(AccountPath & Cuenta & ".cnt", Cuenta, "Conectada", "0")
                        Call FlushBuffer(PersonajeIndex)
                        Call CloseSocket(PersonajeIndex)
                        
                        Call WriteConsoleMsg(UserIndex, "Desconectaste a " & Nombre, FontTypeNames.FONTTYPE_INFO)
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes desconectarte a ti mismo.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                
                    Call WriteConsoleMsg(UserIndex, "Usuario offline o inexistente.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
            End If
            
 
        Call .incomingData.CopyBuffer(Buffer)
        End With
        
ErrHandler:

        Dim Error As Long

148     Error = Err.Number

        On Error GoTo 0
    
        'Destroy auxiliar buffer
150     Set Buffer = Nothing
    
152     If Error <> 0 Then Err.Raise Error

End Sub

Public Sub WriteEjecutarAccion(ByVal UserIndex As Integer, ByVal Accion As Byte, Optional ByVal Extra As String = "")

On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.EjecutarAccion)
    Call UserList(UserIndex).outgoingData.WriteByte(Accion)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(Extra)
    
Exit Sub

ErrHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
