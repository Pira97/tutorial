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
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

Public InBytes            As Long
Public OutBytes           As Long

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Enum Stat
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

Private Enum StatEx
    Paralizado = &H1
    Inmovilizado = &H2
    Hombre = &H4
    Mujer = &H8
End Enum

Private Enum ServerPacketID
    LoggedSuccessful        ' 0
    Logged                  ' 1
    RemoveDialogs           ' 2
    RemoveCharDialog        ' 3
    NavigateToggle          ' 4
    MontateToggle           ' 5
    Disconnect              ' 6
    CommerceEnd             ' 7
    BankEnd                 ' 8
    CommerceInit            ' 9
    BankInit                ' 10
    UpdateSta               ' 11
    UpdateMana              ' 12
    UpdateHP                ' 13
    UpdateGold              ' 14
    UpdateExp               ' 16
    ChangeMap               ' 17
    PosUpdate               ' 18
    ChatOverHead            ' 19
    ChatOverHeadLocale      ' 20
    ConsoleMsg              ' 21
    GuildChat               ' 22
    ShowMessageBox          ' 23
    UserIndexInServer       ' 24
    UserCharIndexInServer   ' 25
    CharacterCreate         ' 26
    CharacterRemove         ' 27
    CharacterMove           ' 29
    ForceCharMove           ' 30
    CharacterChange         ' 31
    CharacterChangeSlot     ' 32
    ObjectCreate            ' 33
    ObjectDelete            ' 34
    BlockPosition           ' 35
    PlayMIDI                ' 36
    PlayWave                ' 37
    guildList               ' 38
    AreaChanged             ' 39
    PauseToggle             ' 40
    RainToggle              ' 41
    CreateFX                ' 42
    UpdateUserStats         ' 43
    UpdateUserStatsForLevel ' 44
    WorkRequestTarget       ' 45
    ChangeInventorySlot     ' 46
    ChangeBankSlot          ' 47
    ChangeSpellSlot         ' 48
    atributes               ' 49
    BlacksmithWeapons       ' 50
    BlacksmithArmors        ' 51
    BlacksmithHelmet        ' 52
    BlacksmithShield        ' 53
    CarpenterObjects        ' 54
    SastreObjects           ' 55
    AlquimiaObjects         ' 56
    RestOK                  ' 57
    SendMsgBox              ' 58
    Blind                   ' 59
    Dumb                    ' 60
    ChangeNPCInventorySlot  ' 61
    UpdateHungerAndThirst   ' 62
    MiniStats               ' 63
    LevelUp                 ' 64
    SetInvisible            ' 65
    MeditateToggle          ' 66
    BlindNoMore             ' 67
    DumbNoMore              ' 68
    SendSkills              ' 69
    TrainerCreatureList     ' 70
    guildNews               ' 71
    OfferDetails            ' 72
    AlianceProposalsList    ' 73
    PeaceProposalsList      ' 74
    CharacterInfo           ' 75
    GuildLeaderInfo         ' 76
    GuildMemberInfo         ' 77
    GuildDetails            ' 78
    ParalizeOK              ' 79
    ShowUserRequest         ' 80
    TradeOK                 ' 81
    BankOK                  ' 82
    Pong                    ' 84
    UpdateTagAndStatus      ' 85
    LocaleMsg               ' 86
    
    'GM messages
    ShowSOSForm             ' 88
    UserNameList            ' 89
    correolist              ' 90
    UpdateStrenght          ' 91
    UpdateDexterity         ' 92
    Premios                 ' 95
    EfectoCharParticula     ' 96
    AddPJ                   ' 97
    EfectoTerrenoParticula  ' 99
    EfectoTerrenoFX         ' 100
    CharStatus              ' 101
    MensajeSigno            ' 102
    MarcamosSkin            ' 106
    MostrarUbicacion        ' 107
    CargarSkin              ' 108
    CharMsgStatus           ' 109
    CharMsgStatusNPC        ' 110
    AbrirFormularios        ' 111
    ChangeInventorySlotUser ' 112
    AuraToChar              ' 113
    UpdateSed               ' 114
    UpdateHambre            ' 115
    EjecutarAccion          ' 116
End Enum
 
Private Enum ClientPacketID
    Walk                    '5
    LoginExistingChar       '0
    LoginNewChar            '1
    Talk                    '3
    Whisper                 '4
    RequestPositionUpdate   '6
    Attack                  '7
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
    RegresarHogar           '110
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
    TeleportCreate = 20 ' Asignamos un número único al paquete de teletransportación
End Enum


Public Sub Connect(ByVal Modo As E_MODO)
    '*********************************************************************
    'Author: Jopi
    'Conexion al servidor mediante la API de Windows.
    '*********************************************************************
    Debug.Print "Conectando en Sub Connect"
    

    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    EstadoLogin = Modo
    frmMain.Socket1.HostName = CurServerIP
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
End Sub
Public Sub ReConnect(ByVal Modo As E_MODO)

    Debug.Print "Conectando en Sub Re-Connect"
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    EstadoLogin = Modo
    frmMain.Socket1.HostName = CurServerIP
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error Resume Next

    Dim Paquete As Long

    Paquete = CLng(incomingData.PeekByte())
    
    
    #If Debugger = 1 Then
    Call AddtoRichTextBox("El servidor mandó el paquete Nº " & Paquete & " (" & incomingData.Length & " Bytes)" & " - " & Time, 0, 0, 0, 0, 0, 0, 17)
    #End If
    
    Select Case Paquete
         
        Case ServerPacketID.LoggedSuccessful
            Call HandleLoggedSuccessful
            
        Case ServerPacketID.Logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
        
        Case ServerPacketID.MontateToggle           ' MONT
            Call HandleMontateToggle
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd

        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit

        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold

        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
                
        Case ServerPacketID.ChatOverHeadLocale            ' ||
            Call HandleChatOverHeadLocale
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.CharacterChangeSlot
            Call HandleCharacterChangeSlot
          
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              ' LLU
            Call HandleRainToggle
            
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.UpdateUserStatsForLevel
            Call HandleUpdateUserStatsForLevel
            
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       ' LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        ' LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.BlacksmithHelmet
            Call HandleBlacksmithHelmet
        
        Case ServerPacketID.BlacksmithShield
            Call HandleBlacksmithShield
        
        Case ServerPacketID.CarpenterObjects        ' OBR
            Call HandleCarpenterObjects
        
        Case ServerPacketID.SastreObjects
            Call HandleSastreObjects
            
        Case ServerPacketID.AlquimiaObjects
            Call HandleAlquimiaObjects
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK

        Case ServerPacketID.SendMsgBox                ' ERR
            Call HandleSendMsgBox
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
            
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
            
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
   
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible

        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews               ' GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      ' PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo

        Case ServerPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.TradeOK                 ' TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  ' BANCOOK
            Call HandleBankOK
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
 
        Case ServerPacketID.LocaleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleLocaleMsg
            
        'GM messages
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm

        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList

        Case ServerPacketID.correolist
            Call HandleCorreoList
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity

        Case ServerPacketID.Premios                 ' CANJE
            Call HandlePremios
 
        Case ServerPacketID.EfectoCharParticula      ' CPC
            Call HandleEfectoCharParticula
         
        Case ServerPacketID.AddPJ
            Call HandleAddPj
            
        Case ServerPacketID.EfectoTerrenoParticula            ' HO
            Call HandleEfectoTerrenoParticula

        Case ServerPacketID.EfectoTerrenoFX           ' HO
            Call HandleEfectoTerrenoFX
            
        Case ServerPacketID.CharStatus
            Call HandleCharStatus
 
        Case ServerPacketID.MensajeSigno
            Call HandleMensajeSigno
 
        Case ServerPacketID.MarcamosSkin
           Call HandleMarcamosSkin
           
        Case ServerPacketID.MostrarUbicacion
           Call Handlemostrarubicacion
           
        Case ServerPacketID.CargarSkin
            Call HandleCargarSkin
        
        Case ServerPacketID.CharMsgStatus
            Call HandleCharMsgStatus

        Case ServerPacketID.CharMsgStatusNPC
            Call HandleCharMsgStatusNPC
        
        Case ServerPacketID.AbrirFormularios
            Call HandleAbrirFormularios
            
        Case ServerPacketID.ChangeInventorySlotUser
            Call HandleChangeInventorySlotUser

        Case ServerPacketID.AuraToChar
            Call HandleAuraToChar

        Case ServerPacketID.UpdateSed
            Call HandleUpdateSed
       
        Case ServerPacketID.UpdateHambre
            Call HandleUpdateHambre
       
        Case ServerPacketID.EjecutarAccion
            Call HandleEjecutarAccion

        Case Else
            'ERROR : Abort!
            Exit Sub

    End Select
    
    'Done with this packet, move on to next one
    If incomingData.Length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData

    End If

End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()

    On Error GoTo HandleLogged_Err
    
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Security.Redundance = incomingData.ReadByte()
    
    ' Variable initialization
    EngineRun = True
    Nombres = True
    
    'Set connected state
    Call SetConnected
    
    Exit Sub

HandleLogged_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleLogged", Erl)
    Resume Next
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()

    On Error GoTo HandleRemoveDialogs_Err
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Char_Dialog_Remove_All
    
    Exit Sub

HandleRemoveDialogs_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleRemoveDialogs", Erl)
    Resume Next
    
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()

    On Error GoTo HandleRemoveCharDialog_Err
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Char_Dialog_Remove(incomingData.ReadInteger)
    
    Exit Sub

HandleRemoveCharDialog_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleRemoveCharDialog", Erl)
    Resume Next
    
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()

    On Error GoTo HandleNavigateToggle_Err
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando

    Exit Sub

HandleNavigateToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleMontateToggle()

    On Error GoTo HandleMontateToggle_Err
 
    Call incomingData.ReadByte
    
    CurrentUser.Montando = Not CurrentUser.Montando

    If CurrentUser.Montando = True Then
        Engine_Scroll_Pixels scroll_pixels_per_frameBackUp * velocidadMontando
    Else
        Engine_Scroll_Pixels scroll_pixels_per_frameBackUp
    End If
    
    Exit Sub

HandleMontateToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleMontateToggle", Erl)
    Resume Next
    
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()

   
    'Remove packet ID
    Call incomingData.ReadByte

   ' Call ResetAllInfo(False)
    
    Call CloseConnectionAndResetAllInfo
    
End Sub
Private Sub CloseConnectionAndResetAllInfo()

  
   ' Call ResetAllInfo(False)
    
    If CheckUserData() Then
        frmMain.Visible = False
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()

    On Error GoTo HandleCommerceEnd_Err
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    
    'Clear item's list
    frmComerciar.List1(0).Clear
    frmComerciar.List1(1).Clear
    
    'Reset vars
    Comerciando = False
    
    'Hide form
    Unload frmComerciar

    Exit Sub

HandleCommerceEnd_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
    Resume Next
    
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()

    On Error GoTo HandleBankEnd_Err
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    
    frmBancoObj.List1(0).Clear
    frmBancoObj.List1(1).Clear
    
    Unload frmBancoObj
    Comerciando = False

    Exit Sub

HandleBankEnd_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBankEnd", Erl)
    Resume Next
    
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
    
    On Error GoTo HandleCommerceInit_Err
 
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            frmComerciar.List1(1).AddItem Inventario.ItemName(i)
        Else
            frmComerciar.List1(1).AddItem "(" & Locale_GUI_Frase(269) & ")"
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain

    Exit Sub

HandleCommerceInit_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCommerceInit", Erl)
    Resume Next
    
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()

    On Error GoTo HandleBankInit_Err
    
    
    'Check packet is complete
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim MiGLD As Long
    Dim MiItem As Byte
    Dim Goliath As Byte
    
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Goliath = incomingData.ReadByte
    MiGLD = incomingData.ReadLong
    MiItem = incomingData.ReadByte
    
    If Goliath Then
        Call frmGoliath.ParseBancoInfo(MiGLD, MiItem)
    Else
        Call frmBancoObj.List1(1).Clear
        
        'Fill the inventory list
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
            Else
                frmBancoObj.List1(1).AddItem "(" & Locale_GUI_Frase(269) & ")"
            End If
        Next i
        
        Call frmBancoObj.List1(0).Clear
        
        'Fill the bank list
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            If UserBancoInventory(i).OBJIndex <> 0 Then
                frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
            Else
                frmBancoObj.List1(0).AddItem "(" & Locale_GUI_Frase(269) & ")"
            End If
        Next i

        'Set state and show form
        Comerciando = True
 
        frmBancoObj.Show vbModeless, frmMain
    End If
    
    Exit Sub

HandleBankInit_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBankInit", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
    
    On Error GoTo HandleUpdateSta_Err

    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , , , incomingData.ReadInteger)
    
    Exit Sub

HandleUpdateSta_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateSta", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()

    On Error GoTo HandleUpdateMana_Err

    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , incomingData.ReadInteger)
    
    Exit Sub

HandleUpdateMana_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateMana", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
    
    On Error GoTo HandleUpdateHP_Err
    
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, incomingData.ReadInteger)
    
    Exit Sub

HandleUpdateHP_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateHP", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
    
    On Error GoTo HandleUpdateGold_Err

    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Call ClientTCP.ActualizarEst(, , , , , , incomingData.ReadLong())

    Exit Sub

HandleUpdateGold_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateGold", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
    
    On Error GoTo HandleUpdateExp_Err

    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , , , , , , , incomingData.ReadLong)

    Exit Sub

HandleUpdateExp_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateExp", Erl)
    Resume Next
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()

    On Error GoTo HandleUpdateStrenght_Err
    
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , , , , , , , , CInt(incomingData.ReadByte))
    
    Exit Sub
  
HandleUpdateStrenght_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateStrenght", Erl)
    Resume Next
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
    
    On Error GoTo HandleUpdateDexterity_Err
    
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
 
    Call ClientTCP.ActualizarEst(, , , , , , , , , , , CInt(incomingData.ReadByte))
    
    Exit Sub

HandleUpdateDexterity_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateDexterity", Erl)
    Resume Next
    
End Sub
Private Sub HandleChangeMap()
 
    On Error GoTo HandleChangeMap_Err
    

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserMap = incomingData.ReadInteger
    
    Call SwitchMap(CurrentUser.UserMap)

    Exit Sub

HandleChangeMap_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChangeMap", Erl)
    Resume Next
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()

    On Error GoTo HandlePosUpdate_Err
    
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Map_RemoveOldUser
    
    '// Seteamos la Posicion en el Mapa
    Call Char_MapPosSet(incomingData.ReadByte, incomingData.ReadByte)

    'Update pos label
    Call Char_UserPos
 
    Exit Sub

HandlePosUpdate_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandlePosUpdate", Erl)
    Resume Next
    
End Sub
 Private Sub HandleChatOverHead()

    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
   End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String, Name As String
    Dim charindex As Integer
    Dim ModeChat As Byte
    
    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()
    ModeChat = buffer.ReadByte()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Name = charlist(charindex).Nombre
 
    If Not NickIgnorado(Name) Then
       If Len(chat) < 1 Then
            Char_Dialog_Remove charindex
            
        Else

            If charlist(charindex).EsUsuario Then
                Dim lC As Byte
                
                Select Case ModeChat
                    Case 1
                        lC = IIf((charlist(charindex).priv > 6), 14, 7)
                        Call AddtoRichTextBox("[" & Name & "] " & chat, 0, 0, 0, 0, 0, 0, lC)

                    Case 2
                        lC = 25 'Gritar
                    
                    Case 3
                        lC = 23 'Global
                    
                    Case 4
                        lC = 22 'MP
                End Select
            End If
            
            If Char_Check(charindex) Then
                Call Char_Dialog_Create(charindex, chat, D3DColorXRGB(FontTypes(lC).red, FontTypes(lC).green, FontTypes(lC).blue))
            End If
            
        End If
    End If
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then Err.Raise error

End Sub
Private Sub HandleChatOverHeadLocale()

     If incomingData.Length < 4 Then
         Err.Raise incomingData.NotEnoughDataErrCode
         Exit Sub
    End If
    
     On Error GoTo ErrHandler

     'Remove packet ID
5    Call incomingData.ReadByte
    
     Dim Modo As Integer
     Dim charindex As Integer
     Dim id As Long
     Dim color As Long
    
6    charindex = incomingData.ReadInteger
7    id = incomingData.ReadLong
     Modo = incomingData.ReadByte
     
     Select Case Modo
     
        Case 1
            color = D3DColorXRGB(128, 128, 0)
     
        Case 2
            color = D3DColorXRGB(128, 128, 0) 'Golpe
     
        Case 3
            color = D3DColorXRGB(50, 100, 200)
     
        Case 4
            color = D3DColorXRGB(255, 255, 255)
     
     End Select
     
     
     If Char_Check(charindex) Then
        Select Case Modo
        
            Case 1 'Palabras mágicas
                Call Char_Dialog_Create(charindex, General_Locale_Spells(id, 5), color, 1)
            
            Case 2, 3 'Daño
                
                Dim daño As String
                daño = CStr(id)
                
                If id > 2 Then daño = "¡" & daño
                
                If id > 2 Then
                    Call Char_Dialog_Create(charindex, daño, color, 3)
                    
                Else
                    If CurrentUser.UserCharIndex <> charindex Then
                        Call Char_Dialog_Create(charindex, CStr(General_Locale_SMG(92, 0)), color, 3) 'Falla
                    Else
                        Call Char_Dialog_Create(charindex, CStr(General_Locale_SMG(86, 0)), color, 3) 'Fallas
                    End If
                    
                End If
                
            Case 4 'Dialogo NPCs
            
                RemoveDialogsNPCArea
                                
                If id = 592 Then
                    Call Char_Dialog_Create(charindex, Locale_GUI_Frase(id), color, 1)
                Else
                    Call Char_Dialog_Create(charindex, General_Locale_NPCs(id, 1), color, 1)
                End If
    
        End Select
        
    End If
    
    Exit Sub
 
ErrHandler:

    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChatOverHeadLocale", Erl)
    Resume Next
End Sub

Private Sub HandleConsoleMessage()

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim fontIndex As Byte
     
    chat = buffer.ReadASCIIString()
    fontIndex = buffer.ReadByte()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    If fontIndex = 12 Then
    
        If HabilitarMensajesGlobales Then
            AddtoRichTextBox chat, 0, 0, 0, 0, 0, 0, fontIndex
        End If
        
    Else
        AddtoRichTextBox chat, 0, 0, 0, 0, 0, 0, fontIndex
    End If

ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    
    chat = buffer.ReadASCIIString()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

    Call AddtoRichTextBox(chat, 0, 0, 0, 0, 0, 0, 13)
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim msg As String
    Dim Pregunta As Boolean
    Dim Accion As Byte
    
    msg = buffer.ReadASCIIString()
    Pregunta = buffer.ReadBoolean()
    Accion = buffer.ReadByte()
 
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
     
    If Pregunta = False Then
        If IsNumeric(msg) Then
            msg = Locale_Error(CInt(msg))
        End If
        
        frmMensaje.msg.Caption = msg
    
    Else
        frmPregunta.SetAccion Accion, msg
    End If
    

    If frmMain.Visible Then
        If Pregunta = False Then
            frmMensaje.Show , frmMain
        Else
            frmPregunta.Show , frmMain
        End If
        
    ElseIf frmCrearCuenta.Visible Then
    
        Call FormParser.Parse_Form(frmCrearCuenta)
        
        If Pregunta = False Then
            frmMensaje.Show vbModal, frmCrearCuenta
        Else
            frmPregunta.Show vbModal, frmCrearCuenta
        End If
        
    ElseIf frmConnect.Visible Then
        
        Call frmMain.ResetCharDisconnect
        
        Call FormParser.Parse_Form(frmConnect)
        
        If Pregunta = False Then
            frmMensaje.Show vbModal, frmConnect
        Else
            frmPregunta.Show vbModal, frmConnect
        End If
        
    ElseIf frmIniciando.Visible Then
        
        Call FormParser.Parse_Form(frmIniciando)
        
        frmCharList.Visible = True
        frmIniciando.Visible = False
        
        If Pregunta = False Then
            frmMensaje.Show vbModal, frmCharList
        Else
            frmPregunta.Show vbModal, frmCharList
        End If
        
    ElseIf frmCrearPersonaje.Visible Then
        
        Call FormParser.Parse_Form(frmCrearPersonaje)
          
        If Pregunta = False Then
            frmMensaje.Show vbModal, frmCrearPersonaje
        Else
            frmPregunta.Show vbModal, frmCrearPersonaje
        End If
        
    ElseIf frmCharList.Visible Then
        
        Call FormParser.Parse_Form(frmCharList)
        If Pregunta = False Then
            frmMensaje.Show vbModal, frmCharList
        Else
            frmPregunta.Show vbModal, frmCharList
        End If
        
    End If

 
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
    
    On Error GoTo HandleUserIndexInServer_Err
    
1    If incomingData.Length < 3 Then
2        Err.Raise incomingData.NotEnoughDataErrCode
3        Exit Sub

4    End If
    
    'Remove packet ID
5    Call incomingData.ReadByte
    
6    UserIndex = incomingData.ReadInteger
    
7    Exit Sub

HandleUserIndexInServer_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUserIndexInServer", Erl)
    Resume Next
    
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()

    On Error GoTo HandleUserCharIndexInServer_Err

1    If incomingData.Length < 3 Then
2        Err.Raise incomingData.NotEnoughDataErrCode
3        Exit Sub

4    End If
    
    'Remove packet ID
5    Call incomingData.ReadByte

    Call Char_UserIndexSet(incomingData.ReadInteger)
    
    'Update pos label
    Call Char_UserPos
    
11  Exit Sub

HandleUserCharIndexInServer_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUserCharIndexInServer", Erl)
    Resume Next
    
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    If incomingData.Length < 31 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim charindex As Integer
    Dim body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim Weapon As Integer
    Dim Shield As Integer
    Dim Helmet As Integer
    Dim Name As String
    Dim ParticulaFx As Byte
    Dim FxIndex As Integer
    Dim FXLoop As Integer
    Dim priv As Byte
    Dim donador As Byte
    Dim Arma_Aura As Byte
    Dim Body_Aura As Byte
    Dim Escudo_Aura As Byte
    Dim Head_Aura As Byte
    Dim Otra_Aura As Byte
    Dim Anillo_Aura As Byte
    
    charindex = buffer.ReadInteger()
    body = buffer.ReadInteger()
    Head = buffer.ReadInteger()
    Heading = buffer.ReadByte()
    X = buffer.ReadByte()
    Y = buffer.ReadByte()
    Weapon = buffer.ReadInteger()
    Shield = buffer.ReadInteger()
    Helmet = buffer.ReadInteger()
    FxIndex = buffer.ReadInteger()
    FXLoop = buffer.ReadInteger()
    Name = buffer.ReadASCIIString()
    priv = buffer.ReadByte()
    donador = buffer.ReadByte()
    ParticulaFx = buffer.ReadByte()
    Arma_Aura = buffer.ReadByte()
    Body_Aura = buffer.ReadByte()
    Escudo_Aura = buffer.ReadByte()
    Head_Aura = buffer.ReadByte()
    Otra_Aura = buffer.ReadByte()
    Anillo_Aura = buffer.ReadByte()

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    With charlist(charindex)
 
        Call SetCharacterFx(charindex, FxIndex, FXLoop)
        
        If IsNumeric(Name) Then 'Si es un NPC damos nombre
            
            Dim Aux As String
            Aux = CInt(Name)
            
            .EsNPC = True
            .EsUsuario = False
            
            Name = General_Locale_NPCs(Aux, 0)
            
            If priv = 1 Then
                .priv = 2
            ElseIf priv = 2 Then
                .priv = 3
            Else
                .priv = priv
            End If
            
            If General_Locale_NPCs(Aux, 4) = 1 Then 'Si es hostil
                .NPCHostil = True
            Else
                .NPCHostil = False
            End If
            
         Else ' Es User
            .EsUsuario = True
            .EsNPC = False
            .priv = priv
         End If
     
        .donador = donador
        
         ParticulaFx = ParticulaFx
        .Arma_Aura = Arma_Aura
        .Body_Aura = Body_Aura
        .Escudo_Aura = Escudo_Aura
        .Head_Aura = Head_Aura
        .Otra_Aura = Otra_Aura
        .Anillo_Aura = Anillo_Aura
    
        If (.Pos.X <> 0 And .Pos.Y <> 0) Then
            If MapData(.Pos.X, .Pos.Y).charindex = charindex Then
            
                'Erase the old character from map
                MapData(charlist(charindex).Pos.X, charlist(charindex).Pos.Y).charindex = 0
            End If

        End If
        
        Call UpdateTagAndNameChar(charindex, Name)
    
        Call MakeChar(charindex, body, Head, Heading, X, Y, Weapon, Shield, Helmet, ParticulaFx)
        
        Call RefreshAllChars
    End With
 
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub
Private Sub HandleCharacterRemove()

    On Error GoTo HandleCharacterRemove_Err

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    Dim Desvanecido As Boolean
    
    charindex = incomingData.ReadInteger
    Desvanecido = incomingData.ReadBoolean
    
    If Rendimiento = 1 Then
        If Desvanecido Then
            Call CrearFantasma(charindex)
        End If
    End If
    
    Call EraseChar(charindex)
    Call RefreshAllChars

    Exit Sub

HandleCharacterRemove_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCharacterRemove", Erl)
    Resume Next
    
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
    
    On Error GoTo HandleCharacterMove_Err

    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    charindex = incomingData.ReadInteger
    X = incomingData.ReadByte
    Y = incomingData.ReadByte
    
    With charlist(charindex)
    
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 7 And .priv <> 8 And .priv <> 9 Then
            Call DoPasosFx(charindex)
        End If

    End With
    
    Call MoveCharbyPos(charindex, X, Y)
    
    Call RefreshAllChars
    
    Exit Sub

HandleCharacterMove_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCharacterMove", Erl)
    Resume Next
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    On Error GoTo HandleForceCharMove_Err
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)
    
    Call MoveCharbyHead(CurrentUser.UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars

    Exit Sub

HandleForceCharMove_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleForceCharMove", Erl)
    Resume Next
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    
    If incomingData.Length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    Dim tempInt As Integer
    Dim headIndex As Integer
    Dim FxIndex As Integer
    Dim FXLoop As Integer
    
    charindex = incomingData.ReadInteger
    
    With charlist(charindex)
    
        tempInt = incomingData.ReadInteger()
        
        If tempInt < LBound(BodyData()) Or tempInt > UBound(BodyData()) Then
            .body = BodyData(0)
            .iBody = 0
        Else
            .body = BodyData(tempInt)
            .iBody = tempInt
        End If
        
        headIndex = incomingData.ReadInteger()
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        Dim oldMuerto As Boolean
        oldMuerto = .Muerto
        
        .Muerto = (headIndex = CASPER_HEAD)

        If .Muerto = False And oldMuerto = True Then
            Call Char_Particle_Group_Remove(charindex, 22)
        End If
        
        .Heading = incomingData.ReadByte()
        
        tempInt = incomingData.ReadInteger()

        If tempInt <> 0 Then .Arma = WeaponAnimData(tempInt)
        
        tempInt = incomingData.ReadInteger()

        If tempInt <> 0 Then .Escudo = ShieldAnimData(tempInt)
        
        tempInt = incomingData.ReadInteger()

        If tempInt <> 0 Then .Casco = CascoAnimData(tempInt)
        
        FxIndex = incomingData.ReadInteger()
        FXLoop = incomingData.ReadInteger()
        
        Call SetCharacterFx(charindex, FxIndex, FXLoop)

    End With
    
    Call RefreshAllChars

    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCharacterChange", Erl)
    Resume Next
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChangeSlot()
    
    On Error GoTo HandleCharacterChangeSlot_Err
    
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer, SlotIndex As Integer
    Dim Index As Integer

    charindex = incomingData.ReadInteger
    SlotIndex = incomingData.ReadInteger
    Index = incomingData.ReadByte
    
    With charlist(charindex)
            
        Select Case Index
        
        Case 1 ' Body
        
            If SlotIndex < LBound(BodyData()) Or SlotIndex > UBound(BodyData()) Then
                .body = BodyData(0)
                .iBody = 0
            Else
                .body = BodyData(SlotIndex)
                .iBody = SlotIndex
            End If
            
        Case 2 ' Head
            
            If SlotIndex < LBound(HeadData()) Or SlotIndex > UBound(HeadData()) Then
                .Head = HeadData(0)
                .iHead = 0
            Else
                .Head = HeadData(SlotIndex)
                .iHead = SlotIndex
            End If
            
            Dim oldMuerto As Boolean: oldMuerto = .Muerto
            
            .Muerto = (SlotIndex = CASPER_HEAD)
    
            If .Muerto = False And oldMuerto = True Then Call Char_Particle_Group_Remove(charindex, 22)
            
        Case 3 'Heading
        
            .Heading = CByte(SlotIndex)
        
        Case 4 ' Weapon
        
            If SlotIndex <> 0 Then
                .Arma = WeaponAnimData(SlotIndex)
                
                If SlotIndex <> 2 Then Call Audio.PlayWave("25.wav", charlist(charindex).Pos.X, charlist(charindex).Pos.Y)
                
            End If
            
            
        Case 5 ' Shield
        
            If SlotIndex <> 0 Then .Escudo = ShieldAnimData(SlotIndex)
            Call Audio.PlayWave("224.wav", charlist(charindex).Pos.X, charlist(charindex).Pos.Y)
            
        Case 6 'Helmet
            
            If SlotIndex <> 0 Then .Casco = CascoAnimData(SlotIndex)
            
        Case 7 ' Fx and Loops
            Call incomingData.ReadInteger
            'Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
            
        Case 8 ' Nudis
        
            If SlotIndex <> 0 Then .Arma = WeaponAnimData(SlotIndex)
                        Call Audio.PlayWave("223.wav", charlist(charindex).Pos.X, charlist(charindex).Pos.Y)

            
        End Select
    
    End With
    
    Call RefreshAllChars

    Exit Sub

HandleCharacterChangeSlot_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCharacterChangeSlot", Erl)
    Resume Next
    
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
    
    On Error GoTo HandleObjectCreate_Err

    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    Dim Obj As Integer
    Dim Amount As Integer
    
    X = incomingData.ReadByte
    Y = incomingData.ReadByte
    
    Obj = incomingData.ReadInteger
    Amount = incomingData.ReadInteger
    
    Map_Obj_Create X, Y, CInt(General_Locale_Obj(Obj, 3)), Obj, CByte(General_Locale_Obj(Obj, 2)), Amount
    
    Exit Sub

HandleObjectCreate_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleObjectCreate", Erl)
    Resume Next
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
    
    On Error GoTo HandleObjectDelete_Err

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    Map_Obj_Delete X, Y
    
    Exit Sub

HandleObjectDelete_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleObjectDelete", Erl)
    Resume Next
    
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
    
    On Error GoTo HandleBlockPosition_Err

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte
    Y = incomingData.ReadByte
    
    If incomingData.ReadBoolean Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If

    Exit Sub

HandleBlockPosition_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlockPosition", Erl)
    Resume Next
    
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
    
    On Error GoTo HandlePlayMIDI_Err

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Byte
    Dim Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadByte
    Loops = incomingData.ReadInteger
    
    If currentMidi > 0 Then
        End If
    
 
    
    Exit Sub

HandlePlayMIDI_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandlePlayMIDI", Erl)
    Resume Next
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()

    On Error GoTo HandlePlayWave_Err

    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Integer
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadInteger
    srcX = incomingData.ReadByte
    srcY = incomingData.ReadByte
    
    If wave = 105 Then 'Trueno
        trueno = 20
        Exit Sub
    End If
    
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
        
    Exit Sub

HandlePlayWave_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandlePlayWave", Erl)
    Resume Next
    
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildAdm
        'Clear guild's list
        .guildslist.Clear
        
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long

        For i = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(buffer)
        
        .Show vbModeless, frmMain

    End With
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
    
    On Error GoTo HandleAreaChanged_Err

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    Call CambioDeArea(X, Y)

    Exit Sub

HandleAreaChanged_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleAreaChanged", Erl)
    Resume Next
    
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()

    On Error GoTo HandlePauseToggle_Err

    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa

    Exit Sub

HandlePauseToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandlePauseToggle", Erl)
    Resume Next
    
End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()

    On Error GoTo HandleRainToggle_Err

    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
        
    'Remove packet ID
    
    Call incomingData.ReadByte

    Queclima = incomingData.ReadByte()
    
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.BAJOTECHO Or _
        MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.trigger_2 Or _
        MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.ZONASEGURA)

    
        
    If (bRain And Queclima = 0) Then
        'Stop playing the rain sound
        Call Audio.StopWave(RainBufferIndex)
        RainBufferIndex = 0
        If bTecho Then
'            Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
        Else
'            Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
        End If
        frmMain.IsPlaying = PlayLoop.plNone
        Particle_Group_Remove meteo_particle
        meteo_particle = 0
          bRain = False
    ElseIf Queclima > 0 Then
        If Queclima = 1 Then
        meteo_particle = SetMapParticle(58, -1, -1) ' 8 o 58
        'frmMain.imgHora.Picture = LoadPicture(App.path & "\Recursos\Interfaces\lluvia.jpg")
        ElseIf Queclima = 2 Then
        meteo_particle = SetMapParticle(8, -1, -1)
        'frmMain.imgHora.Picture = LoadPicture(App.path & "\Recursos\Interfaces\electrica.jpg")
        ElseIf Queclima = 3 Then
        meteo_particle = SetMapParticle(57, -1, -1)  '56, 57,13,
        'frmMain.imgHora.Picture = LoadPicture(App.path & "\Recursos\Interfaces\nieve.jpg")
        End If
          bRain = True
    End If
    
    Exit Sub

HandleRainToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleRainToggle", Erl)
    Resume Next
    
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
    
    On Error GoTo HandleCreateFX_Err

    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    charindex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(charindex, fX, Loops)

    Exit Sub

HandleCreateFX_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCreateFX", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
    
    On Error GoTo HandleUpdateUserStats_Err
    
    If incomingData.Length < 33 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    


    Call ClientTCP.ActualizarEst(incomingData.ReadInteger, incomingData.ReadInteger, incomingData.ReadInteger, incomingData.ReadInteger, _
                                 incomingData.ReadInteger, incomingData.ReadInteger, incomingData.ReadLong, incomingData.ReadInteger, _
                                 incomingData.ReadLong, incomingData.ReadLong, CInt(incomingData.ReadByte), CInt(incomingData.ReadByte), CInt(incomingData.ReadByte), _
                                 CInt(incomingData.ReadByte), CInt(incomingData.ReadByte), CInt(incomingData.ReadByte), True)
    
    Exit Sub
    
HandleUpdateUserStats_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateUserStats", Erl)
    Resume Next
    
End Sub

Private Sub HandleUpdateUserStatsForLevel()
    
    On Error GoTo UpdateUserStatsForLevel_Err
    
    If incomingData.Length < 17 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim HP As Integer
    HP = incomingData.ReadInteger
    
    Call ClientTCP.ActualizarEst(HP, HP, incomingData.ReadInteger, , incomingData.ReadInteger, , , incomingData.ReadInteger, incomingData.ReadLong, incomingData.ReadLong, , , , , , , True)
    
    Exit Sub
    
UpdateUserStatsForLevel_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.UpdateUserStatsForLevel", Erl)
    Resume Next
    
End Sub


''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
    
    On Error GoTo HandleWorkRequestTarget_Err

    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UsingSkill = incomingData.ReadByte()
    
    Select Case CurrentUser.UsingSkill
        
        Case magia
            Call AddtoRichTextBox(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
            If Not FormParser.GetDefaultCursor(frmMain) = E_CAST Then _
                Call FormParser.Parse_Form(frmMain, E_CAST)

        Case pesca, robar, talar, mineria, FundirMetal, domar
            Call AddtoRichTextBox(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
            If Not FormParser.GetDefaultCursor(frmMain) = E_SHOOT Then _
                Call FormParser.Parse_Form(frmMain, E_SHOOT)
                
        Case proyectiles
            Call AddtoRichTextBox(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
            If Not FormParser.GetDefaultCursor(frmMain) = E_ARROW Then _
                Call FormParser.Parse_Form(frmMain, E_ARROW)
            
        Case armasarrojadizas
            Call AddtoRichTextBox(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
            If Not FormParser.GetDefaultCursor(frmMain) = E_ATTACK Then _
                Call FormParser.Parse_Form(frmMain, E_ATTACK)
            
        Case Else
            Call FormParser.Parse_Form(frmMain)
            
    End Select
    
    Exit Sub

HandleWorkRequestTarget_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleWorkRequestTarget", Erl)
    Resume Next
    
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    If incomingData.Length < 12 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Slot As Byte, Puede As Byte
    Dim OBJIndex As Integer, Amount As Integer
    Dim Equipped As Boolean
    Dim Value As Single
    
    Slot = incomingData.ReadByte
    OBJIndex = incomingData.ReadInteger
    Amount = incomingData.ReadInteger
    
    Equipped = incomingData.ReadBoolean
    Value = incomingData.ReadSingle
    Puede = incomingData.ReadByte

    Call Inventario.SetItem(Slot, OBJIndex, Amount, Equipped, Value, Puede)
    
    'If CantidadEnMacros Then Call UpdateMacroLabels(1)
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChangeInventorySlot", Erl)
    Resume Next
    
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    If incomingData.Length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errorhandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Slot As Byte

    Slot = incomingData.ReadByte()
    
    With UserBancoInventory(Slot)
        .OBJIndex = incomingData.ReadInteger()
        .Amount = incomingData.ReadInteger()
        .Valor = incomingData.ReadLong()
        
        .Name = General_Locale_Obj(.OBJIndex, 0)
        .ObjType = CInt(General_Locale_Obj(.OBJIndex, 2))
        .GrhIndex = CInt(General_Locale_Obj(.OBJIndex, 3))
        .MaxDef = CInt(General_Locale_Obj(.OBJIndex, 5))
        .MinDef = CInt(General_Locale_Obj(.OBJIndex, 6))
        .MaxHit = CInt(General_Locale_Obj(.OBJIndex, 7))
        .MinHit = CInt(General_Locale_Obj(.OBJIndex, 8))
        
    If frmBancoObj.List1(0).ListCount >= Slot Then
        Call frmBancoObj.List1(0).RemoveItem(Slot - 1)
    End If
    
    Call frmBancoObj.List1(0).AddItem(IIf(.Name <> "", .Name, "(" & Locale_GUI_Frase(269) & ")"), Slot - 1)
    
    End With
    
    Exit Sub
    
errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChangeBankSlot", Erl)
    Resume Next
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    Dim Slot As Byte
    Dim hechizo As Integer
    
    Call incomingData.ReadByte
 
    Slot = incomingData.ReadByte
 
    UserHechizos(Slot) = incomingData.ReadInteger
    
    If Slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.list(Slot - 1) = General_Locale_Spells(UserHechizos(Slot), 0)
    Else
        Call frmMain.hlst.AddItem(General_Locale_Spells(UserHechizos(Slot), 0))
    End If

    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChangeSpellSlot", Erl)
    Resume Next
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()

On Error GoTo HandleAtributes_Err

    If incomingData.Length < 1 + NUMATRIBUTOS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTOS
        CurrentUser.UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    If LlegaronSkills And LlegaronStats Then
        Alocados = SkillPoints
        frmEstadisticas.Puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show , frmMain
    Else
        LlegaronAtrib = True
    End If
    
    Exit Sub

HandleAtributes_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleAtributes", Erl)
    Resume Next
    
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String
    
    Count = incomingData.ReadInteger()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstArmas.AddItem(tmp)
        ArmasHerrero(i) = incomingData.ReadInteger()
    Next i
    
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i) = 0
    Next i

    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlacksmithWeapons", Erl)
    Resume Next
    
End Sub


Private Sub HandleBlacksmithArmors()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String
    
    Count = incomingData.ReadInteger()
    
    Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstArmaduras.AddItem(tmp)
        ArmadurasHerrero(i) = incomingData.ReadInteger()
    Next i
    
    For i = i To UBound(ArmadurasHerrero())
        ArmadurasHerrero(i) = 0
    Next i
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlacksmithArmors", Erl)
    Resume Next
    
End Sub

Private Sub HandleBlacksmithHelmet()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String
    
    Count = incomingData.ReadInteger()
    
    Call frmHerrero.lstCascos.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstCascos.AddItem(tmp)
        CascosHerrero(i) = incomingData.ReadInteger()
    Next i
    
    For i = i To UBound(CascosHerrero())
        CascosHerrero(i) = 0
    Next i

    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlacksmithHelmet", Erl)
    Resume Next
    
End Sub


Private Sub HandleBlacksmithShield()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String
    
    Count = incomingData.ReadInteger()
    
    Call frmHerrero.lstEscudos.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(incomingData.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstEscudos.AddItem(tmp)
        EscudosHerrero(i) = incomingData.ReadInteger()
    Next i
    
    For i = i To UBound(EscudosHerrero())
        EscudosHerrero(i) = 0
    Next i
 
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlacksmithShield", Erl)
    Resume Next

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String
    
    Count = incomingData.ReadInteger()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & ")"    'The wood needed
        
        Call frmCarp.lstArmas.AddItem(tmp)
        ObjCarpintero(i) = incomingData.ReadInteger()
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i


    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlacksmithShield", Erl)
    Resume Next

End Sub


Private Sub HandleAlquimiaObjects()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String
    
    Count = incomingData.ReadInteger()
    
    Call frmDruida.lstPociones.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & ")"    'The wood needed
        
        Call frmDruida.lstPociones.AddItem(tmp)
        ObjAlquimia(i) = incomingData.ReadInteger()
    Next i
    
    For i = i To UBound(ObjAlquimia())
        ObjAlquimia(i) = 0
    Next i
 
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleAlquimiaObjects", Erl)
    Resume Next

End Sub

Private Sub HandleSastreObjects()

      
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim tmp As String

    Count = incomingData.ReadInteger()

    Call frmSastre.lstRopas.Clear
    
    For i = 1 To Count
        tmp = General_Locale_Obj(incomingData.ReadInteger(), 0) & " ("           'Get the object's name
        tmp = tmp & CStr(incomingData.ReadInteger()) & "/" & _
        CStr(incomingData.ReadInteger()) & "/" & _
        CStr(incomingData.ReadInteger()) & ")"
        
        Call frmSastre.lstRopas.AddItem(tmp)
        ObjSastre(i) = incomingData.ReadInteger()
    Next i
  
    For i = i To UBound(ObjSastre())
        ObjSastre(i) = 0
    Next i
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleSastreObjects", Erl)
    Resume Next

End Sub

Private Sub HandleSendMsgBox()

    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
    On Error GoTo ErrHandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    Dim Mensaje As String
    Dim Modo As Integer
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Mensaje = buffer.ReadASCIIString()
    Modo = CInt(buffer.ReadByte())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    If Modo > 0 Then
        Mensaje = Locale_Error(CInt(Mensaje))
    End If
    
    Call MsgBox(Mensaje)
    
    If frmRecuperarCuenta.Visible = True Then
        Call FormParser.Parse_Form(frmRecuperarCuenta)
    
    ElseIf frmCambiarContraseña.Visible = True Then
        Call FormParser.Parse_Form(frmCambiarContraseña)
        
    ElseIf frmConnect.Visible = True Then
        Call FormParser.Parse_Form(frmConnect)
    
    ElseIf frmCharList.Visible = True Then
        Call FormParser.Parse_Form(frmCharList)
        
    End If
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
    
    On Error GoTo HandleBlind_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True

    Exit Sub

HandleBlind_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlind", Erl)
    Resume Next
    
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
    
    On Error GoTo HandleDumb_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
    
    Exit Sub

HandleDumb_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleDumb", Erl)
    Resume Next
    
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    If incomingData.Length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Slot As Byte

    Slot = incomingData.ReadByte()
    
    With NPCInventory(Slot)
    
        .Amount = incomingData.ReadInteger()
        .Valor = incomingData.ReadSingle()
        .OBJIndex = incomingData.ReadInteger()
        
        .ObjType = CInt(General_Locale_Obj(.OBJIndex, 2))
        .MaxHit = CInt(General_Locale_Obj(.OBJIndex, 7))
        .MinHit = CInt(General_Locale_Obj(.OBJIndex, 8))
        .MaxDef = CInt(General_Locale_Obj(.OBJIndex, 5))
        .MinDef = CInt(General_Locale_Obj(.OBJIndex, 6))
        .GrhIndex = CInt(General_Locale_Obj(.OBJIndex, 3))
        .Name = General_Locale_Obj(.OBJIndex, 0)
        
    
        If frmComerciar.List1(0).ListCount >= Slot Then
            Call frmComerciar.List1(0).RemoveItem(Slot - 1)
        End If
        
        Call frmComerciar.List1(0).AddItem(IIf(.Name <> "", .Name, "(" & Locale_GUI_Frase(269) & ")"), Slot - 1)
    
    End With
        
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChangeNPCInventorySlot", Erl)
    Resume Next
End Sub


''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
    
    On Error GoTo HandleUpdateHungerAndThirst_Err

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , , , , , , , , , , CInt(incomingData.ReadByte), , CInt(incomingData.ReadByte))
    
    Exit Sub

HandleUpdateHungerAndThirst_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateHungerAndThirst", Erl)
    Resume Next
    
End Sub
Private Sub HandleMiniStats()

    On Error GoTo HandleMiniStats_Err

    If incomingData.Length < 39 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .Raza = incomingData.ReadByte
        .Genero = incomingData.ReadByte
        .MuertesUsuario = incomingData.ReadLong()
        .status = incomingData.ReadByte
        .RepublicanosMatados = incomingData.ReadLong()
        .CaosMatados = incomingData.ReadLong()
        .ArmadasRealesMatados = incomingData.ReadLong()
        .MiliciasMatados = incomingData.ReadLong()
    End With

    
    If LlegaronAtrib And LlegaronSkills Then
        Alocados = SkillPoints
        frmEstadisticas.Puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show , frmMain
    Else
        LlegaronStats = True
    End If

    Exit Sub

HandleMiniStats_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleMiniStats", Erl)
    Resume Next
    
End Sub

Private Sub HandleLevelUp()
    
    On Error GoTo HandleLevelUp_Err

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    
    Exit Sub

HandleLevelUp_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleLevelUp", Erl)
    Resume Next
    
End Sub

Private Sub HandleSetInvisible()
    
    On Error GoTo HandleSetInvisible_Err
    
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
        
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    
    charindex = incomingData.ReadInteger()
    charlist(charindex).Invisible = incomingData.ReadBoolean()
    
    Exit Sub

HandleSetInvisible_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleSetInvisible", Erl)
    Resume Next
    
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()

    On Error GoTo HandleMeditateToggle_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar
    
    Exit Sub

HandleMeditateToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleMeditateToggle", Erl)
    Resume Next
    
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()

    On Error GoTo HandleBlindNoMore_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False

    Exit Sub

HandleBlindNoMore_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBlindNoMore", Erl)
    Resume Next
    
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()

    On Error GoTo HandleDumbNoMore_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False

    Exit Sub

HandleDumbNoMore_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleDumbNoMore", Erl)
    Resume Next
    
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()

On Error GoTo HandleSendSkills_Err

    If incomingData.Length < 1 + NUMSKILLS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        CurrentUser.UserSkills(i) = incomingData.ReadByte()
    Next i

    If LlegaronAtrib And LlegaronStats Then
        Alocados = SkillPoints
        frmEstadisticas.Puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show , frmMain
    Else
        LlegaronSkills = True
    End If

    Exit Sub

HandleSendSkills_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleSendSkills", Erl)
    Resume Next
    
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatures() As String

    Dim i As Long
    
    frmEntrenador.lstCriaturas.Clear
    
    creatures = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        creatures(i) = General_Locale_NPCs(CInt(creatures(i)), 0)
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i

    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Is optional show the frmGuildNews form
    '***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim i           As Long

    Dim sTemp       As String
    
    'Get news' string
    frmGuildNews.news = buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
     frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim vsGuildList() As String

    Dim i             As Long
    
    vsGuildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear

    For i = 0 To UBound(vsGuildList())
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear

    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

Private Sub HandleCharacterInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 31 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmCharInfo

        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True

        End If
        
        .Nombre.Caption = "Nombre: " & buffer.ReadASCIIString()
        .Raza.Caption = "Raza: " & ListaRazas(buffer.ReadByte())
        .Clase.Caption = "Clase: " & ListaClases(buffer.ReadByte())
        
        If buffer.ReadByte() = 1 Then
            .Genero.Caption = "Género: " & "Hombre"
        Else
            .Genero.Caption = "Género: " & "Mujer"

        End If
        
        .Nivel.Caption = "Nivel: " & buffer.ReadByte()
        .Oro.Caption = "Oro: " & buffer.ReadLong()
        .Banco.Caption = "Oro en Banco: " & buffer.ReadLong()
        
        
        .txtPeticiones.Text = "Peticiones a Clanes: " & buffer.ReadASCIIString()
        .guildactual.Caption = "Clan Actual: " & buffer.ReadASCIIString()
        .txtMiembro.Text = buffer.ReadASCIIString()
        
        Dim armada As Boolean

        Dim caos   As Boolean
        
        armada = buffer.ReadBoolean()
        caos = buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Faccion: " & "Armada Real - Bug xd"
        ElseIf caos Then
            .ejercito.Caption = "Faccion: " & "Legión Oscura - Bug xd"

        End If
        
        .Ciudadanos.Caption = "Ciudadanos Matados: " & CStr(buffer.ReadLong())
        .criminales.Caption = "Renegados Matados: " & CStr(buffer.ReadLong())

        Call .Show(vbModeless, frmMain)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim i      As Long

    Dim list() As String
    
    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtguildnews = buffer.ReadASCIIString()
        
        'Get list of join requests
        list = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(list())
            Call .solicitudes.AddItem(list(i))
        Next i
        
        .Show , frmMain

    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildBrief
        '.imgDeclararGuerra.Visible = .EsLeader
        '.imgOfrecerAlianza.Visible = .EsLeader
        '.imgOfrecerPaz.Visible = .EsLeader
        
        .Nombre.Caption = buffer.ReadASCIIString()
        .fundador.Caption = buffer.ReadASCIIString()
        .creacion.Caption = buffer.ReadASCIIString()
        .lider.Caption = buffer.ReadASCIIString()
        .web.Caption = buffer.ReadASCIIString()
        .Miembros.Caption = buffer.ReadInteger()
        
        If buffer.ReadBoolean() Then
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"

        End If
        
        .lblAlineacion.Caption = buffer.ReadASCIIString()
        .Enemigos.Caption = buffer.ReadInteger()
        .Aliados.Caption = buffer.ReadInteger()
        .antifaccion.Caption = buffer.ReadASCIIString()
        
        Dim codexStr() As String

        Dim i          As Long
        
        codexStr = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = buffer.ReadASCIIString()

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()

    On Error GoTo HandleParalizeOK_Err

    'Remove packet ID
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
    
    Exit Sub

HandleParalizeOK_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleParalizeOK", Erl)
    Resume Next
    
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()

    On Error GoTo HandleTradeOK_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        Call frmComerciar.List1(1).Clear
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                Call frmComerciar.List1(1).AddItem(Inventario.ItemName(i))
            Else
                Call frmComerciar.List1(1).AddItem("(" & Locale_GUI_Frase(269) & ")")
            End If
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmComerciar.LasActionBuy Then
            frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
            frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
        Else
            frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
            frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
        End If
    End If

    Exit Sub

HandleTradeOK_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleTradeOK", Erl)
    Resume Next
    
End Sub
Private Sub HandleBankOK()

    On Error GoTo HandleBankOK_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    If frmBancoObj.Visible Then
        
        Call frmBancoObj.List1(1).Clear
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                Call frmBancoObj.List1(1).AddItem(Inventario.ItemName(i))
            Else
                Call frmBancoObj.List1(1).AddItem("(" & Locale_GUI_Frase(269) & ")")
            End If
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else
            frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
    End If
       
    Exit Sub

HandleBankOK_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleBankOK", Erl)
    Resume Next
    
End Sub
 
''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    
    sosList = Split(buffer.ReadASCIIString(), SEPARATOR)
 
    For i = 0 To UBound(sosList())
        Call frmConsultas.List1.AddItem(sosList(i))
    Next i
 
    frmConsultas.Show , frmMain
        
    'frmMSG.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim userList() As String

    Dim i          As Long
    
    userList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    If frmGMPanel.Visible Then
        frmGMPanel.cboListaUsus.Clear
    
        For i = 0 To UBound(userList())
            Call frmGMPanel.cboListaUsus.AddItem(userList(i))
        Next i

        If frmGMPanel.cboListaUsus.ListCount > 0 Then frmGMPanel.cboListaUsus.ListIndex = 0

    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()

    On Error GoTo HandlePong_Err

    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte

    Dim Time As Long

    Time = incomingData.ReadLong()
    MSRender = (timeGetTime() And &H7FFFFFFF) - Time
    CurrentUser.Ping = 0

    Exit Sub

HandlePong_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandlePong", Erl)
    Resume Next
    
End Sub


''
' Handles the Pong message.

Private Sub HandleGuildMemberInfo()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long

        For i = 0 To UBound(GuildNames())
            Call .lstClanes.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(buffer)
        
        .Show vbModeless, frmMain

    End With
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim charindex As Integer
    Dim UserTag As String
    Dim priv As Byte
    Dim donador As Byte
    
    charindex = buffer.ReadInteger()
    UserTag = buffer.ReadASCIIString()
    priv = buffer.ReadByte()
    donador = buffer.ReadByte()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    With charlist(charindex)
    
        If .EsNPC = True Then
            If priv = 1 Then
                .priv = 2
            ElseIf priv = 2 Then
                .priv = 3
            End If
    
        Else
            .priv = priv
        End If
    
        .donador = donador
    
     Call ColorNombresPriv(charindex, .priv)
     Call UpdateTagAndNameChar(charindex, UserTag)

     End With
        
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

Public Sub WriteLoginExistingChar()

    On Error GoTo WriteLoginExistingChar_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        Call .WriteASCIIString(Cuenta.UserAccount)
        Call .WriteASCIIString(SEncriptar(Cuenta.UserPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(Cuenta.UserName)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        
    End With

    
    Exit Sub
    
    
WriteLoginExistingChar_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteLoginExistingChar", Erl)
    Resume Next
    
End Sub

Public Sub WriteLoginNewChar()

    On Error GoTo WriteLoginNewChar_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(Cuenta.UserAccount)
        
        Call .WriteASCIIString(SEncriptar(Cuenta.UserPassword))
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
                
        Call .WriteASCIIString(Cuenta.UserName)
         
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)

        Call .WriteByte(UserHogar)
        
        Call .WriteInteger(frmCrearPersonaje.intHeadInd)
        
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        
    End With
    
    Exit Sub

WriteLoginNewChar_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteLoginNewChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String, ByVal mode As Byte)
    
    On Error GoTo WriteTalk_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        Call .WriteASCIIString(chat)
        Call .WriteByte(mode)
    End With

    Exit Sub

WriteTalk_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteTalk", Erl)
    Resume Next
    
End Sub
Public Sub WriteWhisper(ByVal Nombre As String, ByVal chat As String)
    
    On Error GoTo WriteWhisper_Err

    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        Call .WriteASCIIString(Nombre)
        Call .WriteASCIIString(chat)

    End With
    
    Exit Sub

WriteWhisper_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteWhisper", Erl)
    Resume Next
    
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
    
    On Error GoTo WriteWalk_Err

    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        Call .WriteByte(Heading)

    End With

    Exit Sub

WriteWalk_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteWalk", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
    
    On Error GoTo WriteRequestPositionUpdate_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestPositionUpdate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
    
    Exit Sub

WriteRequestPositionUpdate_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteRequestPositionUpdate", Erl)
    Resume Next
    
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
    
    On Error GoTo WriteAttack_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Attack" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)

    Exit Sub

WriteAttack_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteAttack", Erl)
    Resume Next
    
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
    
    On Error GoTo WritePickUp_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PickUp" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)

    Exit Sub

WritePickUp_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WritePickUp", Erl)
    Resume Next
    
End Sub

''
' Writes the "CombatModeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
 
Public Sub WriteCombatModeToggle()

    On Error GoTo WriteCombatModeToggle_Err
     
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CombatModeToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CombatModeToggle)
    
    Exit Sub

WriteCombatModeToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCombatModeToggle", Erl)
    Resume Next
    
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
    
    On Error GoTo WriteResuscitationToggle_Err
    
    '**************************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
    '**************************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
    
    Exit Sub

WriteResuscitationToggle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteResuscitationToggle", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()

    On Error GoTo WriteRequestGuildLeaderInfo_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)

    Exit Sub

WriteRequestGuildLeaderInfo_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteRequestGuildLeaderInfo", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()

    On Error GoTo WriteRequestAtributes_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)

    Exit Sub

WriteRequestAtributes_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteRequestAtributes", Erl)
    Resume Next
    
End Sub
 
''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()

    On Error GoTo WriteRequestSkills_Err

    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)

    Exit Sub

WriteRequestSkills_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteRequestSkills", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()

    On Error GoTo WriteRequestMiniStats_Err

    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)

    Exit Sub

WriteRequestMiniStats_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteRequestMiniStats", Erl)
    Resume Next
    
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()

    On Error GoTo WriteCommerceEnd_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)

    Exit Sub

WriteCommerceEnd_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCommerceEnd", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
    
    On Error GoTo WriteBankEnd_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)

    Exit Sub

WriteBankEnd_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteBankEnd", Erl)
    Resume Next
    
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Long)

    On Error GoTo WriteDrop_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)
    End With

    Exit Sub

WriteDrop_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteDrop", Erl)
    Resume Next
    
End Sub

''
' Writes the "DropDestroy" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDropDestroy(ByVal Slot As Byte, ByVal Amount As Long)
    
    On Error GoTo WriteDropDestroy_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.DropDestroy)
        
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)

    End With

    Exit Sub

WriteDropDestroy_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteDropDestroy", Erl)
    Resume Next
    
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)
    
    On Error GoTo WriteCastSpell_Err

    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(Slot)

    End With

    Exit Sub

WriteCastSpell_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCastSpell", Erl)
    Resume Next
    
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
    
    On Error GoTo WriteLeftClick_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)

    End With

    Exit Sub

WriteLeftClick_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteLeftClick", Erl)
    Resume Next
    
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
    
    On Error GoTo WriteDoubleClick_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoubleClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)

    End With
    
    Exit Sub

WriteDoubleClick_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteDoubleClick", Erl)
    Resume Next
    
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
    
    On Error GoTo WriteWork_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Work" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)

    End With
    
    Exit Sub

WriteWork_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteWork", Erl)
    Resume Next
    
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte)
    
    On Error GoTo WriteUseItem_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(Slot)

    End With
    
    Exit Sub

WriteUseItem_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteUseItem", Erl)
    Resume Next
    
End Sub


''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal item As Integer, cant As Integer)
    
    On Error GoTo WriteCraftBlacksmith_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftBlacksmith" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
        
    End With

    Exit Sub

WriteCraftBlacksmith_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCraftBlacksmith", Erl)
    Resume Next
    
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal item As Integer, ByVal cant As Integer)

    On Error GoTo WriteCraftCarpenter_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
        
    Exit Sub

WriteCraftCarpenter_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCraftCarpenter", Erl)
    Resume Next
    
End Sub

Public Sub WriteCraftalquimia(ByVal item As Integer, ByVal cant As Integer)

    On Error GoTo WriteCraftAlquimista_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Craftalquimia)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
    
    Exit Sub

WriteCraftAlquimista_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCraftAlquimista", Erl)
    Resume Next
    
End Sub

Public Sub WriteCraftSastre(ByVal item As Integer, ByVal cant As Integer)
    
    On Error GoTo WriteCraftSastre_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftSastre)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
    
    Exit Sub

WriteCraftSastre_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCraftSastre", Erl)
    Resume Next
    
End Sub

''
' Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    'Writes the "ShowGuildNews" message to the outgoing data buffer
    '***************************************************
 
    outgoingData.WriteByte (ClientPacketID.ShowGuildNews)

End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
    
    On Error GoTo WriteWorkLeftClick_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)

    End With

    Exit Sub

WriteWorkLeftClick_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteWorkLeftClick", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, _
                               ByVal Name As String, _
                               ByVal Site As String, _
                               ByRef Codex() As String)

    On Error GoTo WriteCreateNewGuild_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNewGuild" message to the outgoing data buffer
    '***************************************************
    Dim temp As String

    Dim i    As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Site)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)

    End With

    Exit Sub

WriteCreateNewGuild_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCreateNewGuild", Erl)
    Resume Next
    
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)
    
    On Error GoTo WriteEquipItem_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        Call .WriteByte(Slot)
    End With

    Exit Sub

WriteEquipItem_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteEquipItem", Erl)
    Resume Next
    
End Sub

Public Sub WriteEquiparSkin(ByVal Equipo As Byte, ByVal BackOrNext As Byte)
    
    On Error GoTo WriteEquiparSkin_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EquipItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquiparSkin)
        Call .WriteByte(Equipo)
        Call .WriteByte(BackOrNext)
    End With

    Exit Sub

WriteEquiparSkin_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteEquiparSkin", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
    
    On Error GoTo WriteChangeHeading_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeHeading" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)

    End With
    
    Exit Sub

WriteChangeHeading_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteChangeHeading", Erl)
    Resume Next
    
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
    
    On Error GoTo WriteModifySkills_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ModifySkills" message to the outgoing data buffer
    '***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i

    End With

    Exit Sub

WriteModifySkills_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteModifySkills", Erl)
    Resume Next
    
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
    
    On Error GoTo WriteTrain_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Train" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)

    End With

    Exit Sub

WriteTrain_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteTrain", Erl)
    Resume Next
    
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteCommerceBuy_Err
    
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceBuy" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)

    End With
    
    Exit Sub

WriteCommerceBuy_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCommerceBuy", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceSell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDeposit" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MoveSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)

    End With

End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal Slot As Byte)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    'Writes the "MoveBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)

    End With

End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ClanCodexUpdate" message to the outgoing data buffer
    '***************************************************
    Dim temp As String

    Dim i    As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)

    End With

End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAllianceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeaceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAlliancePropList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)

End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeacePropList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)

End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDeclareWar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNewWebsite" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(URL)

    End With

End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildKickMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildUpdateNews" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)

    End With

End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOpenElections" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)

End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestMembership" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)

    End With

End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Online" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)

End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/16/08
    'Writes the "Quit" message to the outgoing data buffer
    '***************************************************
    
    Call outgoingData.WriteByte(ClientPacketID.Quit)
    
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeave" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)

End Sub

Public Sub WriteLoginAccount()

    On Error GoTo errorhandler
    

    With outgoingData
        Call .WriteByte(ClientPacketID.ConnectAccount)
        
        Call .WriteASCIIString(Cuenta.UserAccount)
        Call .WriteASCIIString(SEncriptar(Cuenta.UserPassword))
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
 
    End With
    
    Exit Sub

errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteLoginAccount", Erl)
    Resume Next
    
End Sub

Public Sub WriteLoginNewAccount()
    
    On Error GoTo errorhandler
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewAccount)
        
        Call .WriteASCIIString(UCase$(LTrim(RTrim(Cuenta.UserAccount))))
        Call .WriteASCIIString(SEncriptar(Cuenta.UserPassword))
        Call .WriteASCIIString(SEncriptar(Cuenta.UserCode))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
    End With
    
    Exit Sub

errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteLoginNewAccount", Erl)
    Resume Next
    
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Meditate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)

End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Resucitate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)

End Sub
 
''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestStats" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)

End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceStart" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)

End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankStart" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)

End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Enlist" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Enlist)

End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Information" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)

End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Reward" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)

End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpTime" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UpTime)

End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal number As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CentinelReport" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(number)

    End With

End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnline" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)

End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest(ByVal Tipo As Byte, ByVal Soporte As String)

    On Error GoTo errorhandler
    
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
    Call outgoingData.WriteByte(Tipo)
    Call outgoingData.WriteASCIIString(Soporte)
    
    Exit Sub

errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteGMRequest", Erl)
    Resume Next
    
End Sub
 
''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)

    End With

End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub
 
''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Gamble" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)

    End With

End Sub
 
''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)

    End With

End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)

    End With

End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/21/2001
    'Writes the "GuildFundate" message to the outgoing data buffer
    '14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
    '03/21/2001: Pato - Deleted de clanType param.
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)

End Sub

''
' Writes the "GuildFundation" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "GuildFundation" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundation)
        
        Call .WriteByte(clanType)

    End With

End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberList" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildMemberList)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal Cantidad As Long, ByVal NroPorCiclo As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    'Writes the "InitCrafting" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        Call .WriteLong(Cantidad)
        
        Call .WriteInteger(NroPorCiclo)

    End With

End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GMMessage)
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowName" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.showName)

End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoNearby" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoNearby)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Comment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Comment)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerTime" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.serverTime)

End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Where" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Where)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreaturesInMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreaturesInMap)
        
        Call .WriteInteger(Map)

    End With

End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)

    End With

End Sub
 
''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()

    On Error GoTo WriteSOSShowList_err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)

    Exit Sub

WriteSOSShowList_err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteSOSShowList", Erl)
    Resume Next
    
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSRemove" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoToChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "invisible" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Invisible)

End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMPanel" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)

End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestUserList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)

End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Working" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)

End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Hiding" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)

End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(Time)

    End With

End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPC" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)

End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarnUser" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, _
                         ByVal EditOption As eEditOptions, _
                         ByVal arg1 As String, _
                         ByVal arg2 As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EditChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)

    End With

End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInventory" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharSkills" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReviveChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineGM" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)

End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    'Writes the "OnlineMap" message to the outgoing data buffer
    '26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        
        Call .WriteInteger(Map)

    End With

End Sub
''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Kick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Execute" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteBanChar(ByVal UserName As String, ByVal Banear As Byte)
    
    On Error GoTo BanChar
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(Banear)
        
    End With

    Exit Sub

BanChar:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteBanChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCFollow" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)

End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SummonChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub


' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetNPCInventory" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)

End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanWorld" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanWorld)

End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NickToIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.nickToIP)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef IP() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "IPToNick" message to the outgoing data buffer
    '***************************************************
    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        
        For i = LBound(IP()) To UBound(IP())
            Call .WriteByte(IP(i))
        Next i

    End With

End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnlineMembers" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)

    End With




End Sub


''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportDestroy" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)

End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal climas As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)
    Call outgoingData.WriteByte(climas)
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TalkAsNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)

End Sub
 
''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumbNoMore" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetTrigger" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        
        Call .WriteByte(Trigger)

    End With

End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    'Writes the "AskTrigger" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)

End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)

End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPReload" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)

End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildBan" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, _
                      ByRef IP() As Byte, _
                      ByVal Nick As String, _
                      ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanIP" message to the outgoing data buffer
    '***************************************************
    If byIp And UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then

            For i = LBound(IP()) To UBound(IP())
                Call .WriteByte(IP(i))
            Next i

        Else
            Call .WriteASCIIString(Nick)

        End If
        
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef IP() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanIP" message to the outgoing data buffer
    '***************************************************
    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        
        For i = LBound(IP()) To UBound(IP())
            Call .WriteByte(IP(i))
        Next i

    End With

End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal Count As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)
        Call .WriteInteger(Count)
    End With

End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyItems" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)

End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TileBlockedToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)

End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)

End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)

End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LastIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SystemMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        
        Call .WriteInteger(NPCIndex)

    End With

End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)

    End With

End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)

End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)

End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnOffServer" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)

End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterPassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)

    End With

End Sub


''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)

End Sub


''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildMessages" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveMap" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)

End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)

    End With

End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)

    End With

End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)

    End With

End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)

    End With

End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)

    End With

End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)

    End With

End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)

    End With

End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)

    End With

End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveChars" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)

End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanSOS" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)

End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KickAllChars" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)

End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadNPCs" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)

End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadServerIni" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)

End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadSpells" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)

End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadObjects" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)

End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping"   message to the outgoing data buffer
    '***************************************************
    'Prevent the timer from being cut

    Call outgoingData.WriteByte(ClientPacketID.Ping)
    CurrentUser.Ping = timeGetTime() And &H7FFFFFFF
    Call outgoingData.WriteLong(CurrentUser.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents

End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, _
                          ByRef sClave As String, _
                          ByRef sValor As String)

    '***************************************************
    'Author: Brian Chaia (BrianPr)
    'Last Modification: 21/06/2009
    'Writes the "SetIniVar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)

    End With

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************
    Dim sndData As String
    
    With outgoingData

        If .Length = 0 Then Exit Sub
        
        'Call AddtoRichTextBox("Salio paquete con peso de: " & sndData & " bytes", 0, 0, 0, 0, 0, 0, 8)
        'OutBytes = OutBytes + .Length
        'InBytes = 0
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        Call SendData(sndData)

    End With

End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados

        If Not frmMain.Socket1.IsWritable Then
            'Put data back in the bytequeue
            Call outgoingData.WriteASCIIStringFixed(sdData)
        
            Exit Sub

        End If
    
        If Not frmMain.Socket1.Connected Then Exit Sub
 
 
        Dim data() As Byte
    data = StrConv(sdData, vbFromUnicode)
    Security.NAC_E_Byte data, Security.Redundance
    sdData = StrConv(data, vbUnicode)
  
 
        Call frmMain.Socket1.Write(sdData, Len(sdData))

End Sub
Private Sub HandleEfectoCharParticula()

    On Error GoTo HandleCharParticle_Err
    
    If incomingData.Length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer, Particula As Integer
    Dim Life As Long
    Dim Remove As Boolean
    
    charindex = incomingData.ReadInteger()
    Particula = incomingData.ReadInteger()
    Life = incomingData.ReadSingle()
    Remove = incomingData.ReadBoolean()

    If Remove Then
        Call Char_Particle_Group_Remove(charindex, Particula)
        charlist(charindex).Particula = 0
    
    Else
        charlist(charindex).Particula = Particula
        charlist(charindex).ParticulaTime = Time
        Call SetCharacterParticle(Particula, charindex, Life)
    End If
 
    Exit Sub

HandleCharParticle_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCharParticle", Erl)
    Resume Next
    
End Sub

Public Sub HandleAddPj()

On Error GoTo HandleAddPj_Err
    
    If incomingData.Length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Cuenta.UserAccount = buffer.ReadASCIIString()
    NumberOfCharacters = buffer.ReadByte()
    
   
    'Cambiamos al modo cuenta
    frmCharList.Show

    
    If NumberOfCharacters > 0 Then
    
        ReDim cPJ(1 To NumberOfCharacters) As PjCuenta
        
        Dim loopc As Long
        
        For loopc = 1 To NumberOfCharacters
        
            With cPJ(loopc)
                .Nombre = buffer.ReadASCIIString()
                .Head = buffer.ReadInteger()
                .body = buffer.ReadInteger()
                .Helmet = buffer.ReadInteger()
                .Weapon = buffer.ReadInteger()
                .Shield = buffer.ReadInteger()
                .Nivel = buffer.ReadByte()
                .Clase = buffer.ReadByte()
                .Mapa = buffer.ReadInteger()
                .color = buffer.ReadByte()
                .GameMaster = buffer.ReadBoolean()
            End With
            
        Next loopc
        
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    For loopc = 1 To NumberOfCharacters
        Call DrawPJ(loopc)
    Next loopc
 
HandleAddPj_Err:

    Dim error As Long
    error = Err.number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then Err.Raise error
End Sub

 

Public Sub WriteSwapObjects(ByVal ObjSlot1 As Byte, ByVal ObjSlot2 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.SwapObjects)
   
        Call .WriteByte(ObjSlot1)
        Call .WriteByte(ObjSlot2)
    
    End With
End Sub
Public Sub WriteResponderGm(ByVal UserName As String, ByVal MensajeUser As String, ByVal TODOS As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResponderGM)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(MensajeUser)
        Call .WriteASCIIString(TODOS)
    End With
End Sub
 
Public Sub WriteRetirarFaccion()

    On Error GoTo WriteRetirarFaccion_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RetirarFaccion)

    Exit Sub

WriteRetirarFaccion_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.RetirarFaccion", Erl)
    Resume Next
End Sub
Public Sub WriteRegresarHogar()

    On Error GoTo WriteRegresarHogar_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RegresarHogar)

    Exit Sub

WriteRegresarHogar_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteRegresarHogar", Erl)
    Resume Next
End Sub

Public Sub WriteParticulaUsuario(ByVal UserName As String, ByVal Particula As Integer)
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ParticulaUsuario)
        Call .WriteASCIIString(UserName)
        Call .WriteInteger(Particula)

    End With
    
End Sub
Public Sub WriteProcesosLogin()

    On Error GoTo ErorrHandler_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ProcesosLogin)
        Call .WriteASCIIString(Cuenta.UserAccount)
        Call .WriteASCIIString(SEncriptar(Cuenta.UserCode))
        Call .WriteASCIIString(SEncriptar(Cuenta.UserPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteByte(Cuenta.EsChange)
    End With
    
    Exit Sub
ErorrHandler_Err:
     Call RegistrarError(Err.number, Err.Description, "Protocol.WriteProcesosLogin", Erl)
     Resume Next
        
End Sub

Public Sub WriteTransferGold(ByVal UserName As String, ByVal Amount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.TransferGOLD)
        Call .WriteASCIIString(UserName)
        Call .WriteLong(Amount)
    End With
 
End Sub

Private Sub HandleEfectoTerrenoParticula()


 On Error GoTo HandleEfectoTerrenoParticula_Err
 
    If incomingData.Length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
    'Remove packet ID
1    Call incomingData.ReadByte

3    Dim ParticulaFx As Integer
4    Dim X As Byte
5    Dim Y As Byte
6    Dim Time As Long
7    Dim Borrar As Boolean

9    ParticulaFx = incomingData.ReadInteger()
10   X = incomingData.ReadByte()
11   Y = incomingData.ReadByte()
12   Time = incomingData.ReadLong()
 
13   If Time = 1 Then Time = -1

14   If Time = 0 Then Borrar = True
 
 
15   If Borrar Then
16     Particle_Group_Remove (MapData(X, Y).particle_group)
17   Else
       MapData(X, Y).particle_group = 0
       SetMapParticle ParticulaFx, X, Y, Time
24   End If
    
     Exit Sub
     
HandleEfectoTerrenoParticula_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleEfectoTerrenoParticula", Erl)
    Resume Next
End Sub
Private Sub HandleEfectoTerrenoFX()

    On Error GoTo HandleFXTerreno_Err
    
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim fX As Integer, Loops As Integer
    Dim X As Byte, Y As Byte
    
    fX = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    Loops = incomingData.ReadInteger()
    
    Call SetFXMAP(fX, X, Y, Loops)

    Exit Sub

HandleFXTerreno_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleFXTerreno", Erl)
    Resume Next
End Sub

Private Sub HandleCorreoList()

    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
   
    'Remove packet ID
    Call buffer.ReadByte
   
    Dim NumCorreos As Byte
    Dim i As Long

    NumCorreos = buffer.ReadByte()
   
    frmCorreo.lstMsg.Clear
   
    For i = 1 To NumCorreos
        Correos(i).De = buffer.ReadASCIIString()
        Correos(i).Mensaje = buffer.ReadASCIIString()
        
        Correos(i).Leido = buffer.ReadByte()
        
        Correos(i).Cantidad = buffer.ReadInteger()
        
        Correos(i).OBJIndex = buffer.ReadInteger()
            
        Correos(i).Nombre = General_Locale_Obj(Correos(i).OBJIndex, 0)
        Correos(i).GrhIndex = General_Locale_Obj(Correos(i).OBJIndex, 3)
         
        
        If Correos(i).De <> "" Then
        
            If Correos(i).Leido = 0 Then
                Call frmCorreo.lstMsg.AddItem(Correos(i).De & " [" & Locale_GUI_Frase(488) & "]")
            Else
                Call frmCorreo.lstMsg.AddItem(Correos(i).De)
            End If
        End If
        
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmCorreo.ActualizarCorreo
 
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub

Public Sub WritePacketsCorreo(ByVal Index As Byte, ByVal SlotCorreo As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.Packets_Correo)
        Call .WriteByte(Index)
        Call .WriteByte(SlotCorreo)
    End With
End Sub
 
Public Sub WriteEnviarCorreo(ByVal destinatario As String, ByVal Mensaje As String, ByVal ObjetoIndex As Integer, ByVal objetoAmount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.EnviarCorreo)
        Call .WriteASCIIString(destinatario)
        Call .WriteASCIIString(Mensaje)
        Call .WriteInteger(ObjetoIndex)
        Call .WriteInteger(objetoAmount)
    End With
End Sub

Public Sub WriteDonador(ByVal UserName As String)
    
    On Error GoTo WriteDonador_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.donador)
        Call .WriteASCIIString(UserName)
    End With
    
    Exit Sub

WriteDonador_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteDonador", Erl)
    Resume Next
    
    
End Sub
  Public Sub WriteEventoOro(ByVal multi As Byte, ByVal tiempo As Byte)
With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EventoOro)
        Call .WriteByte(multi)
        Call .WriteByte(tiempo)
End With
End Sub

 Public Sub writeEventExperiencia(ByVal multi As Byte, ByVal tiempo As Byte)
With outgoingData
 
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EventoExperiencia)
        Call .WriteByte(multi)
        Call .WriteByte(tiempo)
End With
End Sub
Public Sub HandleCharStatus()

    On Error GoTo HandleCharStatus_Err
    
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    Dim priv As Byte
    
    charindex = incomingData.ReadInteger
    priv = incomingData.ReadByte
    
    charlist(charindex).priv = priv
    Call ColorNombresPriv(charindex, priv)
    
    Exit Sub

HandleCharStatus_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCharStatus", Erl)
    Resume Next
End Sub
 Public Sub WriteSeleccionarHogar(Optional ByVal Mando As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.SeleccionarHogar)
       Call .WriteByte(Mando)
    End With
End Sub


Public Sub WriteCasamiento(ByVal UserName As String, ByVal Modo As Byte)
    
    On Error GoTo WriteCasamiento_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.Casamiento)
        Call .WriteASCIIString(Replace(UserName, " ", "+"))
        Call .WriteByte(Modo)
    End With
    
    Exit Sub

WriteCasamiento_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteCasamiento", Erl)
    Resume Next
End Sub

Public Sub writeDivorciar()
    With outgoingData
        Call .WriteByte(ClientPacketID.divorciar)
    End With
End Sub
Public Sub WriteHayEventos()
    With outgoingData
        Call .WriteByte(ClientPacketID.HayEventos)
    End With
End Sub
 
Public Sub WriteRPremios(ByVal Index As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.RPremios)
        Call .WriteInteger(Index)
    End With
End Sub
 
Public Sub PedirPremios()
    Call outgoingData.WriteByte(ClientPacketID.PidePremios)
End Sub
 
Public Sub HandlePremios()
 
     If incomingData.Length < 11 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
On Error GoTo ErrHandler
    
Dim buffer As New clsByteQueue
Call buffer.CopyBuffer(incomingData)
 
Dim CantPremios As Byte
Dim i As Byte
 
Call buffer.ReadByte
CantPremios = buffer.ReadInteger
CurrentUser.Creditos = buffer.ReadLong
For i = 1 To CantPremios
    With PremiosInv(i)
        .Name = buffer.ReadASCIIString
        .Puntos = buffer.ReadInteger
    End With
    
    If PremiosInv(i).Name <> "" Then
        frmShop.List1.AddItem PremiosInv(i).Name
    Else
        frmShop.List1.AddItem Locale_GUI_Frase(269)
    End If
Next i

frmShop.lstInv.Clear

For i = 1 To MAX_INVENTORY_SLOTS
 If Inventario.ItemName(i) <> "" Then
    frmShop.lstInv.AddItem Inventario.ItemName(i)
 Else
    frmShop.lstInv.AddItem Locale_GUI_Frase(269)
 End If
Next i
 
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub
 
Public Sub WriteDARPUN(ByVal UserName As String, ByVal DAPUN As Long)
With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.DARPUN)
        Call .WriteASCIIString(UserName)
        Call .WriteLong(DAPUN)
End With
End Sub
Private Sub HandleMensajeSigno()

    On Error GoTo HandleMensajeSigno_Err

1    If incomingData.Length < 2 Then
2        Err.Raise incomingData.NotEnoughDataErrCode
3        Exit Sub
4    End If
    
    'Remove packet ID
5    Call incomingData.ReadByte

6    Dim TieneMensaje As Byte

7    TieneMensaje = incomingData.ReadByte

8    If TieneMensaje Then
9       frmMain.nuevocorreo.Visible = True
10   Else
11      frmMain.nuevocorreo.Visible = False
12   End If
        
13   Exit Sub

HandleMensajeSigno_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleMensajeSigno", Erl)
    Resume Next
    
End Sub

Public Sub writeCloseGuild()
With outgoingData
Call .WriteByte(ClientPacketID.CloseGuild)
End With
End Sub
 
  Public Sub WriteCuentaRegresiva(ByVal Second As Byte, ByVal Lugar As Byte)
 
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call outgoingData.WriteByte(eGMCommands.CuentaRegresiva)
        Call .WriteByte(Second)
        Call .WriteByte(Lugar)
    End With
End Sub
 

Private Sub HandleMarcamosSkin()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    On Error GoTo HandleMarcamosSkin_Err
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
    'Remove packet ID
    Call incomingData.ReadByte
   
    Dim Index As Byte
    Index = incomingData.ReadByte()
    
    Select Case Index
    Case 0
    frmSkins.Image15.Visible = False
    
    Case 1
    frmSkins.Image15.Visible = True
        
    Case 2 'armaduras equip
    frmSkins.Image11.Visible = True
    Case 3 'armaduras desek
    frmSkins.Image11.Visible = False

    Case 4
    frmSkins.Image14.Visible = True
     
    
    Case 5
    frmSkins.Image14.Visible = False
    
    
    Case 6
    frmSkins.Image13.Visible = True

    Case 7
    frmSkins.Image13.Visible = False
    
    Case 8
    
    frmSkins.Image12.Visible = True
    
    Case 9
    frmSkins.Image12.Visible = False
    
    End Select
    
    Exit Sub

HandleMarcamosSkin_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleMarcamosSkin", Erl)
    Resume Next
    
End Sub
 
Public Sub WriteAddAmigo(ByVal UserName As String, ByVal Index As Byte)
'***************************************************
'Author: Bateman
'***************************************************
  With outgoingData
  Call .WriteByte(ClientPacketID.AddAmigos)
  Call .WriteASCIIString(UserName)
  Call .WriteByte(Index)
  End With
End Sub
Public Sub WriteDelAmigo(ByVal Index As String)
'***************************************************
'Author: Bateman
'***************************************************
  With outgoingData
  Call .WriteByte(ClientPacketID.DelAmigos)
  Call .WriteASCIIString(Index)
  End With
End Sub
Public Sub WriteOnAmigoandMapa()
'***************************************************
'Author: Bateman
'***************************************************
  With outgoingData
  Call .WriteByte(ClientPacketID.OnAmigos)
  End With
End Sub
Public Sub WriteMsgAmigo(ByVal msg As String)
'***************************************************
'Author: Bateman
'***************************************************
  With outgoingData
  Call .WriteByte(ClientPacketID.MsgAmigos)
  Call .WriteASCIIString(msg)
  End With
End Sub
 
 
Private Sub Handlemostrarubicacion()

    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
     On Error GoTo ErrHandler


Dim miBuffer As New clsByteQueue
 
Call miBuffer.CopyBuffer(incomingData)
 
Call miBuffer.ReadByte
 
Dim Nombre As String
Dim NumAmigo As Byte
Dim Map As Integer
Dim PosX As Byte
Dim PosY As Byte
Nombre = miBuffer.ReadASCIIString
NumAmigo = miBuffer.ReadByte
Map = miBuffer.ReadInteger
PosX = miBuffer.ReadByte
PosY = miBuffer.ReadByte
Nombre = UCase$(Nombre)
        Select Case Map
        Case 0
        frmMain.Shape2(NumAmigo).Visible = False
        frmMain.Label1(NumAmigo).Visible = False
        frmMap.Shape1(NumAmigo).Visible = False
        frmMap.Label1(NumAmigo).Visible = False
        
        Case CurrentUser.UserMap
        'Lo visualizamos en mi minimapa
        frmMain.Shape2(NumAmigo).Visible = True
        frmMain.Shape2(NumAmigo).Left = PosX
        frmMain.Shape2(NumAmigo).Top = PosY
        
        frmMain.Label1(NumAmigo).Caption = Nombre
        frmMain.Label1(NumAmigo).Left = frmMain.Shape2(NumAmigo).Left
        frmMain.Label1(NumAmigo).Top = frmMain.Shape2(NumAmigo).Top

        'Y en el mapa del mundo
        frmMap.Shape1(NumAmigo).Visible = True
        Call frmMap.SetMapPoint2(NumAmigo, Map)
        
        frmMap.Label1(NumAmigo).Caption = Nombre
        frmMap.Label1(NumAmigo).Left = frmMap.Shape1(NumAmigo).Left
        frmMap.Label1(NumAmigo).Top = frmMap.Shape1(NumAmigo).Top + 10

        Case Else
        
        frmMain.Shape2(NumAmigo).Visible = False
        frmMain.Label1(NumAmigo).Visible = False
        
        'Y en el mapa del mundo
        frmMap.Shape1(NumAmigo).Visible = True
        Call frmMap.SetMapPoint2(NumAmigo, Map)
        
        frmMap.Label1(NumAmigo).Caption = Nombre
        frmMap.Label1(NumAmigo).Left = frmMap.Shape1(NumAmigo).Left
        frmMap.Label1(NumAmigo).Top = frmMap.Shape1(NumAmigo).Top + 10
        
        End Select
        frmMain.Minimap.Refresh
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(miBuffer)
    
ErrHandler:

    'If Err.number <> 0 And Err.number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set miBuffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub


 

Public Sub HandleCargarSkin()
    On Error GoTo HandleCargarSkin_Err
    
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
        Dim Head   As Integer, _
        body   As Integer, _
        Casco  As Byte, _
        Arma   As Byte, _
        Escudo As Byte
    
    With incomingData
        Call .ReadByte
        Head = .ReadInteger()
        body = .ReadInteger()
        Casco = .ReadByte()
        Arma = .ReadByte()
        Escudo = .ReadByte()
    End With
   
    RSkin.Head = Head
    RSkin.body = body
    RSkin.Casco = Casco
    RSkin.Weapon = Arma
    RSkin.Shield = Escudo
    
    Call DrawSkinPJ
    
    Exit Sub

HandleCargarSkin_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleCargarSkin", Erl)
    Resume Next
    
End Sub
Public Sub HandleCharMsgStatus()

    '***************************************************
    'Author: Mermas
    'Last Modification: 21/7/21
    '
    '***************************************************
    
    'Check packet is complete
    If incomingData.Length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim tempStr As String, Pareja As String, Desc As String, st As String
    Dim charindex As Integer, St1 As Integer
    Dim lngPorcVida As Long
    Dim btStatus As Byte, St2 As Byte, btClase As Byte, btRaza As Byte, btNivel As Byte, donador As Byte, rangoFaccion As Byte
    Dim btRed As Byte, btGreen As Byte, btBlue As Byte
    
    charindex = buffer.ReadInteger()
    btStatus = buffer.ReadByte()
    lngPorcVida = buffer.ReadLong()
    St1 = buffer.ReadInteger()
    St2 = buffer.ReadByte()
    btClase = buffer.ReadByte()
    btNivel = buffer.ReadInteger()
    btRaza = buffer.ReadByte()
    donador = buffer.ReadByte()
    rangoFaccion = buffer.ReadByte()
    Pareja = buffer.ReadASCIIString()
    Desc = buffer.ReadASCIIString()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    st = Generate_Char_Status(lngPorcVida, BoolToInteger((St2 And StatEx.Paralizado)), BoolToInteger((St2 And StatEx.Inmovilizado)), _
        BoolToInteger((St1 And Stat.Incinerado)), BoolToInteger((St1 And Stat.Envenenado)), BoolToInteger((St1 And Stat.Comerciand)), BoolToInteger((St1 And Stat.Trabajando)), _
        BoolToInteger((St1 And Stat.Combatiendo)), BoolToInteger((St1 And Stat.Ciego)), BoolToInteger((St1 And Stat.Inactivo)), _
        BoolToInteger((St1 And Stat.Resucitando)), BoolToInteger((St1 And Stat.Saliendo)))
    
    tempStr = charlist(charindex).Nombre
    
    If donador Then tempStr = tempStr & " " & Locale_Facc_Frase(32)
    
    tempStr = tempStr & " (" & ListaClases(btClase) & " " & ListaRazas(btRaza) & " " & Locale_GUI_Frase(158) & " "
    
    If btNivel = 255 Then
        tempStr = tempStr & "??"
    Else
        tempStr = tempStr & btNivel
    End If

    tempStr = tempStr & "|" & st & ")"

    Select Case btStatus
    
        Case 1 'Renegado
            tempStr = tempStr & " <" & Locale_GUI_Frase(154) & ">"
            btRed = 114
            btGreen = 115
            btBlue = 108
        Case 2 'Imperial
            tempStr = tempStr & " <" & Locale_GUI_Frase(152) & ">"
            btRed = 32
            btGreen = 81
            btBlue = 251
        Case 3 'Republicano
            tempStr = tempStr & " <" & Locale_GUI_Frase(153) & ">"
            btRed = 204
            btGreen = 107
            btBlue = 0
        Case 5 'Caos
            tempStr = tempStr & " <" & Locale_GUI_Frase(150) & "> <" & Locale_Facc_Frase(rangoFaccion + 10) & ">"
            btRed = 196
            btGreen = 0
            btBlue = 15
        Case 6 'Imperial
            tempStr = tempStr & " <" & Locale_GUI_Frase(148) & "> <" & Locale_Facc_Frase(rangoFaccion) & ">"
            btRed = 32
            btGreen = 81
            btBlue = 251
        Case 7 'Republicano
            tempStr = tempStr & " <" & Locale_GUI_Frase(149) & "> <" & Locale_Facc_Frase(rangoFaccion + 20) & ">"
            btRed = 204
            btGreen = 107
            btBlue = 0
        Case 10
            tempStr = tempStr & " <" & Locale_Facc_Frase(29) & ">"
            btRed = 2
            btGreen = 162
            btBlue = 38
        Case 11
            tempStr = tempStr & " <" & Locale_Facc_Frase(29) & ">"
            btRed = 2
            btGreen = 162
            btBlue = 38
        Case 12
            tempStr = tempStr & " <" & Locale_Facc_Frase(30) & ">"
            btRed = 2
            btGreen = 162
            btBlue = 38
        Case 13
            tempStr = tempStr & " <" & Locale_Facc_Frase(31) & ">"
            btRed = 2
            btGreen = 162
            btBlue = 38
    End Select
    
           
    If charlist(charindex).OffSetClan > 0 Then tempStr = tempStr & charlist(charindex).Clan
    
    If Len(Pareja) > 0 Then tempStr = tempStr & " <" & Locale_GUI_Frase(468) & " " & Pareja & ">"
        
    If Len(Desc) > 0 Then tempStr = tempStr & " - " & Desc
    
    If NickIgnorado(charlist(charindex).Nombre) Then tempStr = tempStr & "(" & Locale_GUI_Frase(557) & ")"
 
    AddtoRichTextBox tempStr, btRed, btGreen, btBlue, 1, 0, 0
        
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error


End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserDescansar = Not CurrentUser.UserDescansar
End Sub


''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub
 Sub HandleCharMsgStatusNPC()

    On Error GoTo HandleCharMsgStatusNPC_Err
    
    'Check packet is complete
    If incomingData.Length < 14 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    'Get data and update form
    Dim charindex As Integer
    Dim btStatus As Byte, PuedeVerVida As Byte, St1 As Byte, btNivel As Byte, MaestroUser As Byte, Owner As Byte
    Dim lngPorcVida As Long
    Dim tempStr As String, st As String
    
    
    charindex = incomingData.ReadInteger
    btStatus = incomingData.ReadByte
    PuedeVerVida = incomingData.ReadByte
    lngPorcVida = incomingData.ReadLong
    St1 = incomingData.ReadByte
    btNivel = incomingData.ReadInteger
    MaestroUser = incomingData.ReadByte
    Owner = incomingData.ReadByte
    
    tempStr = General_Locale_NPCs(charindex, 0)
      
    If btNivel = 255 Then
        tempStr = tempStr & " " & Locale_GUI_Frase(158) & " " & "??"
    Else
        tempStr = tempStr & " " & Locale_GUI_Frase(158) & " " & btNivel
    End If
    
    If PuedeVerVida Then
        st = Generate_Char_StatusNPCs(CLng(((lngPorcVida / 100) / (General_Locale_NPCs(charindex, 6) / 100)) * 100), BoolToByte((St1 And StatEx.Paralizado)), BoolToByte((St1 And StatEx.Inmovilizado)), 0)
        tempStr = tempStr & " (" & LTrim(st) & ")" & " (" & lngPorcVida & "/" & General_Locale_NPCs(charindex, 6) & ")" & " "
    Else
        st = Generate_Char_StatusNPCs(lngPorcVida, BoolToByte((St1 And StatEx.Paralizado)), BoolToByte((St1 And StatEx.Inmovilizado)), PuedeVerVida)
        tempStr = tempStr & " (" & LTrim(st) & ")"
    End If
    
    AddtoRichTextBox tempStr, 0, 0, 0, 0, 0, 0, 4
    
    Exit Sub
    
HandleCharMsgStatusNPC_Err:
    Call RegistrarError(Err.number, Err.Description, "HandleCharMsgStatusNPC", Erl)
    Resume Next

End Sub


Private Sub HandleLocaleMsg()

    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim id As Integer
    Dim Modo As Integer
    Dim fuente As Byte
    
    id = buffer.ReadInteger()
    chat = buffer.ReadASCIIString()
    Modo = CInt(buffer.ReadByte()) 'Siempre que sean numerosos cases por más que no superen os byte, etc, el long o integer procesa más rápido
    fuente = buffer.ReadByte()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Select Case Modo
    
        Case 0 'Usamos init SMG con su fuente
            
            chat = Locale_Parse_ServidorMensaje(id, chat)
            If fuente = 0 Then fuente = IIf(General_Locale_SMG(id, 1) > 0 And General_Locale_SMG(id, 1) < 26, General_Locale_SMG(id, 1), 12)
            Call AddtoRichTextBox(chat, 0, 0, 0, 0, 0, 0, fuente)
    
        Case 1 'Usamos init GUI con fuente_info
        
            Call AddtoRichTextBox(Locale_GUI_Frase(id), 0, 0, 0, 0, 0, 0, IIf(fuente > 25 Or fuente < 1, 12, fuente))
    
        Case 2 'Mensajes ataque criatura
    
            Select Case id
            
                Case bCabeza
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & Locale_GUI_Frase(367) & " " & chat, 0, 0, 0, 0, 0, 0, 2)
                    
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & Locale_GUI_Frase(368) & " " & chat, 0, 0, 0, 0, 0, 0, 2)
                
                Case bBrazoDerecho
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & Locale_GUI_Frase(369) & " " & chat, 0, 0, 0, 0, 0, 0, 2)
                    
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & Locale_GUI_Frase(370) & " " & chat, 0, 0, 0, 0, 0, 0, 2)
                
                Case bPiernaDerecha
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & Locale_GUI_Frase(371) & " " & chat, 0, 0, 0, 0, 0, 0, 2)
                
                Case bTorso
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & Locale_GUI_Frase(372) & " " & chat, 0, 0, 0, 0, 0, 0, 2)
                
                Case Else
                    Call AddtoRichTextBox(Locale_GUI_Frase(497) & " " & Locale_GUI_Frase(374), 0, 0, 0, 0, 0, 0, 2)
            
            End Select
        
    End Select
    
ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub


''
' Handles the Logged message.

Private Sub HandleLoggedSuccessful()

    On Error GoTo HandleLoggedSuccessful_Err
    

    'Remove packet ID
1    Call incomingData.ReadByte
    
3    frmIniciando.Show
   
4    Exit Sub

HandleLoggedSuccessful_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleLoggedSuccessful", Erl)
    Resume Next
End Sub
Public Sub HandleAbrirFormularios()

    '***************************************************
    'Author: Mermas
    'Last Modification: 07/08/21
    'Borramos mas de 15 paquetes y centramos abrir formularios en un Case
    '***************************************************
    
    On Error GoTo HandleAbrirFormulario_Err
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim Formulario As Long

    'Remove packet ID
    Call incomingData.ReadByte
    
    Formulario = CLng(incomingData.ReadByte) 'Pasamos a Long ya que es mucho mas optimo
 
    Select Case Formulario
    
        Case 1 'Show Account
    
            Call FormParser.Parse_Form(frmCharList)
         
            frmCharList.lblAccData(0).Caption = Cuenta.UserAccount
                    
            If Not frmCharList.Visible Then
                frmCharList.Show
            End If
            
            If frmCrearCuenta.Visible = True Then 'Si es primera vez que creó
                Call FormParser.Parse_Form(frmCrearCuenta)
                Unload frmCrearCuenta
            End If
         
            If Not frmConnect Is Nothing Then
            
                Call FormParser.Parse_Form(frmConnect)
                
                If Not RecordarCuentaIni Then
                    frmConnect.txtNombre.Text = ""
                End If
        
                frmConnect.txtPasswd.Text = ""
                frmConnect.Visible = False
        
            End If
            
        Case 2
            frmCarp.Show , frmMain
        
        Case 3
            frmSastre.Show , frmMain
        
        Case 4
            frmDruida.Show , frmMain
        
        Case 5
            frmShop.Show vbModeless, frmMain
        
        Case 6
            If Not frmCorreo.Visible Then
                frmCorreo.Show vbModeless, frmMain
            End If
        
        Case 7
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
    
        Case 8
            frmGMPanel.Show vbModeless, frmMain
        
        Case 9
            frmDonador.Show , frmMain
        
        Case 10
            frmSkins.Show , frmMain
        
        Case 11
            frmHerrero.Show , frmMain
            
    End Select
    
    Exit Sub
    
HandleAbrirFormulario_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleAbrirFormulario & Dato: " & Formulario, Erl)
    Resume Next
End Sub




Public Sub WriteAbrirForms(ByVal Formulario As Byte)

    On Error GoTo errorhandler
    
    Call outgoingData.WriteByte(ClientPacketID.AbrirForms)
    Call outgoingData.WriteByte(Formulario)
    Exit Sub

errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteAbrirForms", Erl)
    Resume Next
End Sub


''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlotUser()

    On Error GoTo errorhandler
    
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Formulario As Long

    'Remove packet ID
    Call incomingData.ReadByte

    Dim Slot As Byte
    Dim Accion As Byte
    Dim Valor As Integer
    
    Slot = incomingData.ReadByte
    Accion = incomingData.ReadByte
    Valor = incomingData.ReadInteger
    
    Call Inventario.SetItemUser(Slot, Accion, Valor)
    
    Exit Sub
    
errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleChangeInventorySlotUser", Erl)
    Resume Next
End Sub


Private Sub HandleAuraToChar()
    
    On Error GoTo HandleAuraToChar_Err

    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    Dim ParticulaIndex As Byte
    Dim Tipo As Integer
     
    charindex = incomingData.ReadInteger
    ParticulaIndex = incomingData.ReadByte
    Tipo = CInt(incomingData.ReadByte)
    
    Select Case Tipo
        Case 1
            charlist(charindex).Arma_Aura = ParticulaIndex
        Case 2
            charlist(charindex).Body_Aura = ParticulaIndex
        Case 3
            charlist(charindex).Escudo_Aura = ParticulaIndex
        Case 4
            charlist(charindex).Head_Aura = ParticulaIndex
        Case 5
            charlist(charindex).Otra_Aura = ParticulaIndex
        Case 6
            charlist(charindex).Anillo_Aura = ParticulaIndex
    End Select
    
    Exit Sub

HandleAuraToChar_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleAuraToChar", Erl)
    Resume Next
    
End Sub



Private Sub HandleUpdateSed()
    
    On Error GoTo errorhandler

    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , , , , , , , , , , , , CInt(incomingData.ReadByte))
    
    Exit Sub

errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateSed", Erl)
    Resume Next
    
End Sub



Private Sub HandleUpdateHambre()
    
    On Error GoTo errorhandler
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call ClientTCP.ActualizarEst(, , , , , , , , , , , , CInt(incomingData.ReadByte))
    
    Exit Sub

errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.HandleUpdateHambre", Erl)
    Resume Next
    
End Sub
Public Sub WriteDesconectarCuenta(ByVal Nombre As String)
    
    On Error GoTo DesconectarCuenta_Err

    With outgoingData
        Call .WriteByte(ClientPacketID.DesconectarCuenta)
        Call .WriteASCIIString(Nombre)
    End With

    
    Exit Sub

DesconectarCuenta_Err:
    Call RegistrarError(Err.number, Err.Description, "Protocol.WriteDesconectarCuenta", Erl)
    Resume Next
    
End Sub

Private Sub HandleEjecutarAccion()
    
    On Error GoTo ErrHandler
    
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
          
    'Remove packet ID
    Call buffer.ReadByte
    
    'Funcion para realizar acciones independientes desde el cliente

    Dim Accion As Integer
    Dim Extra As String

    Accion = CInt(buffer.ReadByte())
    Extra = buffer.ReadASCIIString()

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Select Case Accion
    
        Case 1 'Iniciar / cerrar salida.
            If CurrentUser.TiempoSalida = True Then
                CurrentUser.TiempoSalida = False
            Else
                CurrentUser.TiempoSalida = True
            End If
            
        Case 2 'Creditos

        Case 3 'Cartel casamiento
            frmPregunta.SetAccion 8, Extra
            If frmMain.Visible Then frmPregunta.Show , frmMain
            
        Case 4 'Actualizamos info de updates
            If Len(Extra) > 0 Then
                If Len(frmHlp.txtMsg.Text) = 0 Then
                    frmHlp.txtMsg.Text = Extra
                Else
                    frmHlp.txtMsg.Text = frmHlp.txtMsg.Text & vbCrLf & Extra
                End If
            End If
            
        Case 5 'Activamos boton :p Recuperar cuenta
        
            If frmRecuperarCuenta.Visible Then
                frmRecuperarCuenta.cmdAceptar.Caption = Locale_GUI_Frase(643) 'Recuperar
                frmRecuperarCuenta.cmdAceptar.Enabled = True
                
                If Extra = "1" Then
                    frmRecuperarCuenta.txtNombre.Text = ""
                    frmRecuperarCuenta.txtUserCode.Text = ""
                    frmRecuperarCuenta.txtPassword.Text = ""
                    If frmMain.Socket1.State <> sckClosed Then frmMain.Socket1.Disconnect
                    Unload frmRecuperarCuenta
                    
                End If
                
            ElseIf frmCambiarContraseña.Visible Then
                frmCambiarContraseña.cmdAceptar.Caption = Locale_GUI_Frase(640) 'Cambiar contraseña
                frmCambiarContraseña.cmdAceptar.Enabled = True
                
                If Extra = "1" Then
                    frmCambiarContraseña.txtPassword.Text = ""
                    frmCambiarContraseña.txtNewPassword.Text = ""
                    Unload frmCambiarContraseña
                    
                End If
                
            ElseIf frmCharList.Visible Then
                Call FormParser.Parse_Form(frmPregunta)
            End If
            
        Case 6 'Reload charlist
            Call frmCharList.LimpiarPersonajes
               
    End Select

ErrHandler:

    Dim error As Long

    error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then Err.Raise error

End Sub
 
