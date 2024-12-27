Attribute VB_Name = "Mod_General"
Option Explicit
Public IntervaloCaminar As Long

Public Moviendose                               As Boolean

'To load icons
Private Declare Function GetSystemMetrics Lib "user32" ( _
      ByVal nIndex As Long _
   ) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
   
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" ( _
   ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Const GW_OWNER = 4
 
'Mermas, Mouse configurable
    'Set mouse speed
    Private Declare Function SystemParametersInfo Lib "user32" Alias _
        "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
        ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
     
    Private Const SPI_SETMOUSESPEED = 113
    Private Const SPI_GETMOUSESPEED = 112
    ''

     

'***********************************************
'***********************************************

 

'Fin ConfiguracionNueva Mermas
Private keysMovementPressedQueue As clsArrayList

Public bFogata As Boolean
Private Type tMapSize

    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer

End Type
 
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

Private Type tColor

    r As Long
    g As Long
    b As Long
            
End Type
 
Private Type tDatosLuces

    X As Integer
    Y As Integer
    color As Long
    'extra As Byte
    range As Byte

End Type
 
Private Type tDatosParticulas

    X As Integer
    Y As Integer
    Particula As Long

End Type
 
Private Type tDatosNPC

    X As Integer
    Y As Integer
    NPCIndex As Integer

End Type
 
Private Type tDatosObjs

    X As Integer
    Y As Integer
    OBJIndex As Integer
    ObjAmmount As Integer

End Type
 
Private Type tDatosTE

    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer

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
    Base_Light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32

End Type
 
Private MapSize      As tMapSize

Private MapDat       As tMapDat

 



Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Public Function GetRawName(ByRef sName As String) As String
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 13/01/2010
    'Last Modified By: -
    'Returns the char name without the clan name (if it has it).
    '***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim$(Left$(sName, Pos - 1))
    Else
        GetRawName = sName

    End If

End Function

Sub AddtoRichTextBox(Text As String, _
                     Optional ByVal red As Integer = -1, _
                     Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, _
                     Optional ByVal bold As Boolean, _
                     Optional ByVal italic As Boolean, _
                     Optional ByVal bCrLf As Boolean, Optional ByVal FontTypeIndex As Byte = 0)

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '******************************************r
    Dim bUrl As Boolean
    With frmMain.RecTxt

        .SelFontName = "Tahoma"
        .SelFontSize = 8
        
    If FontTypeIndex <= 0 Then

            bUrl = True
            EnableUrlDetect
        If (Len(.Text)) > 20000 Then .Text = vbNullString
        .SelStart = Len(frmMain.RecTxt.Text)
        .SelLength = 0
        .SelBold = IIf(bold, True, False)
        .SelItalic = IIf(italic, True, False)
            
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        
    Else

            If (Len(.Text)) > 20000 Then .Text = vbNullString
            
            If FontTypeIndex = FONTTYPE_SERVER Then Text = "Servidor> " & Text
            
            bUrl = (FontTypeIndex = FONTTYPE_SERVER Or FontTypeIndex = FONTTYPE_TALK Or _
                FontTypeIndex = FONTTYPE_GUILDMSG Or FontTypeIndex = FONTTYPE_PIEL Or _
                FontTypeIndex = FONTTYPE_PIEL2)
                        
            If bUrl Then EnableUrlDetect
            
            .SelStart = Len(frmMain.RecTxt.Text)
            .SelLength = 0

            .SelBold = FontTypes(FontTypeIndex).bold
            .SelItalic = FontTypes(FontTypeIndex).italic
            
            If Not red = -1 Then .SelColor = RGB(FontTypes(FontTypeIndex).red, FontTypes(FontTypeIndex).green, FontTypes(FontTypeIndex).blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        End If
    End With
    
    If bUrl Then DisableUrlDetect
    
End Sub


'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()

    '*****************************************************************
    'Goes through the charlist and replots all the characters on the map
    'Used to make sure everyone is visible
    '*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar

        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).charindex = loopc

        End If

    Next loopc

End Sub

Function AsciiValidos(ByVal Cad As String) As Boolean

    Dim car As Byte

    Dim i   As Long
    
    Cad = LCase$(Cad)
    
    For i = 1 To Len(Cad)
        car = Asc(mid$(Cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function

        End If

    Next i
    
    AsciiValidos = True

End Function

Function CheckUserData() As Boolean

    Dim loopc As Integer
    Dim CharAscii As Integer

 
    If Len(Cuenta.UserPassword) = 0 Then
        MensajeAdvertencia Locale_GUI_Frase(256)
        CheckUserData = False
        Exit Function
    End If

    For loopc = 1 To Len(Cuenta.UserPassword)
        CharAscii = Asc(mid$(Cuenta.UserPassword, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia Locale_GUI_Frase(257)
            CheckUserData = False
            Exit Function
        End If
    Next loopc

    If Len(Cuenta.UserAccount) = 0 Then
        MensajeAdvertencia Locale_GUI_Frase(258)
        CheckUserData = False
        Exit Function
    End If
    
    If Len(Cuenta.UserAccount) > 20 Then
        MensajeAdvertencia Locale_GUI_Frase(259)
        CheckUserData = False
        Exit Function
    End If

    If Len(Cuenta.UserPassword) > 30 Then
        MensajeAdvertencia Locale_GUI_Frase(260)
        CheckUserData = False
        Exit Function
    End If


     For loopc = 1 To Len(Cuenta.UserAccount)
        CharAscii = Asc(mid$(Cuenta.UserAccount, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia Locale_GUI_Frase(251)
            CheckUserData = False
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True

End Function

Sub UnloadAllForms()

On Error Resume Next
 Dim miFrm As Form

For Each miFrm In Forms
    Unload miFrm
    Set miFrm = Nothing
Next

Reset

End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    'if backspace allow

    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters

    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function

    End If
    
    If KeyAscii > 126 Then
        Exit Function

    End If
    
    'Check for bad special characters in between

    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii _
            = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

End Function

Sub SetConnected()

    Debug.Print "Open SetConnected: " & Time
 
    'Set Connected
    Connected = True
    RenderInv = True
    
    CurrentUser.Logged = True
    CurrentUser.LogeoAlgunaVez = True
    
    frmMain.lblName.Caption = Cuenta.UserName
 
    If Len(Cuenta.UserName) > 15 Then
        frmMain.lblName.FontSize = 9
    Else
        frmMain.lblName.FontSize = 14
    End If

    'Vaciamos la cola de movimiento
    keysMovementPressedQueue.Clear
 
    'Actualizamos los macros
    Call LoadMacros(Cuenta.UserName)
    
    'Limpiamos los macros antes de usar
    Dim i As Byte
    
    For i = 1 To 11
       
        'Actualizamos el picmacro
        If MacroList(i).Grh <= 0 Then
        Else
        End If
    Next i
    
    Mod_General.modocombate
    
    If PrimeraVez = 1 Then
    frmElegirTeclas.Show , frmMain

    End If
    
    'If CantidadEnMacros Then Call UpdateMacroLabels(1)

    scroll_pixels_per_frame = scroll_pixels_per_frameBackUp
    'Load main form
    frmMain.Visible = True
    
    frmIniciando.Visible = False
    
    If CurrentUser.CurMapBattle <> 0 Then frmElegirTeclas.Show vbModeless, frmMain
    
    Call FormularioTimer(True)
    
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/28/2008
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
    ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
    ' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
    '***************************************************
    
    On Error GoTo MoveTo_Err
    
    Dim map_x As Byte
    Dim map_y As Byte
    
    Dim LegalOk As Boolean
     
    Select Case Direccion

        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1, Direccion)

        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y, Direccion)

        Case E_Heading.south
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1, Direccion)

        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y, Direccion)

    End Select
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
  
    
        If Not CurrentUser.UserDescansar Then
            Moviendose = True
            Call MainTimer.Restart(TimersIndex.Walk)
            MoveCharbyHead CurrentUser.UserCharIndex, Direccion
            MoveScreen Direccion
            
         End If
    Else
            If CurrentUser.UserDescansar Then
                WriteRest 'Stop resting (we do NOT have the 1 step enforcing anymore) sono como un tema de los guns.
            End If
        
        End If

 
        Call Char_MapPosGet(CurrentUser.UserCharIndex, map_x, map_y)

        If frmMain.UltPos = 0 Then
            If VerLugar = 1 Then
                frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & map_x & ", " & map_y
            End If
        ElseIf VerLugar = 0 Then
            frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & map_x & ", " & map_y
        End If

        Call ActualizarMiniMapa
    
    

        If charlist(CurrentUser.UserCharIndex).Heading <> Direccion Then
            If IntervaloPermiteHeading(True) Then
                Call WriteChangeHeading(Direccion)
            End If
        End If

   

    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)

    Exit Sub

MoveTo_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_General.MoveTo", Erl)
    Resume Next
    
End Sub

Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))

End Sub

Sub CheckKeys()
'***********************
'Checks keys and respond
'***********************
    Static LastMovement As Long
    
    Dim Direccion As E_Heading
    Direccion = charlist(CurrentUser.UserCharIndex).Heading
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
 
   'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
          If BloqueoAlCaminar Then Exit Sub
            
             If Not MainTimer.Check(TimersIndex.Walk, False) Then Exit Sub
             
             Call AddMovementToKeysMovementPressedQueue
        
             Select Case keysMovementPressedQueue.GetLastItem()
             
             Case (CustomKeys.BindedKey(eKeyType.mKeyUp))
               Call MoveTo(NORTH)
               
             Case (CustomKeys.BindedKey(eKeyType.mKeyRight))
               Call MoveTo(EAST)
               
             Case (CustomKeys.BindedKey(eKeyType.mKeyDown))
               Call MoveTo(south)
             Case (CustomKeys.BindedKey(eKeyType.mKeyLeft))
               Call MoveTo(WEST)
               
             End Select
            
         Else
         
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
                
                If kp Then
                Call RandomMove
                Exit Sub
                End If
   
          End If
          
          End If
   
   
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub
Sub SwitchMap(ByVal MapRoute As String)
 
    On Error GoTo SwitchMap_Err

    Char_Clean
    Particle_Group_Remove_All
    Light_Remove_All
 
    
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
    Dim npcs()       As tDatosNPC
    Dim TEs()        As tDatosTE
    Dim i            As Long
    Dim J            As Long
 
    Extract_File Maps, App.Path & "\Recursos", "mapa" & MapRoute & ".csm", Resource_Path

    
    fh = FreeFile

    If FileExist(App.Path & "\Recursos\Mapa" & CurrentUser.UserMap & ".csm", vbNormal) Then
       Open Resource_Path & "mapa" & MapRoute & ".csm" For Binary Access Read As fh
    Else
        'no encontramos el mapa en el hd
        Call MsgBox("Error en los mapas, algún archivo ha sido modificado o esta dañado.")
       ' MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        Call CloseClient
    End If
    
 
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat
   
    MapDat.map_name = Trim$(MapDat.map_name)

    ReDim MapData(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As MapBlock
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
    
    
    Get #fh, , L1
   
    
        For J = MapSize.YMin To MapSize.YMax
            For i = MapSize.XMin To MapSize.XMax
    
                If L1(i, J) > 0 Then
                    InitGrh MapData(i, J).Graphic(1), L1(i, J)
                End If
    
            Next i
        Next J
        
    With MH
     
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i

        End If
       
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex
            Next i

        End If
       
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex
            Next i

        End If
       
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex
            Next i

        End If
       
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i

        End If
       
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
     
            For i = 1 To .NumeroParticulas
                MapData(Particulas(i).X, Particulas(i).Y).particle_group = SetMapParticle(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y, -1)
            Next i

        End If
       
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces

            For i = 1 To .NumeroLuces
                If Not Luces(i).color = 0 Then
                Call Light_Create(Luces(i).X, Luces(i).Y, Luces(i).color, Luces(i).range)
                End If
            Next i

        End If
    
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            For i = 1 To .NumeroOBJs
                Map_Obj_Create Objetos(i).X, Objetos(i).Y, CInt(General_Locale_Obj(Objetos(i).OBJIndex, 3)), Objetos(i).OBJIndex, CInt(General_Locale_Obj(Objetos(i).OBJIndex, 2)), Objetos(i).ObjAmmount, 1
            Next i
        End If
                
        If .NumeroNPCs > 0 Then
            ReDim npcs(1 To .NumeroNPCs)
            Get #fh, , npcs
            For i = 1 To .NumeroNPCs
                MapData(npcs(i).X, npcs(i).Y).NPCIndex = npcs(i).NPCIndex
            Next
                
        End If
        
    End With

    Close fh

    Delete_File Resource_Path & "mapa" & MapRoute & ".csm"
    If FileExist(Resource_Path & "mapa" & MapRoute & ".csm", vbNormal) Then Kill Resource_Path & "mapa" & MapRoute & ".csm"
    
    If VerLugar = 1 Then frmMain.Label2(0).Caption = Map_Name_Get

    Dim r As Integer, g As Integer, b As Integer
    'Common light value verify
    If MapDat.Base_Light = 0 Then
        m_Afecta = True
        meteo_hour = 65
        Meteo_Change_Time
    Else
        General_Long_Color_to_RGB MapDat.Base_Light, r, g, b
        AmbientColor = D3DColorXRGB(r, g, b)
        ambientLight.r = r
        ambientLight.g = g
        ambientLight.b = b
        m_Afecta = False
    End If
    
    If MapDat.letter_grh <> 0 Then
        Map_Letter_Fade_Set MapDat.letter_grh
    Else
        Map_Letter_UnSet
    End If
    
    Light_Render_All
    
    Call DibujarMiniMapa
 
    If MapDat.battle_mode <> CurrentUser.CurMapBattle Then
        
        Dim Mensaje As String
        
        Select Case MapDat.battle_mode
        
            Case 0: Mensaje = Locale_GUI_Frase(556) 'Inse
            Case 1: Mensaje = Locale_GUI_Frase(553) 'impe
            Case 2: Mensaje = Locale_GUI_Frase(554) 'Repu
            Case 3: Mensaje = Locale_GUI_Frase(552) 'Neutral
            Case 4: Mensaje = Locale_GUI_Frase(555) 'caos
            
        End Select
    
    If Mensaje <> "" Then AddtoRichTextBox Mensaje, 0, 0, 0, 0, 0, 0, 12
    
    CurrentUser.CurMapBattle = MapDat.battle_mode
    
    End If

    Call Audio.PlayMIDI(MapDat.music_number)
     
    
    '*******************************
    'Render lights
    'Light_Render_All

    'For i = 1 To 5
    '    frmMain.Shape2(i).Visible = False
    '    frmMain.Label1(i).Visible = False
    '    'frmMap.Shape1(NumAmigo).Visible = False
    '    'frmMap.Label1(NumAmigo).Visible = False
    'Next i
 
    Exit Sub

SwitchMap_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_General.SwitchMap", Erl)
    Resume Next
    
End Sub
Public Sub ActualizarMiniMapa()
    '***************************************************
    'Author: Martin Gomez (Samke)
    'Last Modify Date: 21/03/2020 (ReyarB)
    'Integrado por Reyarb
    'Se agrego campo de vision del render (Recox)
    'Ajustadas las coordenadas para centrarlo (WyroX)
    'Ajuste de coordenadas y tamaÃ±o del visor (ReyarB)
    '***************************************************

    frmMain.UserP.Left = UserPos.X - 1
    frmMain.UserP.Top = UserPos.Y - 1
    frmMain.Minimap.Refresh
End Sub
Public Sub DibujarMiniMapa()

    On Error GoTo DibujarMiniMapa_Err
    
    Dim map_x As Long
    Dim map_y As Long
    frmMain.Minimap.Cls
    For map_y = MapSize.XMin To MapSize.XMax
        For map_x = MapSize.YMin To MapSize.YMax
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then SetPixel frmMain.Minimap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).mini_map_color

            If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then SetPixel frmMain.Minimap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).mini_map_color
 
            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then SetPixel frmMain.Minimap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).mini_map_color
        Next map_x
    Next map_y
    Exit Sub

DibujarMiniMapa_Err:
    Call RegistrarError(Err.number, Err.Description, "mod_General.DibujarMiniMapa", Erl)
    Resume Next
    
End Sub

Public Sub General_Long_Color_to_RGB(ByVal long_color As Long, _
                                     ByRef red As Integer, _
                                     ByRef green As Integer, _
                                     ByRef blue As Integer)
    '***********************************
    'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modified: 2/19/03
    'Takes a long value and separates RGB values to the given variables
    '***********************************
    Dim temp_color As String
    
    temp_color = Hex$(long_color)

    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String$(6 - Len(temp_color), "0") + temp_color

    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))

End Sub
Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, Value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function
Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
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

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim Count     As Long

    Dim curPos    As Long

    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")

End Function
 
Sub Main()

 Call Resolution.SetResolution

    On Error Resume Next

    Form_Caption = "Link-AO 1.4.5"
    
   #If Debugger = 0 Then
        If FindPreviousInstance Then
                Call MsgBox("¡Link-AO ya está corriendo! No es posible correr otra instancia del juego. Relea el reglamento. Haga click en Aceptar para salir." & vbCrLf & vbCrLf & "Link-AO is already running. Game cannot be run. Click OK to quit.", vbApplicationModal + vbInformation + vbOKOnly, "Already running!")
              'End
        End If
    #End If
    
    Set FormParser = New clsCursor
    Call FormParser.Init

    Call LoadCovAOInit
 
    'Security
    

    'AoDefAntiShInitialize
    
    'If AoDefDebugger Then
    '    Call AoDefAntiDebugger
    '    End
    'End If
    
        'Activar IMPORTANTE
    'If AoDefMultiClient Then
    '    Call AoDefMultiClientOn
    '    End
    'End If
    ''Security
    
    MacAdress = GetMacAddress
    HDserial = GetDriveSerialNumber
    
    Call LoadFontTypes

    Call General_SetIcon(frmMain.hwnd, "AAA", True)

    DoEvents
    
    Set ClientTCP = New clsClientTCP
    
    
    
 
       
    frmCargando.Show
 
    'Don't show cursor anymore
    If RunWindowed = 0 Then Call General_Cursor_Render(False)
        
    Call InicializarNombres
    
    Engine_DirectX8_Init
    InitTileEngine
    
    
    'Inicializaciones
     'Sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, Resource_Path, Resource_Path)
    Audio.MusicActivated = Musica
    Audio.SoundActivated = SonidoHabilitado
    Audio.SoundEffectsActivated = Efectos3D
    
    If Extract_File(MP3, App.Path & "\Recursos", "1.mp3", Resource_Path, False) Then
    Call Audio.MusicMP3Play(App.Path & "\RECURSOS\" & "1.mp3")
    Audio.SoundVolume = 100
    Delete_File Resource_Path & "1.mp3"
    End If
    
    'Esto es para el movimiento suave de pjs, para que el pj termine de hacer el movimiento antes de empezar otro
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
 
     'Inicializamos el socket
    Call frmMain.Socket1.Startup
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
 
    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    IntervaloCaminar = 110
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar)
    
    Call MainTimer.Start(TimersIndex.Walk)
    
    'frmPres.Picture = General_Load_Picture_From_Resource_Ex("_41")
    frmPres.Top = 0
    frmPres.Left = 0
    frmPres.Width = 800 * Screen.TwipsPerPixelX
    frmPres.Height = 600 * Screen.TwipsPerPixelY
    
    frmCargando.picLoad.Width = 500
    frmCargando.picLoad.Refresh

    frmPres.Visible = True
    Unload frmCargando
    
    Do While Not FinPres
       DoEvents
    Loop
    
    frmConnect.Visible = True
    Unload frmPres

    'Well let's leave this until GUI is done...
    If RunWindowed = 0 Then Call General_Cursor_Render(True)

    'Inicialización de variables globales
    RenderInv = False
    
    CurrentUser.LogeoAlgunaVez = False
    
    prgRun = True
    pausa = False
    
    Call StartClient
    
End Sub

 
 
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or ( _
            iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or (MapData( _
            X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or (MapData(X, _
            Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And MapData(X, _
            Y).Graphic(2).GrhIndex = 0
                
End Function

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************

        Dim loopc As Long

        Ciudades(eCiudad.cNix) = "Nix (Imperial)"
        Ciudades(eCiudad.cilliandor) = "Illiandor (Republicano)"
        
        ReDim ListaRazas(1 To NUMRAZAS) As String
        
        For loopc = 1 To NUMRAZAS
            ListaRazas(loopc) = Locale_GUI_Frase(130 + loopc)
        Next loopc
        
        ReDim ListaClases(1 To NUMCLASES) As String
        
        For loopc = 1 To NUMCLASES
            ListaClases(loopc) = Locale_GUI_Frase(112 + loopc)
        Next loopc

 
        ReDim Head_Range(1 To NUMRAZAS) As tHeadRange
        
        'Male heads
        Head_Range(Humano).mStart = 1
        Head_Range(Humano).mEnd = 30
        Head_Range(enano).mStart = 301
        Head_Range(enano).mEnd = 315
        Head_Range(Elfo).mStart = 101
        Head_Range(Elfo).mEnd = 121
        Head_Range(ELFOOSCURO).mStart = 202
        Head_Range(ELFOOSCURO).mEnd = 212
        Head_Range(gnomo).mStart = 401
        Head_Range(gnomo).mEnd = 409
        Head_Range(Orco).mStart = 501
        Head_Range(Orco).mEnd = 514
        
        'Female heads
        Head_Range(Humano).fStart = 70
        Head_Range(Humano).fEnd = 80
        Head_Range(enano).fStart = 370
        Head_Range(enano).fEnd = 373
        Head_Range(Elfo).fStart = 170
        Head_Range(Elfo).fEnd = 189
        Head_Range(ELFOOSCURO).fStart = 270
        Head_Range(ELFOOSCURO).fEnd = 278
        Head_Range(gnomo).fStart = 470
        Head_Range(gnomo).fEnd = 481
        Head_Range(Orco).fStart = 570
        Head_Range(Orco).fEnd = 573
    
        ReDim SkillsNames(1 To NUMSKILLS) As String

        For loopc = 1 To NUMSKILLS
            SkillsNames(loopc) = Locale_GUI_Frase(302 + loopc)
        Next loopc

 
End Sub
 
Public Sub CloseClient(Optional ByVal Closed_ByUser As Boolean = False, Optional ByVal Init_Launcher As Boolean = False)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    ' Allow new instances of the client to be opened
 
    EngineRun = False
    
    '0. Cerramos el socket
    If frmMain.Socket1.State <> sckClosed Then frmMain.Socket1.Disconnect

    '1. Guardamos datos si se cerró correctamente
    If Closed_ByUser Then Call SaveCovAOInit
    
    '2. Eliminamos objetos DX
    Call Directx_DeInitialize
   
    
    
    'Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Set FormParser = Nothing
    Set ClientTCP = Nothing
    
    'Closed Timer
    Call FXTimer(False)
    Call MinutoTimer(False)
    Call FormularioTimer(False)
    
    Call Resolution.ResetResolution
    
    Call UnloadAllForms
        
    KillTimer 0, thFPSAndHour
    
    Call PrevInstance.ReleaseInstance
    
    Call General_Set_Mouse_Speed(10) 'volvemos a la posicion original del mouse
    '//Mermas, optimización de carga de sonidos, en vez de bajar todos los wavs y borrarlos al cierre, bajamos el necesario, _
     y sumamos indices para verificar si sirve y borrarlo, en caso que no, sumamos el indice + 1
 
    
    '8. ¿Había que prender el launcher?
    If Init_Launcher Then ShellExecute GetDesktopWindow, "open", App.Path & "\Link-AOLauncher.exe", vbNullString, vbNullString, 1
    
    End
End Sub
 
Public Function getTagPosition(ByVal Nick As String) As Integer

    Dim buf As Integer

    buf = InStr(Nick, "<")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function

    End If

    buf = InStr(Nick, "[")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function

    End If

    getTagPosition = Len(Nick) + 2

End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer

    Dim i As Long

    For i = 1 To LastChar

        If charlist(i).Nombre = Name Then
            getCharIndexByName = i
            Exit Function

        End If

    Next i

End Function

Public Sub LogError(ByVal Desc As String)

    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\" & "error.log" For Append As #nFile
    Print #nFile, Desc
    Close #nFile

End Sub

Public Sub LogCustom(ByVal Desc As String, ByVal File As String)

    Dim nFile As Integer
    nFile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\" & File & ".log" For Append As #nFile
    Print #nFile, Now & " " & Desc
    Close #nFile

End Sub

Private Sub AddMovementToKeysMovementPressedQueue()
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Remueve la tecla que teniamos presionada
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''clsMeteorologic''''''''''''''''''''''''''''''''''''''''''''''''''''
  Public Function General_Get_Mouse_Speed() As Long
 
    SystemParametersInfo SPI_GETMOUSESPEED, 0, General_Get_Mouse_Speed, 0
     
    End Function
     
    Public Sub General_Set_Mouse_Speed(ByVal lngSpeed As Long)
 
     
    SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
     
    End Sub
Public Sub SetFXMAP(ByVal fX As Integer, ByVal X As Integer, ByVal Y As Byte, ByVal Loops As Integer)

    If X = 0 Or Y = 0 Then Exit Sub
    
    MapData(X, Y).FxIndex = fX
    
    If MapData(X, Y).FxIndex > 0 Then
        InitGrh MapData(X, Y).fX, FxData(fX).Animacion
        
        MapData(X, Y).fX.GrhIndex = FxData(fX).Animacion
        
        MapData(X, Y).fX.Loops = Loops
    End If
    
End Sub

 


 Sub UpdateMacroLabels(ByVal Index As Byte)
 
    Dim loopc As Long
    
    
    For loopc = 1 To 11
    
    If Index = 0 Then 'Apagado
    
        If MacroList(loopc).mTipe = 3 Or MacroList(loopc).mTipe = 4 Then
        End If
    Else
    End If
    
    Next loopc
End Sub
 
Public Function CerrarJuego()
   ' frmPres.Picture = General_Load_Picture_From_Resource_Ex("_13")
    frmPres.Visible = True
    prgRun = False
    
End Function

Public Function modocombate()
If IScombate Then
frmMain.modocombate.Visible = True
frmMain.nomodocombate.Visible = False
Else
frmMain.modocombate.Visible = False
frmMain.nomodocombate.Visible = True
End If
End Function


Public Sub General_SetIcon( _
      ByVal hwnd As Long, _
      ByVal sIconResName As String, _
      Optional ByVal bSetAsAppIcon As Boolean = True)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: Unknown
'
'**************************************************************

Dim lhWndTop As Long
Dim lhWnd As Long
Dim cX As Long
Dim cY As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hwnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindow(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cX = GetSystemMetrics(SM_CXICON)
   cY = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cX, cY, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cX = GetSystemMetrics(SM_CXSMICON)
   cY = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cX, cY, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub

Public Sub LoadCovAOInit()

On Error Resume Next

    If Not LoadLocales Then
        MsgBox "¡No se ha logrado realizar la carga del archivo de idioma! Verifique la integridad del sistema. Si el problema persiste por favor consulte los foros de soporte." & vbCrLf & vbCrLf & "Locale data file could not be loaded. Please refer to tech support if the problem persists.", vbCritical, "Saliendo / Quitting"
        Call CerrarJuego
    End If

    InitCommonControls 'Inicia manifiesto (ejecuta como admin y dar controladores actualizados)
 
    
    'Init Paths
    Resource_Path = App.Path & "\Recursos\"
    DirSounds = App.Path & "\Recursos\Sonidos\"
    DirMidi = App.Path & "\Recursos\Sonidos\"
    DirInit = App.Path & "\Init\"
 
  
    Dim Leer As New clsIniManager
    
    Call Leer.Initialize(DirInit & "CovAoInit.ini")
    
    'Si tiene WindowsXP
    Win2kXP = General_Windows_Is_2000XP
    
    
    'Inicializar configuración de Launcher
    Musica = val(Leer.GetValue("LAUNCHER", "Musica")) 'Music
    
    SonidoHabilitado = val(Leer.GetValue("LAUNCHER", "SonidoHabilitado")) 'Sounds

    Efectos3D = val(Leer.GetValue("LAUNCHER", "Efectos3D")) 'Effects

    VSYNC = val(Leer.GetValue("LAUNCHER", "VSYNC")) 'VSYNC

    RunWindowed = val(Leer.GetValue("LAUNCHER", "RunWindowed")) 'Fullscreen
 
    DeviceIndex = val(Leer.GetValue("LAUNCHER", "DeviceIndex")) 'Software, Hardware, Mixed, Automatico

    Pixels = val(Leer.GetValue("LAUNCHER", "Pixels")) 'Pixels 16/32
    
    FXVolume = val(Leer.GetValue("LAUNCHER", "FXVolume")) 'Pixels 16/32

    MusicVolume = val(Leer.GetValue("LAUNCHER", "MusicVolume")) ' Al iniciar esta variable damos tutoriales, etc

    NombreSkin = Leer.GetValue("LAUNCHER", "NombreSkin") 'Cargar nombre de Skins

    PrimeraVez = val(Leer.GetValue("LAUNCHER", "PrimeraVez")) ' Al iniciar esta variable damos tutoriales, etc
    
    Desvanecimiento = val(Leer.GetValue("LAUNCHER", "Desvanecimiento")) 'Desvanecimientos efectos que consumen FPS
 
    'Fin config Launcher
    
    
    'Configuración InGame
    
    HabilitarMensajesGlobales = val(Leer.GetValue("CONFIG", "HabilitarMensajesGlobales"))
    VerLugar = val(Leer.GetValue("CONFIG", "VerLugar"))
    RecordarCuentaIni = val(Leer.GetValue("CONFIG", "RecordarCuenta"))
    If RecordarCuentaIni = True Then UserAccountRecorded = Leer.GetValue("CONFIG", "Cuenta")
    
    NickModerno = val(Leer.GetValue("CONFIG", "NickModerno"))
    CargarMedidasNombresModernos
    NombresModernos = IIf((NickModerno = True), 5, 1)
    
    ListaIgnorados = Leer.GetValue("CONFIG", "ListaIgnorados")
    CursorHabilitado = val(Leer.GetValue("CONFIG", "CursorHabilitado")) 'Cursor habilitado
    MouseSpeed = val(Leer.GetValue("CONFIG", "MouseSpeed")) 'Sensibilidad del mouse
    BloqueoAlCaminar = val(Leer.GetValue("CONFIG", "BloqueoAlCaminar")) 'Bloqueo al caminar
    CantidadEnMacros = val(Leer.GetValue("CONFIG", "CantidadEnMacros"))
    AutoUsarActivado = val(Leer.GetValue("CONFIG", "AutoUsarActivado"))
    accionMouseUno = val(Leer.GetValue("CONFIG", "AccionMouseUno"))
    accionMousedos = val(Leer.GetValue("CONFIG", "AccionMouseDos"))
    Rendimiento = val(Leer.GetValue("CONFIG", "Rendimiento"))
    
    If Not General_File_Exists(App.Path & "\Skins\" & NombreSkin & ".ias", vbNormal) Then
        Call MsgBox(Locale_GUI_Frase(332) & " " & NombreSkin & ".ias " & Locale_GUI_Frase(333), vbCritical, Locale_GUI_Frase(333))
        Call CerrarJuego
    End If
    
    Call Set_Skin_Name(NombreSkin)
    
    Call General_Set_Mouse_Speed(MouseSpeed)
    


    Load frmCargando
    'Load frmPres
    Load frmConnect
    Load frmMensaje
    Load frmMain
    Load frmCharList
    Load frmOpciones
    Load frmHlp
    
    Leer = Nothing
End Sub

Public Function Tilde(data As String) As String
 
    Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function


Public Function MostrarCantidad(ByVal i As Integer) As Boolean
    Dim ObjType As Integer
    ObjType = CInt(General_Locale_Obj(i, 2))
    
    MostrarCantidad = ObjType <> eObjType.otPuertas And _
            ObjType <> eObjType.otCarteles And _
            ObjType <> eObjType.otArboles And _
            ObjType <> eObjType.otYacimiento And _
            ObjType <> eObjType.otTeleport And _
            ObjType <> eObjType.otcorreo
End Function

Public Function BoolToByte(ByVal val As Boolean) As Byte
    BoolToByte = IIf(val = True, 1, 0)
End Function
Public Function BoolToInteger(ByVal val As Boolean) As Integer
    BoolToInteger = IIf(val = True, 1, 0)
End Function
Function Generate_Char_Status(ByVal PercVida As Long, ByVal Paralizado As Byte, ByVal Inmovilizado As Byte, _
                              Optional ByVal Incinerado As Byte = 0, Optional ByVal Envenenado As Byte = 0, _
                              Optional ByVal Comerciando As Byte = 0, Optional ByVal Trabajando As Byte = 0, _
                              Optional ByVal Combatiendo As Byte = 0, Optional ByVal Ciego As Byte = 0, _
                              Optional ByVal Inactivo As Byte = 0, Optional ByVal Resucitando As Byte = 0, Optional ByVal Saliendo As Byte = 0) As String

If PercVida <> -1 Then
    If PercVida = 100 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(445) & " "
    ElseIf PercVida >= 80 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(446) & " "
    ElseIf PercVida >= 50 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(447) & " "
    ElseIf PercVida >= 30 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(448) & " "
    ElseIf PercVida <> 0 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(449) & " "
    Else
        Generate_Char_Status = " " & Locale_GUI_Frase(450) & " "
    End If
End If

If Paralizado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(451) & " "
End If

If Inmovilizado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(452) & " "
End If

If Incinerado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(454) & " "
End If

If Envenenado > 0 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(453) & " "
End If

If Comerciando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(462) & " "
End If

If Trabajando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(463) & " "
End If
 
If Combatiendo = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(547) & " "
End If
 
If Ciego = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(455) & " "
End If
 
If Inactivo = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(464) & " "
End If

If Resucitando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(551) & " "
End If

If Saliendo = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(331) & " "
End If

Generate_Char_Status = RTrim(Generate_Char_Status)
End Function

Public Function TituloCaos(ByVal rango As Byte) As String
    
    Select Case rango
        Case 1
            TituloCaos = "Miembro de las Hordas"
        Case 2
            TituloCaos = "Guerrero del Caos"
        Case 3
            TituloCaos = "Teniente del Caos"
        Case 4
            TituloCaos = "Comandante del Caos"
        Case 5
            TituloCaos = "General del Caos"
        Case 6
            TituloCaos = "Elite del Caos"
        Case 7
            TituloCaos = "Asolador de las Sombras"
        Case 8
            TituloCaos = "Caballero Negro"
        Case 9
            TituloCaos = "Emisario de las Sombras"
        Case 10
            TituloCaos = "Avatar del Apocalipsis"
    End Select

       
End Function
Public Function TituloReal(ByVal rango As Byte) As String

    Select Case rango
        Case 1
            TituloReal = "Legionario"
        Case 2
            TituloReal = "Soldado Real"
        Case 3
            TituloReal = "Teniente Real"
        Case 4
            TituloReal = "Comandante Real"
        Case 5
            TituloReal = "General Real"
        Case 6
            TituloReal = "Elite Real"
        Case 7
            TituloReal = "Guardian del Bien"
        Case 8
            TituloReal = "Caballero Imperial"
        Case 9
            TituloReal = "Justiciero"
        Case 10
            TituloReal = "Guardia Imperial"
    End Select
    
End Function
Public Function TituloMilicia(ByVal rango As Byte) As String

    Select Case rango
        Case 1
            TituloMilicia = "Milicia de Reserva"
        Case 2
            TituloMilicia = "Miliciano"
        Case 3
            TituloMilicia = "Miliciano Elite"
        Case 4
            TituloMilicia = "Soldado de la República"
        Case 5
            TituloMilicia = "Soldado Raso"
        Case 6
            TituloMilicia = "Soldado Elite"
        Case 7
            TituloMilicia = "Comandante de la República"
    End Select
    

       
End Function



Public Sub LoadFontTypes()

On Error GoTo errorhandler

Dim lC As Integer, Arch As String, tempStr As String

If Not Extract_File(Scripts, App.Path & "\Recursos", "fonttypes.ind", Resource_Path, False) Then
    Err.Description = "No se ha logrado extraer el archivo de recurso."
    GoTo errorhandler
End If

Arch = Resource_Path & "fonttypes.ind"
NUMFONTS = val(General_Var_Get(Arch, "INIT", "NumFonts"))
ReDim Preserve FontTypes(1 To NUMFONTS) As tFontType
 
For lC = 1 To NUMFONTS

1    tempStr = General_Var_Get(Arch, "INIT", str(lC))
2    FontTypes(lC).red = val(General_Field_Read(2, tempStr, "~"))
3    FontTypes(lC).green = val(General_Field_Read(3, tempStr, "~"))
4    FontTypes(lC).blue = val(General_Field_Read(4, tempStr, "~"))
5    FontTypes(lC).bold = val(General_Field_Read(5, tempStr, "~"))
6    FontTypes(lC).italic = val(General_Field_Read(6, tempStr, "~"))
Next lC

Delete_File Resource_Path & "fonttypes.ind"

Exit Sub
    
errorhandler:
    Call RegistrarError(Err.number, Err.Description, "Protocol.LoadFontypes", Erl)
    If General_File_Exists(Resource_Path & "fonttypes.ind", vbNormal) Then Delete_File Resource_Path & "fonttypes.ind"

End Sub

Public Sub SaveCovAOInit()

Dim lC As Integer, Arch As String

Arch = DirInit & "CovAoInit.ini"

'Call General_Var_Write(Arch, "INIT", "NUMBINDS", str(NUMBINDS))
'Call General_Var_Write(Arch, "INIT", "NUMBOTONES", str(NUMBOTONES))
'Call General_Var_Write(Arch, "INIT", "VerLugar", str(VerLugar))
'Call General_Var_Write(Arch, "INIT", "FxNavega", str(FxNavega))
'Call General_Var_Write(Arch, "INIT", "DefaultMidi", str(DefMidi))
'Call General_Var_Write(Arch, "INIT", "gldf", str(gldf))
'Call General_Var_Write(Arch, "INIT", "CopiarDialogos", str(CopiarDialogos))
'Call General_Var_Write(Arch, "INIT", "MensajesGlobales", str(MensajesGlobales))
'Call General_Var_Write(Arch, "INIT", "MensajesFaccionarios", str(MensajesFaccionarios))
'Call General_Var_Write(Arch, "INIT", "CopiarDialogos", str(CopiarDialogos))
'Call General_Var_Write(Arch, "INIT", "MusicVolume", str(MusicVolume))
'Call General_Var_Write(Arch, "INIT", "FXVolume", str(FXVolume))
'Call General_Var_Write(Arch, "INIT", "InvertirSonido", IIf(InvertirSonido = True, "1", "0"))
'Call General_Var_Write(Arch, "INIT", "Musica", str(sMusica))
'Call General_Var_Write(Arch, "INIT", "SonidoHabilitado", str(Audio))
'Call General_Var_Write(Arch, "INIT", "NombreSkin", NombreSkin)
'Call General_Var_Write(Arch, "INIT", "NombresSimples", str(NombresSimples))
'Call General_Var_Write(Arch, "INIT", "MouseSpeed", str(MouseS))
'Call General_Var_Write(Arch, "INIT", "Publicidad_Contenido", str(Publicidad_Contenido))
'Call General_Var_Write(Arch, "INIT", "CursoresStandar", str(CursoresStandar))
'Call General_Var_Write(Arch, "INIT", "GameLocale", GameLocale)
'Call General_Var_Write(Arch, "INIT", "LastRunDate", str(LastRunDate))

Call General_Var_Write(Arch, "CONFIG", "NickModerno", IIf((NickModerno = True), 1, 0))

If RecordarCuentaIni = True Then
    If Len(UserAccountRecorded) > 0 Then
        Call General_Var_Write(Arch, "CONFIG", "Cuenta", UserAccountRecorded)
    Else
        RecordarCuentaIni = False
        Call General_Var_Write(Arch, "CONFIG", "Cuenta", "")
    End If
    
Else
    Call General_Var_Write(Arch, "CONFIG", "Cuenta", "")
    
End If

Call General_Var_Write(Arch, "CONFIG", "RecordarCuenta", IIf((RecordarCuentaIni = True), 1, 0))

Call General_Var_Write(Arch, "CONFIG", "VerLugar", VerLugar)
Call General_Var_Write(Arch, "CONFIG", "HabilitarMensajesGlobales", IIf((HabilitarMensajesGlobales = True), 1, 0))


'For lc = 1 To NUMBINDS
'    Call General_Var_Write(Arch, "User", str(lc), str(BindKeys(lc).KeyCode) & "," & BindKeys(lc).Name)
'Next lc
 
'lc = 0

'For lc = 1 To NUMBOTONES
'    Call General_Var_Write(Arch, "Bind" & lc, "Accion", str(MacroKeys(lc).TipoAccion))
'    Call General_Var_Write(Arch, "Bind" & lc, "hlist", str(MacroKeys(lc).hlist))
'    Call General_Var_Write(Arch, "Bind" & lc, "invslot", str(MacroKeys(lc).invslot))
'    Call General_Var_Write(Arch, "Bind" & lc, "SndString", MacroKeys(lc).SendString)
'Next lc

ListaIgnorados = vbNullString

For lC = 0 To frmOpciones.lstIgnore.ListCount
    If frmOpciones.lstIgnore.list(lC) <> vbNullString Then
        ListaIgnorados = ListaIgnorados & frmOpciones.lstIgnore.list(lC) & "¬"
    End If
Next lC

If ListaIgnorados <> vbNullString Then _
    ListaIgnorados = Left$(ListaIgnorados, Len(ListaIgnorados) - 1)

Call WriteVar(Arch, "CONFIG", "ListaIgnorados", ListaIgnorados)
End Sub

'Carga de nombre de mapas
Public Function Map_NameLoad(ByVal map_num As Integer) As String
On Error GoTo errorhandler
    
    Map_NameLoad = MapNames(map_num)
    Exit Function

errorhandler:
    Map_NameLoad = "Mapa Desconocido"

End Function
Public Function Map_Name_Get() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/9/2005
'
'**************************************************************
    Map_Name_Get = Trim$(MapDat.map_name)
End Function
Function Generate_Char_StatusNPCs(ByVal PercVida As Long, ByVal Paralizado As Byte, ByVal Inmovilizado As Byte, ByVal PuedeVerVida As Byte) As String

If PuedeVerVida = 0 Then
    If PercVida <> -1 Then
        If PercVida = 100 Then
            Generate_Char_StatusNPCs = " " & Locale_GUI_Frase(445) & " "
        ElseIf PercVida >= 80 Then
            Generate_Char_StatusNPCs = " " & Locale_GUI_Frase(446) & " "
        ElseIf PercVida >= 50 Then
            Generate_Char_StatusNPCs = " " & Locale_GUI_Frase(447) & " "
        ElseIf PercVida >= 30 Then
            Generate_Char_StatusNPCs = " " & Locale_GUI_Frase(448) & " "
        ElseIf PercVida <> 0 Then
            Generate_Char_StatusNPCs = " " & Locale_GUI_Frase(449) & " "
        Else
            Generate_Char_StatusNPCs = " " & Locale_GUI_Frase(450) & " "
        End If
    End If
End If

If Paralizado = 1 Then
    Generate_Char_StatusNPCs = Generate_Char_StatusNPCs & "| " & Locale_GUI_Frase(451) & " "
End If

If Inmovilizado = 1 Then
    Generate_Char_StatusNPCs = Generate_Char_StatusNPCs & "| " & Locale_GUI_Frase(452) & " "
End If

Generate_Char_StatusNPCs = RTrim(Generate_Char_StatusNPCs)
End Function

