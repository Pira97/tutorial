Attribute VB_Name = "Extra"
Option Explicit

Public Function ClaseToEnum(ByVal Clase As String) As eClass
Dim i As Byte
    For i = 1 To NUMCLASES
        If UCase$(ListaClases(i)) = UCase$(Clase) Then
            ClaseToEnum = i
        End If
    Next i
End Function

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    esArmada = (UserList(UserIndex).Faccion.Status = 5)
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    esCaos = (UserList(UserIndex).Faccion.Status = 4)
End Function

Public Function esMili(ByVal UserIndex As Integer) As Boolean
    esMili = (UserList(UserIndex).Faccion.Status = 6)
End Function

Public Function esFaccion(ByVal UserIndex As Integer) As Boolean
    esFaccion = (UserList(UserIndex).Faccion.Status = 4 Or UserList(UserIndex).Faccion.Status = 5 Or UserList(UserIndex).Faccion.Status = 6)
End Function

Public Function esRene(ByVal UserIndex As Integer) As Boolean
    esRene = (UserList(UserIndex).Faccion.Status = 1)
End Function

Public Function esCiuda(ByVal UserIndex As Integer) As Boolean
    esCiuda = (UserList(UserIndex).Faccion.Status = 2)
End Function

Public Function esRepu(ByVal UserIndex As Integer) As Boolean
    esRepu = (UserList(UserIndex).Faccion.Status = 3)
End Function

Public Function EsGm(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    EsGm = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or _
            PlayerType.Consejero))

End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    
    On Error GoTo ErrHandler
 
    'Controla las salidas
    If InMapBounds(Map, X, Y) Then

        With MapData(Map, X, Y)
        
            Dim nPos       As WorldPos
            Dim FxFlag     As Boolean
            
             
            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
            End If
           
            If .TileExit.Map > 0 And .TileExit.Map <= NumMaps Then
                 
                 '¿Es mapa de newbies?
                 If .TileExit.Map = 37 Or .TileExit.Map = 208 Then 'UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Zona) = "NEWBIE" Then
                      
                    '¿El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)

                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                            
                        End If
                        
                    Else 'No es newbie
                        Call WriteLocaleMsg(UserIndex, 152, "1%14")
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        End If
                        
                    End If

                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.

                    If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)

                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If

                    End If

                End If
                
                Select Case MapData(Map, X, Y).TileExit.Map
                    Case 849 'cementerio nix
                        If UserList(UserIndex).Stats.ELV > 24 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 24 o anterior.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
        
                    Case 116 'maravel
                        If UserList(UserIndex).Stats.ELV < 24 And UserList(UserIndex).Stats.ELV > 40 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 24 a 40.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
            
                    Case 230 ' DC
                        If UserList(UserIndex).Stats.ELV < 42 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 42 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
        
                    Case 758 'dungeon arriva de nueva
                        If UserList(UserIndex).Stats.ELV < 30 And UserList(UserIndex).Stats.ELV > 40 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 30 a 40.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
            
                    Case 209 'DZ
                        If UserList(UserIndex).Stats.ELV < 40 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 40 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 755 'DF
                        If UserList(UserIndex).Stats.ELV < 35 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 35 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 756 'df 2do piso
                        If UserList(UserIndex).Stats.ELV < 38 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 38 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 760 'draogn legend
                        If UserList(UserIndex).Stats.ELV < 45 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 45 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 205 'minas illiandor
                        If UserList(UserIndex).Stats.ELV < 35 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 35 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 207 'cueva iliandor izq
                        If UserList(UserIndex).Stats.ELV < 23 And UserList(UserIndex).Stats.ELV > 35 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 22 a 35.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 140 'DV
                        If UserList(UserIndex).Stats.ELV < 35 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 35 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 830 'templo kalat
                        If UserList(UserIndex).Stats.ELV < 45 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 45 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 833 'templo kalat 2do piso
                        If UserList(UserIndex).Stats.ELV < 50 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 50 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                    Case 827 ' casti del vampi
                        If UserList(UserIndex).Stats.ELV < 48 Then
                            Call WriteConsoleMsg(UserIndex, "Mapa exclusivamente para personajes de nivel 48 a 60.", FontTypeNames.FONTTYPE_INFO)
                            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False)
                        End If
                End Select
                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
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

            End If

        End With

    End If
 
    Exit Sub

ErrHandler:
132     Call RegistrarError(Err.Number, Err.description, "Extra.DoTileEvents", Erl)
134     Resume Next
End Sub

Function InRangoVision(ByVal UserIndex As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
        If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function

        End If

    End If

    InRangoVision = False

End Function

Public Function InVisionRangeAndMap(ByVal UserIndex As Integer, _
                                    ByRef OtherUserPos As WorldPos) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        ' Same map?
        If .Pos.Map <> OtherUserPos.Map Then Exit Function
    
        ' In x range?
        If OtherUserPos.X < .Pos.X - MinXBorder Or OtherUserPos.X > .Pos.X + MinXBorder Then Exit Function
        
        ' In y range?
        If OtherUserPos.Y < .Pos.Y - MinYBorder And OtherUserPos.Y > .Pos.Y + MinYBorder Then Exit Function

    End With

    InVisionRangeAndMap = True
    
End Function

Function InRangoVisionNPC(ByVal npcindex As Integer, _
                          X As Integer, _
                          Y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If X > Npclist(npcindex).Pos.X - MinXBorder And X < Npclist(npcindex).Pos.X + MinXBorder Then
        If Y > Npclist(npcindex).Pos.Y - MinYBorder And Y < Npclist(npcindex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function

        End If

    End If

    InRangoVisionNPC = False

End Function

Function InMapBounds(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True

    End If
    
End Function

Sub ClosestLegalPos(Pos As WorldPos, _
                    ByRef nPos As WorldPos, _
                    Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True)
    '*****************************************************************
    'Author: Unknown (original version)
    'Last Modification: 24/01/2007 (ToxicWaste)
    'Encuentra la posicion legal mas cercana y la guarda en nPos
    '*****************************************************************

    Dim Notfound As Boolean
    Dim loopc    As Integer
    Dim tX       As Long
    Dim tY       As Long

    nPos.Map = Pos.Map

    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)

        If loopc > 12 Then
            Notfound = True
            Exit Do

        End If
    
        For tY = Pos.Y - loopc To Pos.Y + loopc
            For tX = Pos.X - loopc To Pos.X + loopc
            
                If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra) Then
                    nPos.X = tX
                    nPos.Y = tY
                    '¿Hay objeto?
                
                    tX = Pos.X + loopc
                    tY = Pos.Y + loopc

                End If

            Next tX
        Next tY
    
        loopc = loopc + 1
    Loop

    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0

    End If

End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
    '*****************************************************************

    Dim Notfound As Boolean
    Dim loopc    As Integer
    Dim tX       As Long
    Dim tY       As Long
    
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)

        If loopc > 12 Then
            Notfound = True
            Exit Do

        End If
        
        For tY = Pos.Y - loopc To Pos.Y + loopc
            For tX = Pos.X - loopc To Pos.X + loopc
                
                If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                    nPos.X = tX
                    nPos.Y = tY
                    '¿Hay objeto?
                    
                    tX = Pos.X + loopc
                    tY = Pos.Y + loopc

                End If

            Next tX
        Next tY
        
        loopc = loopc + 1
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0

    End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim UserIndex As Long
    
    '¿Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function

    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))

    End If
    
    UserIndex = 1

    Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function

        End If

    Loop
     
    NameIndex = UserIndex

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Long
    
    For loopc = 1 To MaxUsers

        If UserList(loopc).flags.UserLogged = True Then
            If UserList(loopc).ip = UserIP And UserIndex <> loopc Then
                CheckForSameIP = True
                Exit Function

            End If

        End If

    Next loopc
    
    CheckForSameIP = False

End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
        On Error GoTo HeadtoPos_Err
        
        '*****************************************************************
        'Toma una posicion y se mueve hacia donde esta perfilado
        '*****************************************************************
        Dim X  As Integer

        Dim Y  As Integer

        Dim nX As Integer

        Dim nY As Integer
        
100     X = Pos.X
102     Y = Pos.Y

    Select Case Head

        Case eHeading.NORTH
106         nX = X
108         nY = Y - 1
        
        Case eHeading.SOUTH
112         nX = X
114         nY = Y + 1
        
        Case eHeading.EAST
118         nX = X + 1
120         nY = Y
        Case eHeading.WEST
124         nX = X - 1
126         nY = Y

    End Select

        'Devuelve valores
128     Pos.X = nX
130     Pos.Y = nY

        
        Exit Sub

HeadtoPos_Err:
132     Call RegistrarError(Err.Number, Err.description, "Extra.HeadtoPos", Erl)
134     Resume Next
        
End Sub

Function LegalPos(ByVal Map As Integer, _
                  ByVal X As Integer, _
                  ByVal Y As Integer, _
                  Optional ByVal PuedeAgua As Boolean = False, _
                  Optional ByVal PuedeTierra As Boolean = True) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Checks if the position is Legal.
    '***************************************************

    '¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPos = False
    Else

        With MapData(Map, X, Y)

            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.npcindex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.npcindex = 0) And (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.npcindex = 0) And (HayAgua(Map, X, Y))
            Else
                LegalPos = False

            End If

        End With

    End If

End Function

Function MoveToLegalPos(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        Optional ByVal PuedeAgua As Boolean = False, _
                        Optional ByVal PuedeTierra As Boolean = True) As Boolean
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 13/07/2009
    'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
    '13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
    '***************************************************

    Dim UserIndex        As Integer
    Dim IsDeadChar       As Boolean
    Dim IsAdminInvisible As Boolean

    '¿Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        MoveToLegalPos = False
    Else

        With MapData(Map, X, Y)
            UserIndex = .UserIndex
        
            If UserIndex > 0 Then
                IsDeadChar = (UserList(UserIndex).flags.Muerto = 1)
                IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
            Else
                IsDeadChar = False
                IsAdminInvisible = False

            End If
        
            If PuedeAgua And PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.npcindex _
                        = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.npcindex _
                        = 0) And (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.npcindex _
                        = 0) And (HayAgua(Map, X, Y))
            Else
                MoveToLegalPos = False

            End If

        End With

    End If

End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByRef X As Integer, _
                        ByRef Y As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 26/03/2009
    'Search for a Legal pos for the user who is being teleported.
    '***************************************************

    If MapData(Map, X, Y).UserIndex <> 0 Or MapData(Map, X, Y).npcindex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace     As Boolean
        Dim tX             As Long
        Dim tY             As Long
        Dim Rango          As Long
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango

                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).npcindex = 0 Then
                        
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                        Exit For

                    End If

                Next tX
        
                If FoundPlace Then Exit For
            Next tY
            
            If FoundPlace Then Exit For
        Next Rango
    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            Dim OtherUserIndex As Integer
            OtherUserIndex = MapData(Map, X, Y).UserIndex

            If OtherUserIndex <> 0 Then
            
                Call CloseSocket(OtherUserIndex)

            End If

        End If

    End If

End Sub

Function LegalPosNPC(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal AguaValida As Byte, _
                     Optional ByVal IsPet As Boolean = False) As Boolean
    '***************************************************
    'Autor: Unkwnown
    'Last Modification: 09/23/2009
    'Checks if it's a Legal pos for the npc to move to.
    '09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
    '***************************************************
    Dim IsDeadChar       As Boolean
    Dim UserIndex        As Integer
    Dim IsAdminInvisible As Boolean
    
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function

    End If

    With MapData(Map, X, Y)
        UserIndex = .UserIndex

        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False

        End If
    
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.npcindex = 0) _
                    And (.Trigger <> eTrigger.POSINVALIDA Or IsPet) And Not HayAgua(Map, X, Y)
        Else
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.npcindex = 0) _
                    And (.Trigger <> eTrigger.POSINVALIDA Or IsPet)

        End If

    End With

End Function
 

Sub LookatTile(ByVal UserIndex As Integer, _
               ByVal Map As Integer, _
               ByVal X As Integer, _
               ByVal Y As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 26/03/2009
    '13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
    '***************************************************

    
    On Error GoTo LookatTile_Err

    'Responde al click del usuario sobre el mapa
    Dim FoundChar      As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex  As Integer

    With UserList(UserIndex)

        '¿Rango Visión? (ToxicWaste)
        If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then Exit Sub

    
        '¿Posicion valida?
        If InMapBounds(Map, X, Y) Then

            With .flags
                .TargetMap = Map
                .TargetX = X
                .TargetY = Y

                '¿Es un obj?
                If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y
                    FoundSomething = 1
                ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then

                    'Informa el nombre
                    If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y
                        FoundSomething = 1
                    End If

                ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then

                    If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = X + 1
                        .TargetObjY = Y + 1
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then

                    If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = X
                        .TargetObjY = Y + 1
                        FoundSomething = 1

                    End If

                End If
                
                 If FoundSomething = 1 Then
                    UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                End If

                '¿Es un personaje?
                If Y + 1 <= YMaxMapSize Then
                    If MapData(Map, X, Y + 1).UserIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).UserIndex
                        FoundChar = 1

                    End If

                    If MapData(Map, X, Y + 1).npcindex > 0 Then
                        TempCharIndex = MapData(Map, X, Y + 1).npcindex
                        FoundChar = 2

                    End If

                End If

                '¿Es un personaje?
                If FoundChar = 0 Then
                    If MapData(Map, X, Y).UserIndex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).UserIndex
                        FoundChar = 1

                    End If

                    If MapData(Map, X, Y).npcindex > 0 Then
                        TempCharIndex = MapData(Map, X, Y).npcindex
                        FoundChar = 2

                    End If

                End If

            End With
        
        Select Case FoundChar
        
        Case 1 '  ¿Encontro un Usuario?
            
        If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            'if UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo) = 0 Then
                If UserList(TempCharIndex).showName Then
                    WriteCharMsgStatus UserIndex, TempCharIndex
                End If
            'End If
 
         FoundSomething = 1
        .flags.TargetUser = TempCharIndex
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        
        End If
        
        Case 2 '¿Encontro un NPC?
        
        With .flags
        
        'If UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo) = 0 Then
           WriteCharMsgStatusNPC UserIndex, TempCharIndex
        'End If
        
         'Si el NPC se corre mandamos el texto sino lo lee el cliente :p
        If (Npclist(TempCharIndex).StartPos.X <> Npclist(TempCharIndex).Pos.X) Or (Npclist(TempCharIndex).StartPos.Y <> Npclist(TempCharIndex).Pos.Y) Then
            If Len(Npclist(TempCharIndex).desc) > 1 Then Call WriteChatOverHeadLocale(UserIndex, Npclist(TempCharIndex).Char.CharIndex, Npclist(TempCharIndex).Numero, 4)
        End If
         
         FoundSomething = 1
        .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
        .TargetNPC = TempCharIndex
        .TargetUser = 0
        .TargetObj = 0
        
        End With
        
        
        End Select
        
        With .flags
        If FoundChar = 0 Then
            .TargetNPC = 0
            .TargetNpcTipo = eNPCType.Comun
            .TargetUser = 0
        End If
        End With
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0

    End If
    
    Else
    
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        

    End If
    
    End If
 
    End With

    Exit Sub

LookatTile_Err:
Call RegistrarError(Err.Number, Err.description, "Extra.LookAtTile", Erl)
Resume Next
End Sub

Function FindDirection(ByVal NPCI As Integer, Target As WorldPos) As eHeading

''*****************************************************************
''Devuelve la direccion en la cual el target se encuentra
''desde pos, 0 si la direc es igual
''*****************************************************************
'


Dim X As Integer
Dim Y As Integer
Dim Pos As WorldPos
Dim puedeX As Boolean
Dim puedeY As Boolean

Pos = Npclist(NPCI).Pos
X = Npclist(NPCI).Pos.X - Target.X
Y = Npclist(NPCI).Pos.Y - Target.Y
'
'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If
'
''Lo tenemos al lado
If Distancia(Pos, Target) = 1 Then
    FindDirection = 0
    Exit Function
End If
'

If Rodeado(Target) Then
    FindDirection = 0
    Exit Function
End If
'
'
''Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    If Not PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y) Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.WEST: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If

''norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    If Not PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y) Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.WEST: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

''oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    If Not PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1) Then
                FindDirection = eHeading.SOUTH: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.WEST: Exit Function
    End If
End If

''este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    If Not PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1) Then
                FindDirection = eHeading.SOUTH: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.EAST: Exit Function
    End If
End If
'
''NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y - 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.WEST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If

'    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If
'
''NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y - 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.EAST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If

'    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If
'
''SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y + 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
       Else
            If puedeX Then
                FindDirection = eHeading.WEST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If

'    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

''SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y + 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
           End If
        Else
            If puedeX Then
                FindDirection = eHeading.EAST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If

    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

End Function
Function Rodeado(ByRef Pos As WorldPos) As Boolean
   
    If Not PuedeNpc(Pos.Map, Pos.X + 1, Pos.Y) Then
        If Not PuedeNpc(Pos.Map, Pos.X - 1, Pos.Y) Then
            If Not PuedeNpc(Pos.Map, Pos.X, Pos.Y + 1) Then
                If Not PuedeNpc(Pos.Map, Pos.X, Pos.Y - 1) Then
                    Rodeado = True
                End If
            End If
        End If
    End If
End Function
Function PuedeNpc(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo hayerror
    'Add Marius agregamos el if para la validacion asi no se va el indice a la mierda.
    If X > 0 And Y > 0 Then
        PuedeNpc = (MapData(Map, X, Y).npcindex = 0 And _
                    MapData(Map, X, Y).Blocked = 0 And _
                    MapData(Map, X, Y).UserIndex = 0)
    Else
        PuedeNpc = False
    End If
                
    Exit Function
hayerror:
     LogError ("Error en PuedeNPC:" & Err.Number & " Descripcion: " & Err.description & " map:" & Map & " x:" & X & " y:" & Y)
End Function
Function FindDonde(ByVal NPCI As Integer, Target As WorldPos) As eHeading
Dim X As Integer
Dim Y As Integer
Dim Pos As WorldPos

    Pos = Npclist(NPCI).Pos
    X = Npclist(NPCI).Pos.X - Target.X
    Y = Npclist(NPCI).Pos.Y - Target.Y
    
    If Sgn(X) = 0 And Sgn(Y) = 0 Then FindDonde = 0: Exit Function
    
    If Sgn(X) = 0 And Sgn(Y) = -1 Then FindDonde = eHeading.SOUTH
    If Sgn(X) = 0 And Sgn(Y) = 1 Then FindDonde = eHeading.NORTH
    If Sgn(X) = 1 And Sgn(Y) = 0 Then FindDonde = eHeading.WEST
    If Sgn(X) = -1 And Sgn(Y) = 0 Then FindDonde = eHeading.EAST

End Function
Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
ItemNoEsDeMapa = ObjData(index).OBJType <> eOBJType.otPuertas And ObjData(index).OBJType <> eOBJType.otCarteles And ObjData(index).OBJType <> eOBJType.otArboles And ObjData(index).OBJType <> eOBJType.otYacimiento And ObjData(index).OBJType <> eOBJType.otTeleport And ObjData(index).OBJType <> eOBJType.otCorreo And ObjData(index).OBJType <> eOBJType.otYacimiento And ObjData(index).OBJType <> eOBJType.otMuebles And ObjData(index).OBJType <> eOBJType.otFragua And ObjData(index).OBJType <> eOBJType.otYunque


End Function
 
Public Function NoTieneEspacioAmigos(ByVal Usuario As Integer) As Boolean
  Dim i As Long
  Dim count As Byte

  For i = 1 To MAXAMIGOS
  If Not UserList(Usuario).Amigos(i).Nombre = "Vacío" Then
  count = count + 1
  End If
  Next i

  If count = MAXAMIGOS Then
  NoTieneEspacioAmigos = True
  End If

End Function
Public Function BuscarSlotAmigoVacio(ByVal Usuario As Integer) As Byte
  Dim i As Long

  For i = 1 To MAXAMIGOS
  If UserList(Usuario).Amigos(i).Nombre = "Vacío" Then
  BuscarSlotAmigoVacio = i
  Exit Function
  End If
  Next i

End Function
Public Function BuscarSlotAmigoName(ByVal Usuario As Integer, ByVal Nombre As String) As Boolean
  Dim i As Long

  For i = 1 To MAXAMIGOS
  If UCase$(UserList(Usuario).Amigos(i).Nombre) = UCase$(Nombre) Then
  BuscarSlotAmigoName = True
  Exit Function
  End If
  Next i

End Function


Public Function BuscarSlotAmigoNameSlot(ByVal Usuario As Integer, ByVal Nombre As String) As Byte
  Dim i As Long

  For i = 1 To MAXAMIGOS
  If UCase$(UserList(Usuario).Amigos(i).Nombre) = UCase$(Nombre) Then
  BuscarSlotAmigoNameSlot = i
  Exit Function
  End If
  Next i

End Function
Public Sub delAmigoOfli(ByVal charName As String, ByVal Amigo As String)
  Dim CharFile As String
  Dim i As Byte
  Dim Cantidad As Byte
  Dim Amigos As String
  Dim Tiene As Boolean
  Dim slot As Byte
  CharFile = CharPath & charName & ".chr"
  Cantidad = GetVar(CharFile, "FLAGS", "CantidadAmigos")
If FileExist(CharFile) Then

    For i = 1 To Cantidad
    If UCase$(CStr(GetVar(CharFile, "AMIGOS", "NOMBRE" & i))) = UCase$(Amigo) Then
    Tiene = True
    Exit For
    End If
    Next i
    Debug.Print i & "tenemos y existe"
  If Tiene Then
    If i = Cantidad Then
     Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & i, "Vacío")
    Else
        For slot = i To Cantidad
                If slot = Cantidad Then Exit For
                Amigos = GetVar(CharFile, "AMIGOS", "NOMBRE" & slot + 1)
                Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & slot, Amigos)
                Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & slot + 1, "Vacío")
        Next slot
    End If
 Call WriteVar(CharFile, "FLAGS", "CantidadAmigos", Cantidad - 1)
 End If
  
End If
End Sub
Public Function IntentarAgregarAmigo(ByVal Usuario As Integer, ByVal Otro As Integer, ByRef razon As String) As Boolean
  With UserList(Usuario)
 If Otro = 0 Or Usuario = 0 Then
  razon = "Usuario offline."
  IntentarAgregarAmigo = False
  Exit Function

  ElseIf Usuario = Otro Then
  razon = "No puedes agregarte a tu propia lista de amigos."
  IntentarAgregarAmigo = False
  Exit Function
  
  ElseIf NoTieneEspacioAmigos(Usuario) = True Then
  razon = "La lista de amigos está llena."
  IntentarAgregarAmigo = False
  Exit Function

  ElseIf NoTieneEspacioAmigos(Otro) = True Then
  razon = "La lista de amigos del jugador está llena."
  IntentarAgregarAmigo = False
  Exit Function

  ElseIf BuscarSlotAmigoName(Usuario, UserList(Otro).Name) = True Then
  razon = UserList(Otro).Name & " ya está en tu lista de amigos"
  IntentarAgregarAmigo = False
  Exit Function
  End If

  IntentarAgregarAmigo = True
  End With
End Function
 Public Function ObtenerIndexLibre(ByVal Usuario As Integer) As Integer
  Dim i As Long

  For i = 1 To MAXAMIGOS
  If UserList(Usuario).Amigos(i).index <= 0 Then
  ObtenerIndexLibre = i
  Exit Function
  End If
  Next i

End Function
Public Function ObtenerIndexUsuado(ByVal Usuario As Integer, ByVal Otro As Integer) As Integer
  Dim i As Long

  For i = 1 To MAXAMIGOS
  If UserList(Usuario).Amigos(i).index = Otro Then
  ObtenerIndexUsuado = i
  Exit Function
  End If
  Next i

End Function

Public Sub ObtenerIndexAmigos(ByVal Usuario As Integer, ByVal Desconectar As Boolean)
    Dim i As Long
    Dim slot As Byte
    Dim Cantidad As String
    Dim tUser2 As Integer
    Dim j As Long
  With UserList(Usuario)

  If Desconectar = False Then
  For i = 1 To MAXAMIGOS
  If LenB(UserList(i).Name) > 0 Then
  If BuscarSlotAmigoName(Usuario, UserList(i).Name) Then
  'Lo encontro y agregamos el index
  slot = ObtenerIndexLibre(Usuario)
  'Por las dudas
  If slot > 0 Then _
  .Amigos(slot).index = i
  If BuscarSlotAmigoName(i, .Name) Then
  'Actualizamos la lista del otro
  slot = ObtenerIndexLibre(i)
  If slot > 0 Then
  UserList(i).Amigos(slot).index = Usuario
  'Informamos al otro de nuestra presencia
  Call WriteLocaleMsg(i, 239, .Name)
  'slot = BuscarSlotAmigoNameSlot(.Amigos(i).index, .name)
  ' Writemostrarubicacion .Amigos(i).index, UserList(Usuario).name, i, .Pos.map, .Pos.x, .Pos.y
  End If
  End If
  End If
  End If
  Next i
  Else
  For i = 1 To .flags.CantidadAmigos
  'Antes q nada
  If .Amigos(i).index > 0 Then
  Call WriteLocaleMsg(.Amigos(i).index, 240, .Name)
   slot = BuscarSlotAmigoNameSlot(.Amigos(i).index, .Name)
  Writemostrarubicacion .Amigos(i).index, UserList(Usuario).Name, slot, 0, 0, 0
  'Actualizamos la lista de index de los amigos
  slot = ObtenerIndexUsuado(.Amigos(i).index, Usuario)
    UserList(.Amigos(i).index).Amigos(slot).index = 0
  End If
  Next i
  End If
  End With
End Sub

 
Public Sub Sum(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal FX As Boolean)
On Error Resume Next
Dim NuevaPos As WorldPos
Dim FuturePos As WorldPos
    
    FuturePos.Map = Map
    FuturePos.X = X
    FuturePos.Y = Y
    
    If UserIndex <> 0 Then
        Call ClosestLegalPos(FuturePos, NuevaPos)
        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
            Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, FX)
        Else
            Call WarpUserChar(UserIndex, FuturePos.Map, FuturePos.X, FuturePos.Y, FX)
        End If
    End If
End Sub

Sub AddtoRichTextBox(Text As String)


    With frmRecibeDatos.RecTxt

    .SelFontName = "Tahoma"
    .SelFontSize = 8
    
    If (Len(.Text)) > 20000 Then .Text = vbNullString
    .SelStart = Len(frmRecibeDatos.RecTxt.Text)
    .SelLength = 0
    

    .SelBold = 0
    .SelItalic = 0
            
    .SelColor = RGB(200, 185, 10)
    
    .SelText = Text & vbCrLf
 
    End With
    
End Sub


