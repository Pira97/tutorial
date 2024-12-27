Attribute VB_Name = "AI"
'Argentum Online 0.12.2
'Copyright (C) 2002 Mï¿½rquez Pablo Ignacio
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
'Calle 3 nï¿½mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Cï¿½digo Postal 1900
'Pablo Ignacio Mï¿½rquez

Option Explicit

Public Enum TipoAI
    NpcMaloAtacaUsersBuenos = 1
    ESTATICO = 1
    MueveAlAzar = 3
    
    NPCDEFENSA = 4 'mascota
    NpcObjeto = 6 'aun no se
    SigueAmo = 8 'mascota
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visiï¿½n que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'                        Modulo AI_NPC
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'AI de los NPC
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½
'?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½

Private Sub GuardiasAI(ByVal npcindex As Integer, ByVal DelCaos As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    With Npclist(npcindex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                            'ï¿½ES CRIMINAL?
                            If Not DelCaos Then
                                If esRene(UI) Then
                                    If NpcAtacaUser(npcindex, UI) Then
                                        Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(npcindex, UI) Then
                                        Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            Else
                                If Not esRene(UI) Then
                                    If NpcAtacaUser(npcindex, UI) Then
                                        Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(npcindex, UI) Then
                                        Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If  'not inmovil
        Next headingloop
    End With
    
    Call RestoreOldMovement(npcindex)
End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal npcindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/01/2010 (ZaMa)
'28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
'14/09/200*: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'**************************************************************
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    Dim UserProtected As Boolean
    
    atacoPJ = False
    
    With Npclist(npcindex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            'If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.X, nPos.Y).npcindex
                    If UI > 0 And Not atacoPJ Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            
                            atacoPJ = True
                            If .Movement = NpcObjeto Then
                                ' Los npc objeto no atacan siempre al mismo usuario
                                If RandomNumber(1, 3) = 3 Then atacoPJ = False
                            End If
                            
                            If atacoPJ Then
                            If NpcAtacaUser(npcindex, UI) Then
                                If .flags.LanzaSpells <> 0 Then
                                    If .flags.AtacaDoble Then
                                        If (RandomNumber(0, 1)) Then
                                            If NpcAtacaUser(npcindex, UI) Then
                                            '.CanAttack = 1
                                                Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    
                                    Call NpcLanzaUnSpell(npcindex, UI)
                                    Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                   
                                Else
                                Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                End If
                            End If
                            
                                Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                            End If
                            Exit Sub

                       End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(npcindex, NPCI, False)
                            Exit Sub
                        End If
                    End If
                End If
            'End If  'inmo
        Next headingloop
    End With
    
    Call RestoreOldMovement(npcindex)
End Sub

Private Sub HostilBuenoAI(ByVal npcindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't atack protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
    Dim nPos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    Dim UserProtected As Boolean
    
    With Npclist(npcindex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UI).flags.Ignorado
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(npcindex, UI)
                                End If
                                
                                If NpcAtacaUser(npcindex, UI) Then
                                    Call ChangeNPCChar(npcindex, .Char.body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    Call RestoreOldMovement(npcindex)
End Sub

Private Sub IrUsuarioCercano(ByVal npcindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 12/01/2010 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
'***************************************************
    Dim tHeading As Byte
    Dim UserIndex As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    Dim UserProtected As Boolean
    
    With Npclist(npcindex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado
                        
                        If UserList(UserIndex).flags.Muerto = 0 Then
                            If Not UserProtected Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(npcindex, UserIndex)
                                Exit Sub
                            End If
                        End If
                        
                    End If
                End If
            Next i
            
        ' No esta inmobilizado
        Else
            
            ' Tiene prioridad de seguir al usuario al que le pertenece si esta en el rango de vision
            Dim OwnerIndex As Integer
            
            OwnerIndex = .Owner
            If OwnerIndex > 0 Then
            
                'Is it in it's range of vision??
                If Abs(UserList(OwnerIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(OwnerIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        ' va hacia el si o esta invi ni oculto
                        If UserList(OwnerIndex).flags.Invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.Ignorado Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(npcindex, OwnerIndex)
                                
                            tHeading = FindDirection(npcindex, UserList(OwnerIndex).Pos)
                            Call MoveNPCChar(npcindex, tHeading)
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
            
            ' No le pertenece a nadie o el dueño no esta en el rango de vision, sigue a cualquiera
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        With UserList(UserIndex)
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or .flags.Ignorado
                            
                            If .flags.Muerto = 0 And .flags.Invisible = 0 And .flags.Oculto = 0 And _
                                .flags.AdminPerseguible And Not UserProtected Then
                                
                                If Npclist(npcindex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(npcindex, UserIndex)
                                
                                tHeading = FindDirection(npcindex, .Pos)
                                Call MoveNPCChar(npcindex, tHeading)
                                Exit Sub
                            End If
                            
                        End With
                        
                    End If
                End If
            Next i
            
            'Si llega aca es que no habï¿½a ningï¿½n usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(npcindex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
            
        End If
    End With
    
    Call RestoreOldMovement(npcindex)
End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal npcindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: Marco Vanotti (MarKoxX)
'Last Modify Date: 08/16/2008
'08/16/2008: MarKoxX - Now pets that do melï¿½ attacks have to be near the enemy to attack.
'**************************************************************
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim i As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer

    With Npclist(npcindex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select

            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not esRene(.MaestroUser) And Not esRene(UI) And esArmada(.MaestroUser) Then
                                    'Call WriteMensajes(.MaestroUser, eMensajes.Mensaje020)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(npcindex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(npcindex).Pos) <= 1 Then
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(npcindex).Numero <> 92 Then
                                            Call NpcAtacaUser(npcindex, UI)
                                        End If
                                    End If
                                 End If
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not esRene(.MaestroUser) And Not esRene(UI) And (esArmada(.MaestroUser)) Then
                                    Call WriteLocaleMsg(.MaestroUser, 8)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(npcindex)
                                    Exit Sub
                                End If
                            End If
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(npcindex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(npcindex).Pos) <= 1 Then
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(npcindex).Numero <> 92 Then
                                            Call NpcAtacaUser(npcindex, UI)
                                        End If
                                    End If
                                 End If
                                 
                                 tHeading = FindDirection(npcindex, UserList(UI).Pos)
                                 Call MoveNPCChar(npcindex, tHeading)
                                 
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(npcindex)
End Sub

Private Sub RestoreOldMovement(ByVal npcindex As Integer)
    With Npclist(npcindex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        End If
    End With
End Sub


Private Sub SeguirAmo(ByVal npcindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tHeading As Byte
    Dim UI As Integer
    
    With Npclist(npcindex)
        If .Target = 0 And .TargetNPC = 0 Then
            UI = .MaestroUser

            If UI > 0 Then
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        If UserList(UI).flags.Muerto = 0 _
                                And UserList(UI).flags.Invisible = 0 _
                                And UserList(UI).flags.Oculto = 0 _
                                And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            tHeading = FindDirection(npcindex, UserList(UI).Pos)
                            Call MoveNPCChar(npcindex, tHeading)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(npcindex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal npcindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tHeading As Byte
    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With Npclist(npcindex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).npcindex
                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(npcindex, NI)
                                    If Npclist(NI).NPCtype = Dragon Then

                                        Call NpcLanzaUnSpellSobreNpc(NI, npcindex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(npcindex, NI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
                For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                       NI = MapData(.Pos.Map, X, Y).npcindex
                       If NI > 0 Then
                            If .TargetNPC = NI Then
                                 bNoEsta = True
                                 If .Numero = ELEMENTALFUEGO Then
                                     Call NpcLanzaUnSpellSobreNpc(npcindex, NI)
                                     If Npclist(NI).NPCtype = Dragon Then
 
                                        Call NpcLanzaUnSpellSobreNpc(NI, npcindex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(npcindex, NI)
                                    End If
                                 End If
                                 If .flags.Inmovilizado = 1 Then Exit Sub
                                 If .TargetNPC = 0 Then Exit Sub
                                 tHeading = FindDirection(npcindex, Npclist(MapData(.Pos.Map, X, Y).npcindex).Pos)
                                 Call MoveNPCChar(npcindex, tHeading)
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(npcindex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Public Sub AiNpcObjeto(ByVal npcindex As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 14/09/2009 (ZaMa)
'14/09/2009: ZaMa - Now npcs don't follow protected users.
'***************************************************
    Dim UserIndex As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim UserProtected As Boolean
    
    With Npclist(npcindex)
        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
            
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                    With UserList(UserIndex)
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                        
                        If .flags.Muerto = 0 And .flags.Invisible = 0 And _
                            .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            
                            ' No quiero que ataque siempre al primero
                            If RandomNumber(1, 3) < 3 Then
                                If Npclist(npcindex).flags.LanzaSpells > 0 Then
                                     Call NpcLanzaUnSpell(npcindex, UserIndex)
                                End If
                            
                                Exit Sub
                            End If
                        End If
                    End With
               End If
            End If
            
        Next i
    End With

End Sub

Sub NPCAI(ByVal npcindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: ZaMa
'Last Modify Date: 15/11/2009
'08/16/2008: MarKoxX - Now pets that do melï¿½ attacks have to be near the enemy to attack.
'15/11/2009: ZaMa - Implementacion de npc objetos ai.
'**************************************************************
On Error GoTo ErrorHandler
    With Npclist(npcindex)
    
    If .NPCtype = eNPCType.GuardiasCity Or .NPCtype = eNPCType.Convertidores Or .NPCtype = eNPCType.facciones Then Exit Sub   'Mermas temporal para que los bichos nose muevan
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then
            'Busca a alguien para atacar
            'ï¿½Es un guardia?
            If .Hostile <> 0 Then
                Call HostilMalvadoAI(npcindex)
            ElseIf .Hostile = 0 Then
                Call HostilBuenoAI(npcindex)
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement
            Case TipoAI.MueveAlAzar
                If .flags.Inmovilizado = 1 Then Exit Sub

                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(npcindex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    
            
            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(npcindex)
            
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(npcindex)
            

            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(npcindex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(npcindex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(npcindex)
                
            Case TipoAI.NpcObjeto
                Call AiNpcObjeto(npcindex)
                
            Case TipoAI.NpcPathfinding
                If .flags.Inmovilizado = 1 Then Exit Sub
                If ReCalculatePath(npcindex) Then
                    Call PathFindingAI(npcindex)
                    'Existe el camino?
                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(npcindex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                    End If
                Else
                    If Not PathEnd(npcindex) Then
                        Call FollowPath(npcindex)
                    Else
                        .PFINFO.PathLenght = 0
                    End If
                End If
        End Select
    End With
Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & Npclist(npcindex).Name & " " & Npclist(npcindex).MaestroUser & " " & Npclist(npcindex).MaestroNpc & " mapa:" & Npclist(npcindex).Pos.Map & " x:" & Npclist(npcindex).Pos.X & " y:" & Npclist(npcindex).Pos.Y & " Mov:" & Npclist(npcindex).Movement & " TargU:" & Npclist(npcindex).Target & " TargN:" & Npclist(npcindex).TargetNPC)
    Dim MiNPC As npc
    MiNPC = Npclist(npcindex)
    Call QuitarNPC(npcindex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function UserNear(ByVal npcindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'Returns True if there is an user adjacent to the npc position.
'***************************************************

    With Npclist(npcindex)
        UserNear = Not Int(Distance(.Pos.X, .Pos.Y, UserList(.PFINFO.TargetUser).Pos.X, _
                    UserList(.PFINFO.TargetUser).Pos.Y)) > 1
    End With
End Function

Function ReCalculatePath(ByVal npcindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'Returns true if we have to seek a new path
'***************************************************

    If Npclist(npcindex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(npcindex) And Npclist(npcindex).PFINFO.PathLenght = Npclist(npcindex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True
    End If
End Function

Function PathEnd(ByVal npcindex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'Returns if the npc has arrived to the end of its path
'***************************************************
    PathEnd = Npclist(npcindex).PFINFO.CurPos = Npclist(npcindex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal npcindex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'Moves the npc.
'***************************************************
    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    With Npclist(npcindex)
        tmpPos.Map = .Pos.Map
        tmpPos.X = .PFINFO.Path(.PFINFO.CurPos).Y ' invertï¿½ las coordenadas
        tmpPos.Y = .PFINFO.Path(.PFINFO.CurPos).X
        
        'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
        
        tHeading = FindDirection(npcindex, tmpPos)
        
        MoveNPCChar npcindex, tHeading
        
        .PFINFO.CurPos = .PFINFO.CurPos + 1
    End With
End Function

Function PathFindingAI(ByVal npcindex As Integer) As Boolean
'***************************************************
'Author: Gulfas Morgolock
'Last Modification: -
'This function seeks the shortest path from the Npc
'to the user's location.
'***************************************************
    Dim Y As Long
    Dim X As Long
    
    With Npclist(npcindex)
        For Y = .Pos.Y - 10 To .Pos.Y + 10    'Makes a loop that looks at
             For X = .Pos.X - 10 To .Pos.X + 10   '5 tiles in every direction
                
                 'Make sure tile is legal
                 If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                    
                     'look for a user
                     If MapData(.Pos.Map, X, Y).UserIndex > 0 Then
                         'Move towards user
                          Dim tmpUserIndex As Integer
                          tmpUserIndex = MapData(.Pos.Map, X, Y).UserIndex
                          With UserList(tmpUserIndex)
                            If .flags.Muerto = 0 And .flags.Invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
                                Npclist(npcindex).PFINFO.Target.X = .Pos.Y
                                Npclist(npcindex).PFINFO.Target.Y = .Pos.X 'ops!
                                Npclist(npcindex).PFINFO.TargetUser = tmpUserIndex
                                Call SeekPath(npcindex)
                                Exit Function
                            End If
                        End With
                    End If
                End If
            Next X
        Next Y
    End With
End Function

Sub NpcLanzaUnSpell(ByVal npcindex As Integer, ByVal UserIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: -
'Last Modify Date: -
'**************************************************************
    With UserList(UserIndex)
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
    End With
    
    Dim K As Integer
    K = RandomNumber(1, Npclist(npcindex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(npcindex, UserIndex, Npclist(npcindex).Spells(K))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal npcindex As Integer, ByVal TargetNPC As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim K As Integer
    K = RandomNumber(1, Npclist(npcindex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(npcindex, TargetNPC, Npclist(npcindex).Spells(K))
End Sub




