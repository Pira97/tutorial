Attribute VB_Name = "modHechizos"

Option Explicit
Public Const ANILLO_ESPECTRAL As Integer = 1329
Public Const ANILLO_PENUMBRAS As Integer = 1330

Public Sub NpcLanzaSpellSobreUser(ByVal npcindex As Integer, _
                           ByVal UserIndex As Integer, _
                           ByVal Spell As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 13/02/2009
    '13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
    '***************************************************
    On Error GoTo NpcLanzaSpellSobreUser_Err
    
100    If Not IntervaloPermiteAtacarNpc(npcindex) Then Exit Sub
        
102    If Spell = 0 Then Exit Sub
104    If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub
    
       ' Si no se peude usar magia en el mapa, no le deja hacerlo.
106    If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

        If UserList(UserIndex).Invent.MagicIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.MagicasNoAtacan Then
                Exit Sub
            End If
        End If
    
108    Dim daño As Integer


110    With UserList(UserIndex)

                If Hechizos(Spell).SubeHP = 1 Then
114                 daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
116                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.X, .Pos.Y))
118                 If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
120                 If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
  
122                 .Stats.MinHP = .Stats.MinHP + daño

124                 If .Stats.MinHP > .Stats.MaxHP Then .Stats.MinHP = .Stats.MaxHP

126                 Call WriteLocaleMsg(UserIndex, 34, "$" & Npclist(npcindex).Numero & "%" & daño)
128                 Call WriteUpdateHP(UserIndex)
            
130                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(Npclist(npcindex).Char.CharIndex, Spell, 1))
            
                ElseIf Hechizos(Spell).SubeHP = 2 Then

134                 daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
 
136                 If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
138                     daño = daño - ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).ResistenciaMagica
140                 End If
                
142                 If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
144                     daño = daño - ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica
146                 End If
        
148                 If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
150                     daño = daño - ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica
152                 End If
                
154                 If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
156                     daño = daño - ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica
158                 End If
            
166                 If .Invent.AnilloEqpObjIndex > 0 Then
168                     daño = daño - ObjData(.Invent.AnilloEqpObjIndex).ResistenciaMagica
170                 End If
            
172                 If daño < 0 Then daño = 0
            
174                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.X, .Pos.Y))
176                 If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
178                 If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
        
180                 .Stats.MinHP = .Stats.MinHP - daño

182                 Call WriteLocaleMsg(UserIndex, 34, "$" & Npclist(npcindex).Numero & "%" & daño)
                        
184                 Call WriteUpdateHP(UserIndex)
                
186                 Call SubirSkill(UserIndex, eSkill.Resistencia)
188                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(Npclist(npcindex).Char.CharIndex, Spell, 1))

                    'Muere
190                 If .Stats.MinHP < 1 Then
192                     .Stats.MinHP = 0
194                     Call UserDie(UserIndex)
204                 End If
        
        
208             ElseIf Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then

210                If .flags.Paralizado = 0 Then
212                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.X, .Pos.Y))
214                    If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
216                    If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
           
218                    If Hechizos(Spell).Inmoviliza = 1 Then .flags.Inmovilizado = 1
                  
220                    .flags.Paralizado = 1
224                    .Counters.Paralisis = IntervaloParalizado
                  
226                    Call WriteParalizeOK(UserIndex)
228                End If
    
232             ElseIf Hechizos(Spell).Estupidez = 1 Then   ' turbacion

234                If .flags.Estupidez = 0 Then
236                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, .Pos.X, .Pos.Y))
238                    If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
240                    If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
       
242                    .flags.Estupidez = 1
244                    .Counters.Ceguera = IntervaloInvisible
     
246                    Call WriteDumb(UserIndex)
248                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(Npclist(npcindex).Char.CharIndex, Spell, 1))
                    
250                End If

252             End If

    End With

    Exit Sub

NpcLanzaSpellSobreUser_Err:
254       Call RegistrarError(Err.Number, Err.description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreUser", Erl)
256     Resume Next
        
End Sub

Public Sub NpcLanzaSpellSobreNpc(ByVal npcindex As Integer, _
                          ByVal TargetNPC As Integer, _
                          ByVal Spell As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo NpcLanzaSpellSobreNpc_Err
    
    'solo hechizos ofensivos!
    If Not IntervaloPermiteAtacarNpc(npcindex) Then Exit Sub


    Dim daño As Integer

    Select Case Hechizos(Spell).SubeHP
    
    Case 2
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageEfectoCharParticula(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
    
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(npcindex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(npcindex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
    End Select
    
    If Hechizos(Spell).Inmoviliza = 1 Then
        If Npclist(TargetNPC).flags.AfectaParalisis = 0 And Npclist(TargetNPC).flags.Inmovilizado = 0 Then
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageEfectoCharParticula(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
            Npclist(TargetNPC).flags.Inmovilizado = 1
            Npclist(TargetNPC).flags.Paralizado = 0
            Npclist(TargetNPC).Contadores.Paralisis = IntervaloParalizado
        End If
    End If
    
    If Hechizos(Spell).Paraliza = 1 Then
        If Npclist(TargetNPC).flags.AfectaParalisis = 0 And Npclist(TargetNPC).flags.Paralizado = 0 Then
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).Loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageEfectoCharParticula(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, False))
            Npclist(TargetNPC).flags.Paralizado = 1
            Npclist(TargetNPC).flags.Inmovilizado = 0
            Npclist(TargetNPC).Contadores.Paralisis = IntervaloParalizado
        End If
    End If
    
    
    Exit Sub

NpcLanzaSpellSobreNpc_Err:
194     Call RegistrarError(Err.Number, Err.description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreNpc", Erl)

196     Resume Next
        
End Sub

Private Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next
    Exit Function

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.TieneHechizo", Erl)

196     Resume Next
        
End Function

Public Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal slot As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim hIndex As Integer
    Dim j      As Integer
    
    With UserList(UserIndex)
    
    hIndex = ObjData(.Invent.Object(slot).ObjIndex).HechizoIndex
    
        If Not TieneHechizo(hIndex, UserIndex) Then
        
            'Buscamos un slot vacio
            For j = 1 To MAXUSERHECHIZOS
                If .Stats.UserHechizos(j) = 0 Then Exit For
            Next j
            
            If .Stats.UserHechizos(j) <> 0 Then
                Call WriteLocaleMsg(UserIndex, 395)
            Else
                .Stats.UserHechizos(j) = hIndex
                Call UpdateUserHechizos(False, UserIndex, CByte(j))
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, CByte(slot), 1)
            End If
        Else
           Call WriteLocaleMsg(UserIndex, 396)
        End If

    End With

        
    Exit Sub

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.AgregarHechizo", Erl)

196     Resume Next
        
End Sub
            
Private Sub DecirPalabrasMagicas(ByVal spellindex As Integer, ByVal UserIndex As Integer)

    'Mermas 23/7/21 //Ahora mandamos solo el index del hechizo a los usuarios y que cada cpu lo procese a su velocidad
    
    On Error GoTo ErrHandler

    If UserList(UserIndex).Invent.MagicIndex <> 0 Then
        If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.Silencio Then
            Exit Sub
        End If
    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHeadLocale(UserList(UserIndex).Char.CharIndex, spellindex, 1))

    Exit Sub

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.DecirPalabrasMagicas", Erl)

196     Resume Next
        
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

    On Error GoTo ErrHandler

    With UserList(UserIndex)

    If DeadCheck(UserIndex) Then Exit Function
        
    If EsGm(UserIndex) Then
        PuedeLanzar = True
        Exit Function
    End If
    
    If .Stats.UserSkills(eSkill.magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteLocaleMsg(UserIndex, 221)
        PuedeLanzar = False
        Exit Function
    End If
    
    
    If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
        Call WriteLocaleMsg(UserIndex, 93)
        PuedeLanzar = False
        Exit Function
    End If
        
    If .Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido Then
        Call WriteLocaleMsg(UserIndex, 222)
        PuedeLanzar = False
        Exit Function
    End If

    If .Stats.UserSkills(eSkill.magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteLocaleMsg(UserIndex, 221)
        PuedeLanzar = False
        Exit Function
    End If
    
    
    If Hechizos(HechizoIndex).Anillo > 0 Then
    
        Select Case Hechizos(HechizoIndex).Anillo
        
        Case 1
        
        If Not TieneObjetos(ANILLO_ESPECTRAL, 1, UserIndex) Or TieneObjetos(ANILLO_PENUMBRAS, 1, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 440)
            PuedeLanzar = False
            Exit Function
        End If
        
        Case 2
        
        If Not TieneObjetos(ANILLO_PENUMBRAS, 1, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 441)
            PuedeLanzar = False
            Exit Function
        End If
        
        End Select
    
    End If
    
    End With
    
    PuedeLanzar = True

    Exit Function

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.PuedeLanzar", Erl)

196     Resume Next
        
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)

    On Error GoTo ErrHandler
    
1    Dim PosCasteadaX As Integer
2    Dim PosCasteadaY As Integer
3    Dim PosCasteadaM As Integer
4    Dim h As Integer
5    Dim TempX As Integer
6    Dim TempY As Integer
7    Dim TargetUser As Integer
8    Dim TargetNPC As Integer
9    Dim daño As Long
    
    With UserList(UserIndex)
10    PosCasteadaX = .flags.TargetX
11    PosCasteadaY = .flags.TargetY
12    PosCasteadaM = .flags.TargetMap
    
13    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
      
14   If Hechizos(h).HechizoDeArea Then
15        If Hechizos(h).AreaEfecto <> 0 Then
16            b = True
17            For TempX = PosCasteadaX - 1 To PosCasteadaX + 1
18                For TempY = PosCasteadaY - 1 To PosCasteadaY + 1
19                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
20                        TargetUser = MapData(PosCasteadaM, TempX, TempY).UserIndex
21                        If TargetUser > 0 Then

22                            If UserList(TargetUser).flags.Muerto = 0 Then
                                 
23                              Select Case Hechizos(h).SubeHP

                                Case 1 'cura
                                    If PuedeAyudar(UserIndex, TargetUser) Then
                                        Call HechizoTerrenoHP(UserIndex, TargetUser, h, True)
                                    End If
291
 
                                Case 2 'ataca
                                     If Not TargetUser = UserIndex Then
28                                        Call HechizoTerrenoHP(UserIndex, TargetUser, h, False)
                                          b = True
                                     End If
                                    
                                End Select
                                
                            End If
                        End If
                        
31                        TargetNPC = MapData(PosCasteadaM, TempX, TempY).npcindex
                        
32                        If TargetNPC <> 0 Then
33                            Call HechizoPropNPC(h, TargetNPC, UserIndex, b)
                              
                                If Npclist(TargetNPC).flags.Snd2 > 0 Then
                                    Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Npclist(TargetNPC).flags.Snd2, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
                                End If
                                
34                        End If
35                    End If
36                Next TempY
37            Next TempX
38        End If
39    End If
    
40    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
          b = True
42        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
43            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
44                If InMapBounds(PosCasteadaM, TempX, TempY) Then
45                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
46                        'hay un user
47                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
48                            If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).Loops))
49                            If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False, True))
50                        End If
51                    End If
52                End If
53            Next TempY
54        Next TempX
    
55        Call InfoHechizo(UserIndex)
56    End If
    
    
57    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoTerrenoParticula(Hechizos(h).Particle, PosCasteadaX, PosCasteadaY, Hechizos(h).TimeParticula))
58    If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoTerrenoFX(Hechizos(h).FXgrh, PosCasteadaX, PosCasteadaY + 1, Hechizos(h).Loops))
59    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, PosCasteadaX, PosCasteadaY))
      Call DecirPalabrasMagicas(h, UserIndex)
'
    End With

    Exit Sub

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoEstado", Erl)

196     Resume Next
        
End Sub


Sub HechizoTerrenoHP(ByVal UserIndex As Integer, ByVal TargetUser As Integer, ByVal h As Integer, ByVal Accion As Boolean)
    
    On Error GoTo Error
    
    Dim daño As Long
        
    If Accion = True Then  ' Sube HP
    
        daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MinHP + daño
        If UserList(TargetUser).Stats.MinHP > UserList(TargetUser).Stats.MaxHP Then UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MaxHP
        Call WriteUpdateHP(TargetUser)
        
        If UserIndex <> TargetUser Then
            Call WriteLocaleMsg(UserIndex, 445, daño & "%" & UserList(TargetUser).Name)
            Call WriteLocaleMsg(TargetUser, 32, UserList(TargetUser).Name & "%" & daño)
        Else
            Call WriteLocaleMsg(UserIndex, 33, daño)
        End If

        Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 3)
        

        Call FlushBuffer(UserIndex):   If UserIndex <> TargetUser Then Call FlushBuffer(TargetUser)
        

    Else 'Baja HP
    
        If TriggerZonaPelea(UserIndex, TargetUser) Then
    
            daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
            
            daño = daño + Porcentaje(daño, 2 * UserList(UserIndex).Stats.ELV)
            
            'Baculos DM + X
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                    daño = daño + (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
                End If
            End If
            
            If (UserList(TargetUser).Invent.CascoEqpObjIndex > 0) Then _
                daño = daño - ObjData(UserList(TargetUser).Invent.CascoEqpObjIndex).ResistenciaMagica
                
            If UserList(TargetUser).Invent.EscudoEqpObjIndex > 0 Then _
                daño = daño - ObjData(UserList(TargetUser).Invent.EscudoEqpObjIndex).ResistenciaMagica

            If UserList(TargetUser).Invent.ArmourEqpObjIndex > 0 Then _
                daño = daño - ObjData(UserList(TargetUser).Invent.ArmourEqpObjIndex).ResistenciaMagica

            If UserList(TargetUser).Invent.MonturaObjIndex > 0 Then _
                daño = daño - ObjData(UserList(TargetUser).Invent.MonturaObjIndex).ResistenciaMagica
            
            If UserList(TargetUser).Invent.AnilloEqpObjIndex > 0 Then _
                daño = daño - ObjData(UserList(TargetUser).Invent.AnilloEqpObjIndex).ResistenciaMagica
            
            
            If daño < 0 Then daño = 0
            
            If Not PuedeAtacar(UserIndex, TargetUser) Then Exit Sub
            
            If UserIndex <> TargetUser Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetUser)
            End If
            

            UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MinHP - daño
            
            Call WriteUpdateHP(TargetUser)
            
369         Call InfoHechizo(UserIndex)

390         Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 2) 'Dibuja daño

371         Call WriteLocaleMsg(TargetUser, 34, UserList(UserIndex).Name & "%" & daño) '%5 te ha quitado #1 puntos de vida

            'Muere
            If UserList(TargetUser).Stats.MinHP < 1 Then
                Call ContarMuerte(TargetUser, UserIndex)
                UserList(TargetUser).Stats.MinHP = 0
                Call ActStats(TargetUser, UserIndex)
                Call UserDie(TargetUser)
            End If
        End If
        
        If Hechizos(h).Envenena > 0 Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TargetUser)
            UserList(TargetUser).flags.Envenenado = Hechizos(h).Envenena
        End If
  
        Call FlushBuffer(UserIndex):   If UserIndex <> TargetUser Then Call FlushBuffer(TargetUser)
        
    End If
    
    Exit Sub

Error:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoHP", Erl)

196     Resume Next
        
End Sub


''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Author: Uknown
    'Last modification: 18/11/2009
    'Sale del sub si no hay una posición valida.
    '18/11/2009: Optimizacion de codigo.
    '***************************************************

    On Error GoTo Error

    With UserList(UserIndex)

        'No permitimos se invoquen criaturas en zonas seguras
        If MapInfo(.Pos.Map).Pk = False Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
            Exit Sub

        End If
    
        Dim spellindex As Integer, NroNpcs As Integer, npcindex As Integer, PetIndex As Integer
        Dim TargetPos  As WorldPos
    
        TargetPos.Map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
    
        spellindex = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
        ' Warp de mascotas
        If Hechizos(spellindex).Warp = 1 Then
            PetIndex = FarthestPet(UserIndex)
        
            ' La invoco cerca mio
            If PetIndex > 0 Then
                Call WarpMascota(UserIndex, PetIndex)

            End If
        
            ' Invocacion normal
        Else

            If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        
            For NroNpcs = 1 To Hechizos(spellindex).cant
            
                If .NroMascotas < MAXMASCOTAS Then
                    npcindex = SpawnNpc(Hechizos(spellindex).NumNpc, TargetPos, True, False)

                    If npcindex > 0 Then
                        .NroMascotas = .NroMascotas + 1
                    
                        PetIndex = FreeMascotaIndex(UserIndex)
                    
                        .MascotasIndex(PetIndex) = npcindex
                        .MascotasType(PetIndex) = Npclist(npcindex).Numero
                    
                        With Npclist(npcindex)
                            .MaestroUser = UserIndex
                            .Contadores.TiempoExistencia = IntervaloInvocacion
                            .GiveGLD = 0

                        End With
                    
                        Call FollowAmo(npcindex)
                    Else
                        Exit Sub

                    End If

                Else
                    Exit For

                End If
        
            Next NroNpcs

        End If

    End With

    Call InfoHechizo(UserIndex)
    HechizoCasteado = True

    Exit Sub

Error:

    With UserList(UserIndex)
        LogError ("[" & Err.Number & "] " & Err.description & " por el usuario " & .Name & "(" & UserIndex & ") en (" _
                & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & Hechizos( _
                spellindex).Nombre & "(" & spellindex & ") en la posicion ( " & .flags.TargetX & ", " & _
                .flags.TargetY & ")")

    End With

End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal spellindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    
    If SeguroCheck(UserIndex, 1) Then Exit Sub
    
    Dim HechizoCasteado As Boolean
    
    
    Select Case Hechizos(spellindex).Tipo

        Case TipoHechizo.uInvocacion
            Call WriteLocaleMsg(UserIndex, 435)
            Call HechizoInvocacion(UserIndex, HechizoCasteado)
            
        Case TipoHechizo.uEstado, uPropiedades
            Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)
        
        Case TipoHechizo.uCreateTelep
            Call HechizoCreateTelep(UserIndex, HechizoCasteado)
            
        Case TipoHechizo.uMaterializa
            'Call HechizoTerrenoMaterializa(UserIndex, HechizoCasteado)

        Case TipoHechizo.uDetectarInvis
            ' Call HechizoDetectaInvis(UserIndex, HechizoCasteado)
            
    End Select
    
     

    If HechizoCasteado Then

        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.magia)
 
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - Hechizos(spellindex).ManaRequerido

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(spellindex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateSta(UserIndex)
            Call WriteUpdateMana(UserIndex)

        End With

    End If
    
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal spellindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
    '***************************************************
    
    On Error GoTo ErrHandler
    
    If SeguroCheck(UserIndex, 1) Then Exit Sub
    
    Dim HechizoCasteado As Boolean

    Select Case Hechizos(spellindex).Tipo

        Case TipoHechizo.uEstado, TipoHechizo.uPropEsta, TipoHechizo.uPropiedades
            Call HechizoEstadoUsuario(UserIndex, HechizoCasteado, 0)
       
        Case TipoHechizo.uCreateMagic
            'Call HechizoCreateMagic(UserIndex, HechizoCasteado)
        
    End Select
    
    
    If HechizoCasteado Then

        With UserList(UserIndex)
            
            Call SubirSkill(UserIndex, eSkill.magia)
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - Hechizos(spellindex).ManaRequerido

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(spellindex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateSta(UserIndex)
            Call WriteUpdateMana(UserIndex)
            .flags.TargetUser = 0
            
        End With

    End If

    Exit Sub
    

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoUsuario", Erl)

196     Resume Next
        
End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
    '17/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Bonificacion para druidas de 10% para todos hechizos excepto apoca y descarga.
    '12/01/2010: ZaMa - Los druidas mimetizados con npcs ahora son ignorados.
    '***************************************************
    
    On Error GoTo ErrHandler
    
1    Dim HechizoCasteado As Boolean
 
    
    With UserList(UserIndex)

2        Select Case Hechizos(HechizoIndex).Tipo

            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, UserIndex)
                
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, UserIndex, HechizoCasteado)

        End Select
        
3        If HechizoCasteado Then
        
4            Call SubirSkill(UserIndex, eSkill.magia)

6            .flags.TargetNPC = 0
   
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - Hechizos(HechizoIndex).ManaRequerido

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
7            Call WriteUpdateSta(UserIndex)
8            Call WriteUpdateMana(UserIndex)
            
        End If

    End With

    Exit Sub

ErrHandler:
194     Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoNPC", Erl)

196     Resume Next
End Sub


Sub LanzarHechizo(ByVal spellindex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 02/16/2010
    '24/01/2007 ZaMa - Optimizacion de codigo.
    '02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
    '***************************************************
    On Error GoTo ErrHandler
    
    spellindex = UserList(UserIndex).Stats.UserHechizos(spellindex)

    With UserList(UserIndex)
            
        If spellindex = 0 Then Exit Sub
        
        If PuedeLanzar(UserIndex, spellindex) Then
 
            Select Case Hechizos(spellindex).Target
                 
                Case TargetType.uUsuarios '1

                    If .flags.TargetUser > 0 Then
                        If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            Call HandleHechizoUsuario(UserIndex, spellindex)
                        Else
                            Call WriteLocaleMsg(UserIndex, 8)
                        End If
                    Else
                        Call WriteLocaleMsg(UserIndex, 223)
                    End If
            
                Case TargetType.uNPC '2

                    If .flags.TargetNPC > 0 Then
                        If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            Call HandleHechizoNPC(UserIndex, spellindex)
                        Else
                            Call WriteLocaleMsg(UserIndex, 8)
                        End If
                    Else
                        Call WriteLocaleMsg(UserIndex, 224)
                    End If
            
                Case TargetType.uUsuariosYnpc '3

                    If .flags.TargetUser > 0 Then
                        If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            Call HandleHechizoUsuario(UserIndex, spellindex)
                        Else
                            Call WriteLocaleMsg(UserIndex, 8)
                        End If

                    ElseIf .flags.TargetNPC > 0 Then

                        If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            Call HandleHechizoNPC(UserIndex, spellindex)
                        Else
                            Call WriteLocaleMsg(UserIndex, 8)
                        End If
                    End If
            
                Case TargetType.uTerreno '4
                    Call HandleHechizoTerreno(UserIndex, spellindex)
            End Select
        
        End If
    
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
    
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

    End With

    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.description, "modHechizos.LanzarHechizo", Erl)
    Resume Next
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean, ByVal Hechizo As Integer)


    On Error GoTo ErrorHandler
    
    Dim HechizoIndex As Integer, targetIndex  As Integer
    Dim daño As Long
    Dim color As Byte
    color = 2
    
1    With UserList(UserIndex)
         If Hechizo > 0 Then
2           HechizoIndex = Hechizo
         Else
3           HechizoIndex = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
        End If
        
        targetIndex = .flags.TargetUser
4
5        If DeadCheck(UserIndex) Then
6            HechizoCasteado = False
7            Exit Sub
8        End If
            
9        ' <-------- Revive ---------->
10        If Hechizos(HechizoIndex).Revivir = 1 Then
11            If UserList(targetIndex).flags.Muerto = 1 Then
12
13                If MapInfo(UserList(targetIndex).Pos.Map).ResuSinEfecto > 0 Then
14                    Call WriteLocaleMsg(UserIndex, 448)
15                    HechizoCasteado = False
16                    Exit Sub
17                End If
        

18                If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
19                    If Not PuedeAyudar(UserIndex, targetIndex) Then
20                        HechizoCasteado = False
21                        Exit Sub
22                    End If
23                End If

                'Add Nod Kopfnickend Solo revive si no esta en modo combate
24                If UserList(targetIndex).flags.ModoCombate Then
25                    Call WriteLocaleMsg(UserIndex, 44)
26                    Call WriteLocaleMsg(targetIndex, 450, UserList(UserIndex).Name)
27                    HechizoCasteado = False
28                    Exit Sub
29                End If
                
30                Call RevivirUsuario(targetIndex)
                
31                UserList(targetIndex).flags.Sed = 0
32                UserList(targetIndex).flags.Hambre = 0
33                UserList(targetIndex).Stats.MinAGU = UserList(targetIndex).Stats.MaxAGU
34                UserList(targetIndex).Stats.MinHam = UserList(targetIndex).Stats.MaxHam
35
36                Call WriteUpdateHungerAndThirst(targetIndex)
                
37                If HechizoIndex = (54) Then
38                    UserList(targetIndex).Stats.MinMAN = UserList(targetIndex).Stats.MaxMAN
39                    UserList(targetIndex).Stats.MinSta = UserList(targetIndex).Stats.MaxSta
40                    Call WriteUpdateMana(targetIndex)
41                    Call WriteUpdateSta(targetIndex)
42                End If
43
44                If HechizoIndex > 0 Then
45                  Call InfoHechizo(UserIndex, HechizoIndex)
                  Else
                    Call InfoHechizo(UserIndex)
                  End If
                  HechizoCasteado = True
47                Exit Sub
48
49            Else
50
51                HechizoCasteado = False
52
53            End If
54
55        End If

58        If UserList(targetIndex).flags.Muerto = 1 Then

59            Call WriteLocaleMsg(UserIndex, 7) ' ¡El usuario está muerto!
               
60            HechizoCasteado = False
61            Exit Sub
62        End If
        
        
        ' <-------- Sanacion ---------->
63        If Hechizos(HechizoIndex).Sanacion = 1 Then
        
            
64            If Not PuedeAyudar(UserIndex, targetIndex) Then
65                HechizoCasteado = False
66                Exit Sub
67            End If
            
68            If UserList(targetIndex).flags.Incinerado Then UserList(targetIndex).flags.Incinerado = 0
            
69            If UserList(targetIndex).flags.Envenenado Then UserList(targetIndex).flags.Envenenado = 0
  
70            HechizoCasteado = True
            
71        End If
 
 
            
        ' <-------- Revive Familiar ----------> 'Solo damos la animacion xd
72        If Hechizos(HechizoIndex).ResucitaFamiliar = 1 Then
73            If UserIndex = targetIndex Then
74                    HechizoCasteado = True
75            End If
76        End If
        
        ' <-------- Agrega Invisibilidad ---------->
77        If Hechizos(HechizoIndex).Invisibilidad = 1 Then
            
78            If .flags.Invisible = True Then
79                Call WriteLocaleMsg(UserIndex, 68)
80                HechizoCasteado = False
81                Exit Sub
82            End If

83            If MapInfo(.Pos.Map).InviSinEfecto > 0 Then
84                Call WriteLocaleMsg(UserIndex, 448)
85                HechizoCasteado = False
86                Exit Sub
87            End If
            
88            .flags.Invisible = 1
89            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
90            HechizoCasteado = True

91        End If
        

        ' <-------- Agrega Envenenamiento ---------->
92        If Hechizos(HechizoIndex).Envenena > 0 Then
96            If UserIndex = targetIndex Then
95               Call WriteLocaleMsg(UserIndex, 298)
94               HechizoCasteado = False
93               Exit Sub
97            End If
        
98            If UserList(targetIndex).flags.Envenenado = 1 Then
99                Call WriteLocaleMsg(UserIndex, 326)
100                HechizoCasteado = False
101                Exit Sub
102            End If
            
106            If Not PuedeAtacar(UserIndex, targetIndex) Then
103                HechizoCasteado = False
104                Exit Sub
105            End If
            
107            If UserIndex <> targetIndex Then
108                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
109            End If

110            UserList(targetIndex).flags.Envenenado = 1
            
111            HechizoCasteado = True

112        End If
        
        

        ' <-------- Agrega Incineramiento ---------->
113        If Hechizos(HechizoIndex).Incinera = 1 Then
        
114            If UserIndex = targetIndex Then
115               Call WriteLocaleMsg(UserIndex, 298)
116               HechizoCasteado = False
117               Exit Sub
118            End If
            
119            If UserList(targetIndex).flags.Incinerado = 1 Then
120                Call WriteLocaleMsg(UserIndex, 326)
121                HechizoCasteado = False
122                Exit Sub
123            End If
            
124            If Not PuedeAtacar(UserIndex, targetIndex) Then
125                HechizoCasteado = False
126                Exit Sub
127            End If
            
128            If UserIndex <> targetIndex Then
129                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
130            End If

131            UserList(targetIndex).flags.Incinerado = 1
            
132            HechizoCasteado = True

133        End If
        
        ' <-------- Cura Envenenamiento ---------->
134        If Hechizos(HechizoIndex).CuraVeneno = 1 Then
        
135            If UserList(targetIndex).flags.Envenenado = 0 Then
136                Call WriteLocaleMsg(UserIndex, 451)
137                HechizoCasteado = False
138                Exit Sub
139            End If
            
            'Para poder tirar curar veneno a un pk en el ring
140            If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
141                If Not PuedeAyudar(UserIndex, targetIndex) Then
142                    HechizoCasteado = True
143                    Exit Sub
144                End If
145            End If

146            UserList(targetIndex).flags.Envenenado = 0
147            HechizoCasteado = True

148        End If
    
    
        ' <-------- Inmovilizar o Paralizar ---------->
149        If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
        
150            If UserIndex = targetIndex Then
151               Call WriteLocaleMsg(UserIndex, 298)
152               HechizoCasteado = False
153               Exit Sub
154            End If
            
155            If UserList(targetIndex).flags.Paralizado = 0 And UserList(targetIndex).flags.Inmovilizado = 0 Then
            
156                If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
                
157                If UserIndex <> targetIndex Then
158                    Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
159                End If
                
160                HechizoCasteado = True
                
161                If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(targetIndex).flags.Inmovilizado = 1
162                'UserList(targetIndex).flags.Paralizado = 1
                   If Hechizos(HechizoIndex).Paraliza = 1 Then UserList(targetIndex).flags.Paralizado = 1
 
163                UserList(targetIndex).Counters.Paralisis = IntervaloParalizado
164                UserList(targetIndex).flags.ParalizedByIndex = UserIndex
165                UserList(targetIndex).flags.ParalizedBy = UserList(UserIndex).Name
166                Call WriteParalizeOK(targetIndex)
167            End If
            
168         End If
        
   
    
        ' <-------- Remueve Paralisis/Inmovilidad ---------->
169        If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
            ' Remueve si esta en ese estado
170            If UserList(targetIndex).flags.Paralizado = 1 Or UserList(targetIndex).flags.Inmovilizado = 1 Then
                
                'Para poder tirar remo a un pk en el ring
171                If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
172                    If Not PuedeAyudar(UserIndex, targetIndex) Then
173                        HechizoCasteado = False
174                        Exit Sub
175                    End If
176                End If
                
177                Call RemoveParalisis(targetIndex)
                
178                HechizoCasteado = True
            
179            End If

180        End If
        
     
        ' <-------- Remueve Estupidez (Aturdimiento) ---------->
181        If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then

            ' Remueve si esta en ese estado
182            If UserList(targetIndex).flags.Estupidez = 1 Then
                
                'Para poder tirar remo a un pk en el ring
183                If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
184                    If Not PuedeAyudar(UserIndex, targetIndex) Then
185                        HechizoCasteado = False
186                        Exit Sub
187                    End If
188                End If
 
189                UserList(targetIndex).flags.Estupidez = 0
                
190                HechizoCasteado = True
                
                'no need to crypt this
191                Call WriteDumbNoMore(targetIndex)
192            End If

193        End If
 
        
        ' <-------- Agrega Estupidez (Aturdimiento) ---------->
194        If Hechizos(HechizoIndex).Estupidez = 1 Then
        
195            If UserIndex = targetIndex Then
196               Call WriteLocaleMsg(UserIndex, 298)
197               HechizoCasteado = False
198               Exit Sub
199            End If
200
201            If Not PuedeAtacar(UserIndex, targetIndex) Then
202                HechizoCasteado = False
203                Exit Sub
204            End If
            
205            If UserIndex <> targetIndex Then
206                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
207            End If

208            If UserList(targetIndex).flags.Estupidez = 0 Then
209                UserList(targetIndex).flags.Estupidez = 1
210                UserList(targetIndex).Counters.Estupidez = IntervaloParalizado / 3
211
212                Call WriteDumb(targetIndex)
213            End If

214            HechizoCasteado = True
            
215        End If
        
        ' <-------- Agrega Ceguera ---------->
216        If Hechizos(HechizoIndex).Ceguera = 1 Then
        
217            If UserIndex = targetIndex Then
218               Call WriteLocaleMsg(UserIndex, 298)
219               HechizoCasteado = False
220               Exit Sub
221            End If
        
222            If Not PuedeAtacar(UserIndex, targetIndex) Then
223                HechizoCasteado = False
224                Exit Sub
225            End If
            
226            If UserIndex <> targetIndex Then
227                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
228            End If
    
229            If UserList(targetIndex).flags.Ceguera = 0 Then
230                UserList(targetIndex).flags.Ceguera = 1
231                UserList(targetIndex).Counters.Ceguera = IntervaloParalizado / 3
        
232                Call WriteBlind(targetIndex)

            
233            End If
            
234            HechizoCasteado = True

235        End If
        
        
       ' AGILIDAD
236       Select Case Hechizos(HechizoIndex).SubeAgilidad
       
           Case 1 'Aumenta

            'Para poder tirar cl a un pk en el ring
237            If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
238                If Not PuedeAyudar(UserIndex, targetIndex) Then
239                    HechizoCasteado = False
240                    Exit Sub
241                End If
242            End If

243            daño = RandomNumber(Hechizos(HechizoIndex).MinAgilidad, Hechizos(HechizoIndex).MaxAgilidad)
            
244            UserList(targetIndex).flags.DuracionEfecto = 1200
245            UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) + daño
246            If UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
            
247            Call WriteUpdateDexterity(targetIndex)
            
248            UserList(targetIndex).flags.TomoPocion = True
249            HechizoCasteado = True
        
        Case 2 ' Reduce
        
250            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
        
251            If UserIndex <> targetIndex Then Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        
252            UserList(targetIndex).flags.TomoPocion = True
253            daño = RandomNumber(Hechizos(HechizoIndex).MinAgilidad, Hechizos(HechizoIndex).MaxAgilidad)
254            UserList(targetIndex).flags.DuracionEfecto = 700
255            UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) - daño

256            If UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(targetIndex).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        
258            Call WriteUpdateDexterity(targetIndex)
257            HechizoCasteado = True
            
        End Select

        
        ' FUERZA
259        Select Case Hechizos(HechizoIndex).SubeFuerza
        
        Case 1 'Aumenta
    
            'Para poder tirar fuerza a un pk en el ring
            If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
                If Not PuedeAyudar(UserIndex, targetIndex) Then
                    HechizoCasteado = False
                    Exit Sub
                End If
            End If
            
            daño = RandomNumber(Hechizos(HechizoIndex).MinFuerza, Hechizos(HechizoIndex).MaxFuerza)
        
            UserList(targetIndex).flags.DuracionEfecto = 1200
    
            UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) + daño
            If UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
        
            UserList(targetIndex).flags.TomoPocion = True
            Call WriteUpdateStrenght(targetIndex)
            
            HechizoCasteado = True
            
        Case 2 'Reduce
        
            If Not PuedeAtacar(UserIndex, targetIndex) Then
                HechizoCasteado = False
                Exit Sub
            End If
        
            If UserIndex <> targetIndex Then Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
        
            UserList(targetIndex).flags.TomoPocion = True
        
             daño = RandomNumber(Hechizos(HechizoIndex).MinFuerza, Hechizos(HechizoIndex).MaxFuerza)
            UserList(targetIndex).flags.DuracionEfecto = 700
            UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) - daño

            If UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(targetIndex).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        
            Call WriteUpdateStrenght(targetIndex)
            HechizoCasteado = True
            
        End Select
        
300        ' VIDA
      Select Case Hechizos(HechizoIndex).SubeHP
        
        Case 1 'Aumenta
        
302            'Para poder tirar curar a un pk en el ring
303            If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then
                If Not PuedeAyudar(UserIndex, targetIndex) Then
305                    HechizoCasteado = False
306                    Exit Sub
370                  End If
308            End If
            
309            daño = RandomNumber(Hechizos(HechizoIndex).MinHP, Hechizos(HechizoIndex).MaxHP)
310            daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)

311            UserList(targetIndex).Stats.MinHP = UserList(targetIndex).Stats.MinHP + daño

312            If UserList(targetIndex).Stats.MinHP > UserList(targetIndex).Stats.MaxHP Then UserList(targetIndex).Stats.MinHP = UserList(targetIndex).Stats.MaxHP
        
313            Call WriteUpdateHP(targetIndex)
               
321            Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 3)

               color = 13
               
322            HechizoCasteado = True
            
        Case 2 ' Reduce
        
325            If UserIndex = targetIndex Then
326                Call WriteLocaleMsg(UserIndex, 298) 'Objetivo invalido
327                HechizoCasteado = False
328                Exit Sub
329            End If

               daño = RandomNumber(Hechizos(HechizoIndex).MinHP, Hechizos(HechizoIndex).MaxHP)
330            daño = daño + Porcentaje(daño, 2 * UserList(UserIndex).Stats.ELV) 'esto ya tiene q estar
        
                
                'Baculos DM + X
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                        daño = daño + (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
                    End If
                End If
            
334            If UserList(targetIndex).Invent.CascoEqpObjIndex > 0 Then
335                daño = daño - ObjData(UserList(targetIndex).Invent.CascoEqpObjIndex).ResistenciaMagica
336            End If

337            If UserList(targetIndex).Invent.EscudoEqpObjIndex > 0 Then
338                daño = daño - ObjData(UserList(targetIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica
                End If
                
331            If UserList(targetIndex).Invent.ArmourEqpObjIndex > 0 Then
332                daño = daño - ObjData(UserList(targetIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica
333            End If
 
339            If UserList(targetIndex).Invent.MonturaObjIndex > 0 Then
340                daño = daño - ObjData(UserList(targetIndex).Invent.MonturaObjIndex).ResistenciaMagica
341            End If

357            If (.Invent.AnilloEqpObjIndex > 0) Then
358                daño = daño - ObjData(.Invent.AnilloEqpObjIndex).ResistenciaMagica
359            End If
        
360            If daño < 0 Then daño = 0
        
361            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
        
362            If UserIndex <> targetIndex Then
363                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
364            End If
            
365            HechizoCasteado = True
            
366            Call SubirSkill(targetIndex, eSkill.Resistencia)

367            UserList(targetIndex).Stats.MinHP = UserList(targetIndex).Stats.MinHP - daño
        
368            Call WriteUpdateHP(targetIndex)

390            Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 2) 'Dibuja daño

               'Muere
372            If UserList(targetIndex).Stats.MinHP < 1 Then
376                Call ContarMuerte(targetIndex, UserIndex)
            
378                UserList(targetIndex).Stats.MinHP = 0
379                Call ActStats(targetIndex, UserIndex)
380                Call UserDie(targetIndex)
384            End If
            
        End Select
        
        
383     If HechizoCasteado Then
            If HechizoIndex > 0 Then
                Call InfoHechizo(UserIndex, HechizoIndex, daño, color)
            Else
                Call InfoHechizo(UserIndex, 0, daño, color)
            End If
        End If
        
        Call FlushBuffer(UserIndex)
        
        If UserIndex <> targetIndex Then Call FlushBuffer(targetIndex)
        
    End With

Exit Sub

ErrorHandler:
    Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoUsuario", Erl)
    Resume Next
    
    
End Sub

Sub HechizoEstadoNPC(ByVal npcindex As Integer, ByVal spellindex As Integer, ByRef HechizoCasteado As Boolean, ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 07/07/2008
    'Handles the Spells that afect the Stats of an NPC
    '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
    'removidos por users de su misma faccion.
    '07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
    '***************************************************
    
    On Error GoTo HechizoEstadoNPC_Err
    
    With Npclist(npcindex)
        
 
        If Hechizos(spellindex).Paraliza = 1 Or Hechizos(spellindex).Inmoviliza = 1 Then
        
            If .flags.AfectaParalisis = 0 Then
            
                If Not PuedeAtacarNPC(UserIndex, npcindex, True) Then
                    HechizoCasteado = False
                    Exit Sub
                End If

                Call NPCAtacado(npcindex, UserIndex)
                
                If Hechizos(spellindex).Paraliza = 1 Then
                    .flags.Paralizado = 1
                End If
                
                If Hechizos(spellindex).Inmoviliza = 1 Then
                    .flags.Inmovilizado = 1
                End If
                
                .Contadores.Paralisis = IntervaloParalizado
                HechizoCasteado = True
            Else
            
                Call WriteLocaleMsg(UserIndex, 399) '¡La criatura es inmune a este hechizo!
                
                HechizoCasteado = False
                Exit Sub

            End If

        End If
        
        If Hechizos(spellindex).RemoverParalisis = 1 Then
            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                If .MaestroUser = UserIndex Then
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = 0
                     HechizoCasteado = True

                Else
                    'Provisional hasta q este habilitado eso
                    HechizoCasteado = False
                    'If Npclist(npcindex).NPCtype = eNPCType.GuardiaReal Then
                    '    If esArmada(UserIndex) Then
                    '        Call InfoHechizo(UserIndex)
                    '        Npclist(npcindex).flags.Paralizado = 0
                    '        Npclist(npcindex).Contadores.Paralisis = 0
                    '        HechizoCasteado = True
                    '        Exit Sub
                    '    Else
                    '        Call WriteConsoleMsg(1, UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    '        HechizoCasteado = False
                    '        Exit Sub
                    '    End If
                    '
                    '    Call WriteConsoleMsg(1, UserIndex, "Solo puedes Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
                    '    HechizoCasteado = False
                    '    Exit Sub
                    'Else
                    '    If Npclist(npcindex).NPCtype = eNPCType.Guardiascaos Then
                    '        If esCaos(UserIndex) Then
                    '            Call InfoHechizo(UserIndex)
                    '            Npclist(npcindex).flags.Paralizado = 0
                    '            Npclist(npcindex).Contadores.Paralisis = 0
                    '            HechizoCasteado = True
                    '            Exit Sub
                    '        Else
                    '            Call WriteConsoleMsg(1, UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    '            HechizoCasteado = False
                    '            Exit Sub
                    '        End If
                    '    End If
                    'End If
                End If
            Else
                HechizoCasteado = False
                Exit Sub
            End If

        End If
        
        If HechizoCasteado = True Then Call InfoHechizo(UserIndex)

    End With
    
    Exit Sub

HechizoEstadoNPC_Err:
182     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoNPC", Erl)
184     Resume Next
        
End Sub

Sub HechizoPropNPC(ByVal spellindex As Integer, ByVal npcindex As Integer, ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)

    On Error GoTo HechizoPropNPC_Err
    
100    Dim daño As Long

102    With Npclist(npcindex)
 
        ' VIDA
        Select Case Hechizos(spellindex).SubeHP
        
        Case 1 'Aumenta
        
        If Npclist(npcindex).Hostile > 0 Then
            HechizoCasteado = False
            Exit Sub
        End If
        
        daño = RandomNumber(Hechizos(spellindex).MinHP, Hechizos(spellindex).MaxHP)
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
        .Stats.MinHP = .Stats.MinHP + daño
          
        If .Stats.MinHP > .Stats.MaxHP Then .Stats.MinHP = .Stats.MaxHP
        Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 3)
        HechizoCasteado = True
        
        Case 2 'Reduce
        
        If Not PuedeAtacarNPC(UserIndex, npcindex) Then
            HechizoCasteado = False
            Exit Sub
        End If

        Call NPCAtacado(npcindex, UserIndex)
        daño = RandomNumber(Hechizos(spellindex).MinHP, Hechizos(spellindex).MaxHP)
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)

    
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                daño = (daño + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
            End If
        End If
        
        HechizoCasteado = True
        
        If daño < 0 Then daño = 0
        
        .Stats.MinHP = .Stats.MinHP - daño
        
        Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 2)
        
        Call CalcularDarExp(UserIndex, npcindex, daño)
    
        If .Stats.MinHP < 1 Then
            .Stats.MinHP = 0
            Call MuereNpc(npcindex, UserIndex)
        End If
        
    End Select
    
    If HechizoCasteado = True Then Call InfoHechizo(UserIndex)
    
    End With
    
    Exit Sub

HechizoPropNPC_Err:
182     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropNPC", Erl)
184     Resume Next
        
End Sub

Sub InfoHechizo(ByVal UserIndex As Integer, Optional ByVal HechizoIndex As Integer = 0, Optional ByVal daño As Long = 0, Optional color As Byte = 2)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 25/07/2009
    '25/07/2009: ZaMa - Code improvements.
    '25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
    '***************************************************
    
    On Error GoTo ErrorHandler
    
 
    Dim spellindex As Integer, tUser As Integer, tNPC As Integer
    
    With UserList(UserIndex)
    
         If HechizoIndex > 0 Then
         
1           spellindex = HechizoIndex
2
         Else
         
            spellindex = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
            
         End If
         
         tUser = .flags.TargetUser
3        tNPC = .flags.TargetNPC

4        If Hechizos(spellindex).HechizoDeArea Then Exit Sub
      
5         Call DecirPalabrasMagicas(spellindex, UserIndex)
         
6         If tUser > 0 Then

7            If Hechizos(spellindex).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(spellindex).FXgrh, Hechizos(spellindex).Loops))
8            Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(spellindex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y))
9            If Hechizos(spellindex).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageEfectoCharParticula(UserList(tUser).Char.CharIndex, Hechizos(spellindex).Particle, Hechizos(spellindex).TimeParticula, False, False))

             If tUser > 0 Then
            
                If daño > 0 Then
                   If UserIndex <> tUser Then
                       Call WriteLocaleMsg(UserIndex, 497, "^" & HechizoIndex & "%" & UserList(tUser).Name & "%" & daño, 0, color)  'Has lanzado sobre
                       Call WriteLocaleMsg(tUser, 498, UserList(UserIndex).Name & "%" & "^" & spellindex & "%" & daño, 0, color) 'Te ha lanzado
                   Else
                       Call WriteLocaleMsg(UserIndex, 499, "^" & HechizoIndex & "%" & daño, 0, color)   'Has lanzado sobre ti
                   End If
                Else
                   If UserIndex <> tUser Then
                          Call WriteLocaleMsg(UserIndex, 452, "^" & spellindex & "%" & UserList(tUser).Name, 0, color)  'Has lanzado sobre
                          Call WriteLocaleMsg(tUser, 453, UserList(UserIndex).Name & "%" & "^" & spellindex, 0, color)
                   Else
                       Call WriteLocaleMsg(UserIndex, 0, "=" & spellindex, 0, color) 'Has lanzado sobre ti
                   End If
                   
                End If
              End If
          
10         ElseIf tNPC > 0 Then
            
11            If Hechizos(spellindex).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessageCreateFX(Npclist(tNPC).Char.CharIndex, Hechizos(spellindex).FXgrh, Hechizos(spellindex).Loops))
12            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessagePlayWave(Hechizos(spellindex).WAV, Npclist(tNPC).Pos.X, Npclist(tNPC).Pos.Y))
13            If Hechizos(spellindex).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessageEfectoCharParticula(Npclist(tNPC).Char.CharIndex, Hechizos(spellindex).Particle, Hechizos(spellindex).TimeParticula, False, False))
               
14        End If

    End With

    Exit Sub

ErrorHandler:
182     Call RegistrarError(Err.Number, Err.description, "modHechizos.InfoHechizo", Erl)
184     Resume Next
        
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, _
                       ByVal UserIndex As Integer, _
                       ByVal slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim loopc As Byte

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .Stats.UserHechizos(slot) > 0 Then
                Call ChangeUserHechizo(UserIndex, slot, .Stats.UserHechizos(slot))
            Else
                Call ChangeUserHechizo(UserIndex, slot, 0)

            End If

        Else

            'Actualiza todos los slots
            For loopc = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
                If .Stats.UserHechizos(loopc) > 0 Then
                    Call ChangeUserHechizo(UserIndex, loopc, .Stats.UserHechizos(loopc))
                Else
                    Call ChangeUserHechizo(UserIndex, loopc, 0)

                End If

            Next loopc

        End If

    End With

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, _
                      ByVal slot As Byte, _
                      ByVal Hechizo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    UserList(UserIndex).Stats.UserHechizos(slot) = Hechizo
    
    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call WriteChangeSpellSlot(UserIndex, slot)
    Else
        Call WriteChangeSpellSlot(UserIndex, slot)

    End If

End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, _
                            ByVal Dire As Integer, _
                            ByVal HechizoDesplazado As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (Dire <> 1 And Dire <> -1) Then Exit Sub
    If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    With UserList(UserIndex)

        If Dire = 1 Then 'Mover arriba
            If HechizoDesplazado = 1 Then
                Call WriteLocaleMsg(UserIndex, 10)
                Exit Sub
            Else
            
                 TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
                .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado - 1)
                .Stats.UserHechizos(HechizoDesplazado - 1) = TempHechizo

            'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
            If UserList(UserIndex).flags.Hechizo > 0 Then
                UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1
            End If
            
            End If
 
        Else 'mover abajo

            If HechizoDesplazado = MAXUSERHECHIZOS Then
                Call WriteLocaleMsg(UserIndex, 10)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
                .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado + 1)
                .Stats.UserHechizos(HechizoDesplazado + 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
                If UserList(UserIndex).flags.Hechizo > 0 Then
                    UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1
                End If
                
            End If

        End If

    End With

End Sub

Public Function PuedeAyudar(ByVal UserIndex As Integer, ByVal tU As Integer) As Boolean
    
    If esRene(UserIndex) Then
        PuedeAyudar = True
        Exit Function
    End If

    If esArmada(UserIndex) Or esCiuda(UserIndex) Then 'Armada/Ciuda
        If Not (esArmada(tU) Or esCiuda(tU)) Then
            Call WriteLocaleMsg(UserIndex, 449)
            PuedeAyudar = False
            Exit Function
        End If
    End If
    
    If esRepu(UserIndex) Or esMili(UserIndex) Then
        If Not (esMili(tU) Or esRepu(tU)) Then
            Call WriteLocaleMsg(UserIndex, 449)
            PuedeAyudar = False
            Exit Function
        End If
    End If
    
    PuedeAyudar = True
End Function



 Sub HechizoCreateTelep(UserIndex As Integer, b As Boolean)
    
    On Error GoTo ErrorHandler
    Dim tU As Integer
    Dim HechizoIndex As Integer
    Dim i As Integer

    With UserList(UserIndex)
    
    HechizoIndex = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If HechizoIndex <> 53 Then Exit Sub 'Hechizo TP

      If .Pos.Map = tCiudades.Prision.Map Or .Pos.Map = Ciudades(eCiudad.cIntermundia).Map Or MapInfo(.Pos.Map).Pk = False Then
        Call WriteLocaleMsg(UserIndex, 448)
        b = False
        Exit Sub
      End If
      
      If Not LegalPos(.flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
        b = False
        Exit Sub
      End If
      
     If MapData(.Pos.Map, .flags.TargetX, .flags.TargetY).ObjInfo.ObjIndex Then
         Call WriteLocaleMsg(UserIndex, 257)
         b = False
         Exit Sub
     End If
    
     If MapData(.Pos.Map, .flags.TargetX, .flags.TargetY).TileExit.Map Then
         Call WriteLocaleMsg(UserIndex, 257)
         b = False
         Exit Sub
     End If
    
     If MapData(.Pos.Map, .flags.TargetX, .flags.TargetY).Blocked Then
        Call WriteLocaleMsg(UserIndex, 257)
         b = False
         Exit Sub
     End If
     
     If Not MapaValido(.Pos.Map) Or Not InMapBounds(.flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
     b = False
     Exit Sub
     End If
     
    
      If .Counters.TimeTeleport <> 0 Then
        Call WriteLocaleMsg(UserIndex, 457, (15 - .Counters.TimeTeleport))
        b = False
        Exit Sub
      End If
       
       
    .flags.DondeTiroX = .flags.TargetX
    .flags.DondeTiroY = .flags.TargetY
    .flags.DondeTiroMap = .flags.TargetMap
    
    If Hechizos(HechizoIndex).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(HechizoIndex).WAV, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY))
    If Hechizos(HechizoIndex).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoTerrenoParticula(Hechizos(HechizoIndex).Particle, .flags.DondeTiroX, .flags.DondeTiroY, -1))

    .flags.CasteandoPortal = True
    
    .Counters.CreoTeleport = True
    
    .Counters.TimeTeleport = 0
    
    Call DecirPalabrasMagicas(HechizoIndex, UserIndex)

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageLocaleMsg(456, .Name & "%" & "^" & HechizoIndex))
    
    b = True
    
    End With
    
    Exit Sub
    
ErrorHandler:
182     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoCreateTelep", Erl)
184     Resume Next
        
End Sub

