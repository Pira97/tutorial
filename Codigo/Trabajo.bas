Attribute VB_Name = "Trabajo"
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

Public Const CARROMINERO As Integer = 880

Private Const GASTO_ENERGIA_TRABAJADOR    As Byte = 2
Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 2

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer, ByVal DeltaTick As Single)

    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 11/19/2009
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '11/19/2009: Pato - Ahora el bandido se oculta la mitad del tiempo de las demás clases.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
    '********************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - DeltaTick

        If .Counters.TiempoOculto <= 0 Then
            .Counters.TiempoOculto = IntervaloOculto
 
            .Counters.TiempoOculto = 0
            
            .flags.Oculto = 0
 

            If .flags.Invisible = 0 Then
                Call WriteLocaleMsg(UserIndex, 307)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
                End If

        End If

    End With
    
    Exit Sub

ErrHandler:
104     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoPermanecerOculto", Erl)
106     Resume Next

End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
 
    On Error GoTo ErrHandler

    Dim Suerte As Double
    Dim res    As Integer
    Dim Skill  As Integer
    
    With UserList(UserIndex)
    
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            .Counters.TiempoOculto = IntervaloOculto
            
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
            Call SubirSkill(UserIndex, eSkill.Ocultarse)
        End If
        
        
        
        .Counters.Ocultando = .Counters.Ocultando + 1
        

    End With
    
    Exit Sub

ErrHandler:
104     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoOcultarse", Erl)
106     Resume Next

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)

    With UserList(UserIndex)
    
 
    If Not PuedeUsarSkill(UserIndex, eSkill.Navegacion, Barco) Then Exit Sub

    If .flags.Navegando = 0 Then 'Empieza a navegar
    
        If Not ((LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) And .flags.Navegando = 0) Or .flags.Navegando = 1 Then
            Call WriteLocaleMsg(UserIndex, 394)
            Exit Sub
        End If
        
        If .flags.Montando = 1 Then
            .flags.Montando = 0
            Call WriteMontateToggle(UserIndex)
        End If
        
       If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
           .flags.Oculto = 0
           .flags.Invisible = 0
           .Counters.TiempoOculto = 0
           .Counters.Invisibilidad = 0
           Call WriteLocaleMsg(UserIndex, 307) 'Has vuelto a ser visible.|12|1
           Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
       End If
 
        If .flags.Muerto = 0 Then
            .Char.body = iGraficos.iBarca
        Else
            .Char.body = iGraficos.iFragataFantasmal
        End If
                
        .Char.Head = 0
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
        
        .flags.Navegando = 1
 
        .Invent.BarcoObjIndex = .Invent.Object(slot).ObjIndex
        .Invent.BarcoSlot = slot

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.ShieldAnim, 5))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 4))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.CascoAnim, 6))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.Head, 2))
    Else
    
        If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X - 1, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X + 1, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X, .Pos.Y - 1) = True And HayAgua(.Pos.Map, .Pos.X, .Pos.Y + 1) = True Then
            Call WriteLocaleMsg(UserIndex, 430) ' ¡Debes estar cerca de una costa para bajar de tu barca.!|12|1
            Exit Sub
        End If
        
        .flags.Navegando = 0
        
        If .flags.Muerto = 0 Then
        
            .Char.Head = .OrigChar.Head
            
            If .Invent.ArmourEqpObjIndex > 0 Then
                 .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
                
            .Invent.BarcoObjIndex = 0
            .Invent.BarcoSlot = 0
        
            If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
            If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
            If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
              
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.ShieldAnim, 5))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 4))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.CascoAnim, 6))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.Head, 2))
        Else
             .Char.body = iCuerpoMuerto
             .Char.Head = iCabezaMuerto
             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.Head, 2))
        End If
        
    End If
                
 
    End With
    
    Call WriteNavigateToggle(UserIndex)

End Sub

Public Sub FundirMineral(ByVal UserIndex As Long, ByVal X As Long, ByVal Y As Long)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData( _
                    .flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.mineria) / ModFundicion(.Clase) _
                    Then
                Call DoLingotes(UserIndex)
            Else
                Call WriteLocaleMsg(UserIndex, 194) 'No tenes conocimientos de mineria suficientes para trabajar este mineral.|12|1

            End If
        
        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub

Public Sub FundirArmas(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If .flags.TargetObjInvIndex > 0 Then
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                If ObjData(.flags.TargetObjInvIndex).SkHerreria <= .Stats.UserSkills(eSkill.Herreria) / ModHerreria( _
                        .Clase) Then
                    Call DoFundir(UserIndex)
                Else
                   Call WriteLocaleMsg(UserIndex, 420) 'No tenes conocimientos de herreria suficientes para trabajar.|12|1

                End If

            End If

        End If

    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en FundirArmas. Error " & Err.Number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Integer, _
                      ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i     As Integer
    Dim Total As Long

       For i = 1 To MAX_INVENTORY_SLOTS

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount

        End If

    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, _
                         ByVal cant As Integer, _
                         ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************

    Dim i As Integer

    For i = 1 To MAX_INVENTORY_SLOTS

        With UserList(UserIndex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                
                .Amount = .Amount - cant

                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0

                End If
                
                Call UpdateUserInv(False, UserIndex, i)
                
                If cant = 0 Then Exit Sub

            End If

        End With

    Next i

End Sub

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, _
                            ByVal ItemIndex As Integer, _
                            ByVal CantidadItems As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
    '***************************************************
    With ObjData(ItemIndex)

        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, .LingH * CantidadItems, UserIndex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, .LingP * CantidadItems, UserIndex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, .LingO * CantidadItems, UserIndex)

    End With

End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, _
                               ByVal ItemIndex As Integer, _
                               ByVal CantidadItems As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora quita tambien madera elfica
    '***************************************************
    With ObjData(ItemIndex)

        If .Madera > 0 Then Call QuitarObjetos(Leña, .Madera * CantidadItems, UserIndex)

    End With

End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer)
 
If SastreTieneMateriales(UserIndex, ItemIndex, cant) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) >= _
   ObjData(ItemIndex).SkSastreria And _
   PuedeConstruirSastre(ItemIndex) And _
   UserList(UserIndex).Invent.AnilloEqpObjIndex = COSTURERO Then

    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Sastre Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 50 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 50
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 93) ' Estas muy cansado|12|1
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 50 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 50
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 93) ' Estas muy cansado|12|1
            Exit Sub
        End If
    End If
    
    Call SastreQuitarMateriales(UserIndex, ItemIndex, cant)
    Call WriteLocaleMsg(UserIndex, 431) ' ¡Has construido el objeto!|12|1
    
    Dim MiObj As Obj
    MiObj.Amount = cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    Call SubirSkill(UserIndex, eSkill.Sastreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End If
End Sub


Function ModSastreria(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Sastre
        ModSastreria = 1
    Case Else
        ModSastreria = 1
End Select

End Function

Public Function PuedeConstruirSastre(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjSastre)
    If ObjSastre(i) = ItemIndex Then
        PuedeConstruirSastre = True
        Exit Function
    End If
Next i
PuedeConstruirSastre = False

End Function

Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer) As Boolean
    If ObjData(ItemIndex).PielLobo Then
        If TieneObjetos(PielLobo, CLng(ObjData(ItemIndex).PielLobo) * CLng(cant), UserIndex) = False Then
            Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
            SastreTieneMateriales = False
            Exit Function
        End If
    End If

    If ObjData(ItemIndex).PielOsoPardo Then
        If TieneObjetos(PielOso, CLng(ObjData(ItemIndex).PielOsoPardo) * CLng(cant), UserIndex) = False Then
            Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
            SastreTieneMateriales = False
            Exit Function
        End If
    End If
    
    If ObjData(ItemIndex).PielOsoPolar Then
        If TieneObjetos(PielOsoPolar, CLng(ObjData(ItemIndex).PielOsoPolar) * CLng(cant), UserIndex) = False Then
            Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
            SastreTieneMateriales = False
            Exit Function
        End If
    End If
    
    SastreTieneMateriales = True

End Function

Sub druidaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer)
    If ObjData(ItemIndex).Raies > 0 Then
        Call QuitarObjetos(Raiz, ObjData(ItemIndex).Raies * CLng(cant), UserIndex)
    End If
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer)
    If ObjData(ItemIndex).PielLobo Then _
        Call QuitarObjetos(PielLobo, ObjData(ItemIndex).PielLobo * cant, UserIndex)
        
    If ObjData(ItemIndex).PielOsoPardo Then _
        Call QuitarObjetos(PielOso, ObjData(ItemIndex).PielOsoPardo * CLng(cant), UserIndex)
    
    If ObjData(ItemIndex).PielOsoPolar > 0 Then _
        Call QuitarObjetos(PielOsoPolar, ObjData(ItemIndex).PielOsoPolar * CLng(cant), UserIndex)
End Sub

Function druidaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer) As Boolean
    
    If ObjData(ItemIndex).Raies > 0 Then
        If Not TieneObjetos(Raiz, CLng(ObjData(ItemIndex).Raies) * CLng(cant), UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
            druidaTieneMateriales = False
            Exit Function
    End If
End If
    
    druidaTieneMateriales = True

End Function

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, _
                                   ByVal ItemIndex As Integer, _
                                   ByVal Cantidad As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Agregada validacion a madera elfica.
    '16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
    '***************************************************
    
    With ObjData(ItemIndex)

        If .Madera > 0 Then
            If Not TieneObjetos(Leña, .Madera * Cantidad, UserIndex) Then
                If ShowMsg Then Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
                
                CarpinteroTieneMateriales = False
                Exit Function

            End If

        End If
    
    End With

    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, _
                                ByVal ItemIndex As Integer, _
                                ByVal CantidadItems As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Agregada validacion a madera elfica.
    '***************************************************
    With ObjData(ItemIndex)

        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, .LingH * CantidadItems, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, .LingP * CantidadItems, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, .LingO * CantidadItems, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

    End With

    HerreroTieneMateriales = True

End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, _
                               ByVal ItemIndex As Integer, _
                               ByVal CantidadItems As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 24/08/2009
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
    '***************************************************
    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex, CantidadItems) And Round(UserList( _
            UserIndex).Stats.UserSkills(eSkill.Herreria)) >= ObjData( _
            ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(CascosHerrero)

        If CascosHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i
    
    For i = 1 To UBound(EscudosHerrero)

        If EscudosHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i
    
    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer)


If PuedeConstruir(UserIndex, ItemIndex, cant) And PuedeConstruirHerreria(ItemIndex) Then


    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Herrero Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 50 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 50
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 50 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 50
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 206) ' No tenés suficientes materiales.|12|1
            Exit Sub
        End If
    End If

    Call HerreroQuitarMateriales(UserIndex, ItemIndex, cant)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteLocaleMsg(UserIndex, 431) '¡Has construido el objeto!|12|1
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
       Call WriteLocaleMsg(UserIndex, 431) '¡Has construido el objeto!|12|1
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteLocaleMsg(UserIndex, 431) '¡Has construido el objeto!|12|1
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteLocaleMsg(UserIndex, 431) '¡Has construido el objeto!|12|1
    End If
    Dim MiObj As Obj
    MiObj.Amount = cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    Call SubirSkill(UserIndex, eSkill.Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End If
End Sub

Public Function PuedeConstruirDruida(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjDruida)
    If ObjDruida(i) = ItemIndex Then
        PuedeConstruirDruida = True
        Exit Function
    End If
Next i
PuedeConstruirDruida = False

End Function

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ObjCarpintero)

        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function

        End If

    Next i

    PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer)

 
If CarpinteroTieneMateriales(UserIndex, ItemIndex, cant) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) Then
 
    
    Dim MiObj As Obj
    MiObj.Amount = cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Exit Sub
    
    
    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Carpintero Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 20 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 20
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 206) 'No tenés suficientes materiales.|12|1
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 20 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 20
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 206)  'No tenés suficientes materiales.|12|1
            Exit Sub
        End If
    End If
 
    Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, cant)
    Call WriteLocaleMsg(UserIndex, 431) '¡Has construido el objeto!|12|1
 

    Call SubirSkill(UserIndex, eSkill.Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))


    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Else
Call WriteLocaleMsg(UserIndex, 206)  'No tenés suficientes materiales.|12|1
End If
End Sub


Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14

        Case iMinerales.PlataCruda
            MineralesParaLingote = 20

        Case iMinerales.OroCrudo
            MineralesParaLingote = 35

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '***************************************************
    '    Call LogTarea("Sub DoLingotes")
    Dim slot           As Integer
    Dim obji           As Integer
    Dim CantidadItems  As Integer
    Dim TieneMinerales As Boolean
 
    With UserList(UserIndex)
 
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))
        

        slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales

            If .Invent.Object(slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1

            End If

        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteLocaleMsg(UserIndex, 206)  'No tenés suficientes materiales.|12|1
            Exit Sub

        End If
        
        .Invent.Object(slot).Amount = .Invent.Object(slot).Amount - MineralesParaLingote(obji) * CantidadItems

        If .Invent.Object(slot).Amount < 1 Then
            .Invent.Object(slot).Amount = 0
            .Invent.Object(slot).ObjIndex = 0

        End If
        
        Dim MiObj As Obj
        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If

        Call UpdateUserInv(False, UserIndex, slot)
        Call WriteLocaleMsg(UserIndex, 207, CantidadItems) ' ¡Has obtenido #1 lingotes!|12|1
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Public Sub DoFundir(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 03/06/2010
    '03/06/2010 - Pato: Si es el último ítem a fundir y está equipado lo desequipamos.
    '11/03/2010 - ZaMa: Reemplazo división por producto para uan mejor performanse.
    '***************************************************
    Dim i           As Integer
    Dim num         As Integer
    Dim slot        As Byte
    Dim Lingotes(2) As Integer
 
    With UserList(UserIndex)
 
        slot = .flags.TargetObjInvSlot
        
        With .Invent.Object(slot)
            .Amount = .Amount - 1
            
            If .Amount < 1 Then
                If .Equipped = 1 Then Call Desequipar(UserIndex, slot)
                
                .Amount = 0
                .ObjIndex = 0

            End If

        End With
        
        num = RandomNumber(10, 25)
        
        Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * num) * 0.01
        Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * num) * 0.01
        Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * num) * 0.01
    
        Dim MiObj(2) As Obj
    
        For i = 0 To 2
            MiObj(i).Amount = Lingotes(i)
            MiObj(i).ObjIndex = LingoteHierro + i 'Una gran negrada pero práctica

            If MiObj(i).Amount > 0 Then
                If Not MeterItemEnInventario(UserIndex, MiObj(i)) Then
                    Call TirarItemAlPiso(.Pos, MiObj(i))

                End If

                Call UpdateUserInv(True, UserIndex, slot)

            End If

        Next i
        Call WriteLocaleMsg(UserIndex, 432, num) '¡Has obtenido el #1 % de los lingotes utilizados para la construcción del objeto!|12|1
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Function ModFundicion(ByVal Clase As eClass) As Single
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case Clase

        Case eClass.Guerrero
            ModFundicion = 1

        Case Else
            ModFundicion = 1

    End Select

End Function

Function Modalquimia(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Druida
        Modalquimia = 1
    Case Else
        Modalquimia = 1
End Select

End Function

Function ModCarpinteria(ByVal Clase As eClass) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case Clase

        Case eClass.Carpintero
            ModCarpinteria = 1

        Case Else
            ModCarpinteria = 1

    End Select

End Function

Function ModHerreria(ByVal Clase As eClass) As Single

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Clase

        Case eClass.Minero
            ModHerreria = 1

        Case Else
            ModHerreria = 1

    End Select

End Function

Function ModDomar(ByVal Clase As eClass) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Clase

        Case eClass.Druida
            ModDomar = 6

        Case eClass.Cazador
            ModDomar = 6

        Case eClass.Clerigo
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal npcindex As Integer)
    '***************************************************
    'Author: Nacho (Integer)
    'Last Modification: 01/05/2010
    '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
    '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
    '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
    '***************************************************

    On Error GoTo ErrHandler

    Dim puntosDomar      As Integer
    Dim puntosRequeridos As Integer
    Dim CanStay          As Boolean
    Dim petType          As Integer
    Dim NroPets          As Integer
    
    If Npclist(npcindex).MaestroUser = UserIndex Then
        Call WriteLocaleMsg(UserIndex, 421) '¡Has domado a la criatura!|12|1
        Exit Sub

    End If

    With UserList(UserIndex)

        If .NroMascotas < MAXMASCOTAS Then
            
            If Npclist(npcindex).MaestroNpc > 0 Or Npclist(npcindex).MaestroUser > 0 Then
                Call WriteLocaleMsg(UserIndex, 300) 'No puedes domar esa criatura.|12|1
                Exit Sub

            End If
            
            If Not PuedeDomarMascota(UserIndex, npcindex) Then
               Call WriteLocaleMsg(UserIndex, 254) 'No podes controlar mas criaturas.|12|1
                Exit Sub

            End If
            
            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.domar))
            
            puntosRequeridos = Npclist(npcindex).flags.Domable

            
            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
                Dim index As Integer
                .NroMascotas = .NroMascotas + 1
                index = FreeMascotaIndex(UserIndex)
                .MascotasIndex(index) = npcindex
                .MascotasType(index) = Npclist(npcindex).Numero
                
                Npclist(npcindex).MaestroUser = UserIndex
                
                Call FollowAmo(npcindex)
                Call ReSpawnNpc(Npclist(npcindex))
                
                Call WriteLocaleMsg(UserIndex, 421) '¡Has domado a la criatura!|12|1
                ' Es zona segura?
                CanStay = (MapInfo(.Pos.Map).Pk = True)
                
                If Not CanStay Then
                    petType = Npclist(npcindex).Numero
                    NroPets = .NroMascotas
                    
                    Call QuitarNPC(npcindex)
                    
                    .MascotasType(index) = petType
                    .NroMascotas = NroPets
                    
                   ' Call WriteMensajes(UserIndex, eMensajes.Mensaje467)

                End If
                
                Call SubirSkill(UserIndex, eSkill.domar)
        
            Else

                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5

                End If

            End If

        Else
            Call WriteLocaleMsg(UserIndex, 254) 'No podes controlar mas criaturas.|12|1
        End If

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, _
                                   ByVal npcindex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(i) = Npclist(npcindex).Numero Then
            numMascotas = numMascotas + 1

        End If

    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    'Makes an admin invisible o visible.
    '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '***************************************************
    
    With UserList(UserIndex)

        If .flags.AdminInvisible = 0 Then
         UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.body = 0
        UserList(UserIndex).Char.Head = 0
        UserList(UserIndex).Char.ShieldAnim = 0
        UserList(UserIndex).Char.WeaponAnim = 0
        UserList(UserIndex).Char.CascoAnim = 0
        UserList(UserIndex).showName = False
 
        Else
        .flags.AdminInvisible = 0
        .flags.Invisible = 0
        .flags.Oculto = 0
        .Counters.TiempoOculto = 0
         .Char.body = .flags.OldBody
         .Char.Head = .flags.OldHead
        UserList(UserIndex).showName = True
            End If
        
        
     Call RefreshCharStatus(UserIndex)
     Call ChangeUserCharTodo(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
 
    End With
 End Sub
 
Sub TratarDeHacerFogata(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Suerte    As Byte
    Dim exito     As Byte
    Dim Obj       As Obj
    Dim posMadera As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .X = X
        .Y = Y

    End With

    If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
        Call WriteLocaleMsg(UserIndex, 422) 'Necesitas clickear sobre leña para hacer ramitas.|12|1
        Exit Sub

    End If

    If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, 8)  ' Estas muy lejos.|12|1
        Exit Sub

    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(UserIndex, 77)  ' ¡Estás muerto! Ve al sacerdote más cercano para que puedas ser revivido.|12|1
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
        Call WriteLocaleMsg(UserIndex, 208) ' Debe haber no mas ni menos de tres troncos para hacer una fogata.|12|1
        Exit Sub

    End If

    Dim SupervivenciaSkill As Byte

    SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)

    If SupervivenciaSkill >= 0 And SupervivenciaSkill < 6 Then
        Suerte = 3
    ElseIf SupervivenciaSkill >= 6 And SupervivenciaSkill <= 34 Then
        Suerte = 2
    ElseIf SupervivenciaSkill >= 35 Then
        Suerte = 1

    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        Obj.ObjIndex = FOGATA_APAG
        Obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
        Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
        Call MakeObj(Obj, Map, X, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    
        Call SubirSkill(UserIndex, eSkill.Supervivencia)
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
            Call WriteLocaleMsg(UserIndex, 171) 'No has podido hacer fuego.|12|1
            UserList(UserIndex).flags.UltimoMensaje = 10

        End If
    End If

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer
    Dim res           As Integer
    Dim CantidadItems As Integer
 
    If UserList(UserIndex).Clase = eClass.PescadoR Then
        Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    Else
        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)

    End If

    Dim Skill As Integer
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.pesca)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

    res = RandomNumber(1, Suerte)

    If res <= 6 Then
        Dim MiObj As Obj
    
        If UserList(UserIndex).Clase = eClass.PescadoR Then

            With UserList(UserIndex)
                CantidadItems = 1 + MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

            End With
        
            MiObj.Amount = RandomNumber(10, 30)
        Else
           MiObj.Amount = RandomNumber(10, 30)

        End If

        MiObj.ObjIndex = Pescado
    
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

        End If
    
        Call WriteLocaleMsg(UserIndex, 64) '¡Has pescado algo!|2|0
        Call SubirSkill(UserIndex, eSkill.pesca)
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
           Call WriteLocaleMsg(UserIndex, 59) '¡No has pescado nada!|2|0
            UserList(UserIndex).flags.UltimoMensaje = 6

        End If

    End If


    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    Exit Sub

ErrHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.description)

End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo ErrHandler

    Dim iSkill     As Integer
    Dim Suerte     As Integer
    Dim res        As Integer
    Dim EsPescador As Boolean
 
 '   If .Stats.UserSkills(eSkill.Pesca) < ObjData(318).MinSkill Then
 '      Call WriteConsoleMsg(userindex, "Para usar esta montura necesitas " & ObjData(318).MinSkill & " puntos en pesca.", FontTypeNames.FONTTYPE_INFO)
 '      Exit Sub
  'End If
  
    If UserList(UserIndex).Clase = eClass.PescadoR Then
        Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
        EsPescador = True
    Else
        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
        EsPescador = False

    End If

    iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.pesca)

    ' m = (60-11)/(1-10)
    ' y = mx - m*10 + 11

    Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

    If Suerte > 0 Then
        res = RandomNumber(1, Suerte)
    
        If res < 6 Then
            Dim MiObj                 As Obj
            Dim PecesPosibles(1 To 4) As Integer
        
            PecesPosibles(1) = PESCADO1
            PecesPosibles(2) = PESCADO2
            PecesPosibles(3) = PESCADO3
            PecesPosibles(4) = PESCADO4
        
            If EsPescador = True Then
                MiObj.Amount = RandomNumber(10, 30)
            Else
                MiObj.Amount = 10

            End If

            MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
        
            Call WriteLocaleMsg(UserIndex, 64) ' ¡Has pescado algo!|2|0
        
            Call SubirSkill(UserIndex, eSkill.pesca)
        Else
            Call WriteLocaleMsg(UserIndex, 59) 'NO has pescado nada

        End If

    End If
        
    Exit Sub

ErrHandler:
    Call LogError("Error en DoPescarRed")

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************

    On Error GoTo ErrHandler
    If esRene(LadrOnIndex) Then
    
    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
    
    With UserList(LadrOnIndex)
    
        ' Caos robando a caos?
        If esCaos(LadrOnIndex) And esCaos(VictimaIndex) Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de las fuerzas del caos.", _
                    FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            Exit Sub

        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
         
          
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte     As Integer
            Dim res        As Integer
            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.robar)
                
            If RobarSkill <= 10 And RobarSkill >= -1 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 And RobarSkill >= 11 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 And RobarSkill >= 21 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 And RobarSkill >= 31 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 And RobarSkill >= 41 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 And RobarSkill >= 51 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 And RobarSkill >= 61 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 And RobarSkill >= 71 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 And RobarSkill >= 81 Then
                Suerte = 10
            ElseIf RobarSkill < 100 And RobarSkill >= 91 Then
                Suerte = 7
            ElseIf RobarSkill = 100 Then
                Suerte = 5

            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
               
                If (RandomNumber(1, 50) < 25) And (.Clase = eClass.ladron) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", _
                                FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(VictimaIndex).Stats.GLD > 0 Then
                        Dim n As Integer
                        
                        If .Clase = eClass.ladron Then
 
                
                                n = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            
                        Else
                            n = RandomNumber(1, 100)

                        End If

                        If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                        n = n * 5
  
                        If OroLleno(LadrOnIndex, UserList(LadrOnIndex).Stats.GLD, CLng(n)) Then
                        Call WriteConsoleMsg(LadrOnIndex, "Tienes la cantidad máxima de oro que puedes tener. No has obtenido oro", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                        End If

                        .Stats.GLD = .Stats.GLD + n

                        If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList( _
                                VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Call FlushBuffer(VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", _
                                FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.robar)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
               ' Call WriteConsoleMsg( VictimaIndex, "¡" & .name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
                Call FlushBuffer(VictimaIndex)
                

            End If

        End If

    End With

    Else
    
    Call WriteConsoleMsg(LadrOnIndex, "Para robar debes ser renegado, escribe /RETIRAR para dejar tu ciudadania.", FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub

ErrHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal slot As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' Agregué los barcos
    ' Esta funcion determina qué objetos son robables.
    '***************************************************

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(slot).Equipped = _
            0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).Milicia = 0 And ObjData(OI).Shop = 0 And ObjData(OI).EfectoMagico = eMagicType.Sacrificio And ObjData(OI).OBJType <> eOBJType.otMonturas And ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************

    Dim flag As Boolean
    Dim i    As Integer
    flag = False

    If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
        i = 1

        Do While Not flag And i <= MAX_INVENTORY_SLOTS


            'Hay objeto en este slot?
            If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
                If ObjEsRobable(VictimaIndex, i) Then
                    If RandomNumber(1, 10) < 4 Then flag = True

                End If

            End If

            If Not flag Then i = i + 1
        Loop
    Else
        i = 20

        Do While Not flag And i > 0

            'Hay objeto en este slot?
            If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
                If ObjEsRobable(VictimaIndex, i) Then
                    If RandomNumber(1, 10) < 4 Then flag = True

                End If

            End If

            If Not flag Then i = i - 1
        Loop

    End If

    If flag Then
        Dim MiObj     As Obj
        Dim num       As Byte
        Dim ObjAmount As Integer
    
        ObjAmount = UserList(VictimaIndex).Invent.Object(i).Amount
    
        'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
        num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                
        MiObj.Amount = num
        MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
         If ObjData(UserList(VictimaIndex).Invent.Object(i).ObjIndex).OBJType = eOBJType.otbebidas Or eOBJType.otUseOnce Then Exit Sub
 
        UserList(VictimaIndex).Invent.Object(i).Amount = ObjAmount - num
                
        If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
            Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

        End If
            
        Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
        If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

        End If
    
        If UserList(LadrOnIndex).Clase = eClass.ladron Then
            Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, _
                    FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, _
                    FontTypeNames.FONTTYPE_INFO)

        End If

    End If

    'If exiting, cancel de quien es robado
    Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, _
        ByVal daño As Integer)
    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el daño
    '***************************************************
    Dim Suerte As Integer
    Dim Skill  As Integer

    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

    Select Case UserList(UserIndex).Clase

        Case eClass.Asesino
            Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
        Case eClass.Clerigo, eClass.Paladin, eClass.Sastre
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bardo
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            If UserList(UserIndex).Clase = eClass.Asesino Then
                daño = Round(daño * 1.4, 0)
            Else
                daño = Round(daño * 1.5, 0)

            End If
        
            UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
            Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, _
                    FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño, _
                    FontTypeNames.FONTTYPE_FIGHT)
        
            Call FlushBuffer(VictimUserIndex)
        Else
            Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
            Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 2), _
                    FontTypeNames.FONTTYPE_FIGHT)
            '[Alejo]
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño * 2)

        End If
    
        Call SubirSkill(UserIndex, eSkill.Apuñalar)
    Else
        Call WriteLocaleMsg(UserIndex, 79) '¡No has logrado apuñalar a tu enemigo!|2|0

    End If

End Sub


Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer
    Dim res           As Integer
    Dim CantidadItems As Integer

    If UserList(UserIndex).Clase = eClass.Leñador Then
        Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
    Else
        Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)

    End If

    Dim Skill As Integer
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.talar)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

    res = RandomNumber(1, Suerte)

    If res <= 6 Then
        Dim MiObj As Obj
    
        If UserList(UserIndex).Clase = eClass.Leñador Then

            With UserList(UserIndex)
                CantidadItems = 1 + MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

            End With
        
        MiObj.Amount = RandomNumber(10, 30)
        Else
          MiObj.Amount = RandomNumber(10, 30)

        End If
    
        MiObj.ObjIndex = Leña
    
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
        End If
    
        Call WriteLocaleMsg(UserIndex, 65) '¡Has conseguido algo de leña!|2|0
    
        Call SubirSkill(UserIndex, eSkill.talar)
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
            Call WriteLocaleMsg(UserIndex, 61) '¡No has obtenido leña!|2|0
            UserList(UserIndex).flags.UltimoMensaje = 8

        End If

    End If

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    Exit Sub

ErrHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoBotanica(ByVal UserIndex As Integer)
On Error GoTo ErrHandler

Dim Suerte As Integer
Dim res As Integer

    If UserList(UserIndex).Clase = eClass.Druida Then
        Call QuitarSta(UserIndex, EsfuerzoBotanicaDruida)
    Else
        Call QuitarSta(UserIndex, EsfuerzoBotanicaGeneral)
    End If
    
    Dim Skill As Integer
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.botanica)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        
        If UserList(UserIndex).Clase = eClass.Druida Then
            MiObj.Amount = RandomNumber(10, 30)
        Else
            MiObj.Amount = RandomNumber(10, 30)
        End If
        
        MiObj.ObjIndex = Raiz
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido raíces!(" & MiObj.Amount & ")", FontTypeNames.FONTTYPE_INFO)
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
    
    Else
        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
            Call WriteConsoleMsg(UserIndex, "¡No has obtenido raices!", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 8
        End If
        '[/CDT]
    End If
    
     Call SubirSkill(UserIndex, eSkill.botanica)
        
    
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

ErrHandler:
    Call LogError("Error en DoBotanica")

End Sub


Public Sub DoMineria(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer
    Dim res           As Integer
    Dim CantidadItems As Integer

    With UserList(UserIndex)

        If .Clase = eClass.Minero Then
            Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
        Else
            Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)

        End If
    
        Dim Skill As Integer
        Skill = .Stats.UserSkills(eSkill.mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= 5 Then
            Dim MiObj As Obj
        
            If .flags.TargetObj = 0 Then Exit Sub
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
        
            If UserList(UserIndex).Clase = eClass.Minero Then
                CantidadItems = 1 + MaximoInt(1, CInt((.Stats.ELV - 4) / 5))
            
                MiObj.Amount = RandomNumber(10, 30)
            Else
               MiObj.Amount = RandomNumber(10, 30)

            End If
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteLocaleMsg(UserIndex, 66) '¡Has extraido algunos minerales!|2|0
        
            Call SubirSkill(UserIndex, eSkill.mineria)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteLocaleMsg(UserIndex, 62) '¡No has conseguido nada!|2|0
                .flags.UltimoMensaje = 9

            End If


        End If
        
        .Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoMineria")

End Sub
Public Sub DoMeditar(ByVal UserIndex As Integer, ByVal DeltaTick As Single)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(UserIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If

If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then Exit Sub

If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 7
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Round(Suerte / DeltaTick))

If res = 1 Then
    
     
    cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, PorcentajeRecuperoMana)
    cant = cant / 10
     
    If cant <= 0 Then cant = 1
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
        Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 22
    End If
    
    Call WriteUpdateMana(UserIndex)
     Call SubirSkill(UserIndex, eSkill.Meditar)
End If

End Sub
Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modif: 15/04/2010
    'Unequips either shield, weapon or helmet from target user.
    '***************************************************

    Dim Probabilidad   As Integer
    Dim Resultado      As Integer
    Dim WrestlingSkill As Byte
    Dim AlgoEquipado   As Boolean
    
    With UserList(UserIndex)

        ' Si es bardo o gladi desequipa
    If Not UserList(UserIndex).Clase = eClass.Gladiador Or Not UserList(UserIndex).Clase = eClass.Bardo Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66

    End With
   
    With UserList(VictimIndex)

        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", _
                        FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", _
                            FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", _
                        FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", _
                            FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", _
                    FontTypeNames.FONTTYPE_FIGHT)

        End If
    
    End With

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 03/03/2010
    'Implements the pick pocket skill of the Bandit :)
    '03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
    '***************************************************
    If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(UserIndex).Clase <> eClass.ladron Then Exit Sub
 
    Dim res As Integer
    res = RandomNumber(1, 100)

    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
            Call RobarObjeto(UserIndex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", _
                    FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", _
                    FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************

    Dim Probabilidad   As Integer
    Dim Resultado      As Integer
    Dim WrestlingSkill As Byte
    
    With UserList(UserIndex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
            If UserList(VictimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

            Call FlushBuffer(VictimIndex)

        End If

    End With
    
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '
    '***************************************************
    MaxItemsConstruibles = MaximoInt(1, CInt((UserList(UserIndex).Stats.ELV - 4) / 5))

End Function
Public Function PuedeUsarSkill(ByVal UserIndex As Integer, ByVal Skill As Integer, ByRef Obj As ObjData) As Boolean
        
    On Error GoTo ErrorHandler
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then

         If UserList(UserIndex).Stats.UserSkills(Skill) < Obj.MinSkill Then
            Call WriteLocaleMsg(UserIndex, 477, Obj.MinSkill & "%" & "*" & Skill)
            PuedeUsarSkill = False
            Exit Function
        Else
            PuedeUsarSkill = True
        End If
    
    Else
        PuedeUsarSkill = True
        
    End If
    
    Exit Function

ErrorHandler:
116     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeUsarSkill", Erl)
118     Resume Next
        
End Function

Public Function DoEquita(ByVal UserIndex As Integer, ByRef Montura As ObjData, ByVal slot As Integer)

    With UserList(UserIndex)
    
        If Not PuedeUsarSkill(UserIndex, eSkill.Equitacion, Montura) Then Exit Function
        
        If .flags.Montando = 0 Then
            If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger >= 20) Or (MapInfo(.Pos.Map).Zona = "DUNGEON") Then
                Call WriteLocaleMsg(UserIndex, 253) 'No puedes montara aqui
                Exit Function
            End If
            
        End If
          
        If .flags.Navegando = 1 Then
            Call WriteLocaleMsg(UserIndex, 20)
            Exit Function
        End If
            
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            Call WriteLocaleMsg(UserIndex, 307) 'Has vuelto a ser visible.|12|1
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
                      
        If .flags.Montando = 0 Then
    
            .Char.body = Montura.Ropaje
            .Char.Head = .OrigChar.Head
            .Char.WeaponAnim = NingunArma
                    
            .flags.Montando = 1
            
            .Invent.MonturaObjIndex = .Invent.Object(slot).ObjIndex
            .Invent.MonturaSlot = slot
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 4))
            
       Else
       
            .Char.Head = .OrigChar.Head
            
            .flags.Montando = 0
            
       
            If .Invent.ArmourEqpObjIndex > 0 Then
                .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
            
            
            .Invent.MonturaObjIndex = 0
            .Invent.MonturaSlot = 0
            
            If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
            If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
            If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
              
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.ShieldAnim, 5))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 4))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.CascoAnim, 6))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
            
        End If
    
        Call WriteMontateToggle(UserIndex)
 
    End With
End Function
Function ModEquitacion(ByVal Clase As String) As Integer
Select Case UCase$(Clase)
    Case "1"
        ModEquitacion = 1
    Case Else
        ModEquitacion = 1
End Select
 
End Function

Public Sub druidaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cant As Integer)

If druidaTieneMateriales(UserIndex, ItemIndex, cant) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.alquimia) >= _
   ObjData(ItemIndex).SkPociones And _
   PuedeConstruirDruida(ItemIndex) And _
   UserList(UserIndex).Invent.AnilloEqpObjIndex = OLLA Then
   
    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Druida Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 30 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 30
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 206) 'No tenés suficientes materiales.|12|1
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSta >= 30 Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 30
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 206) 'No tenés suficientes materiales.|12|1
            Exit Sub
        End If
    End If
    
    Call druidaQuitarMateriales(UserIndex, ItemIndex, cant)
    Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = cant
    MiObj.ObjIndex = ItemIndex
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    Call SubirSkill(UserIndex, eSkill.alquimia)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))


    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End If
End Sub

Public Function DoTrabajar(ByVal UserIndex As Integer)
 
    
    If Not IntervaloPermiteTrabajar(UserIndex, True) Then Exit Function
    
    With UserList(UserIndex)
        If .Stats.MinSta < 2 Then
             Call WriteConsoleMsg(UserIndex, "Dejas de trabajar debido a tu poca energía.", FontTypeNames.FONTTYPE_INFO)
            .flags.Trabajando = False
            Exit Function
        End If
        
        If .flags.Lingoteando Then
            Call DoLingotes(UserIndex)
        ElseIf .Invent.AnilloEqpSlot <> 0 Then
            Select Case .Invent.AnilloEqpObjIndex
                Case RED_PESCA
                    Call DoPescar(UserIndex)
                    
                Case CAÑA_PESCA
                    Call DoPescar(UserIndex)
                    
                Case PIQUETE_MINERO
                    Call DoMineria(UserIndex)
                  
                Case HACHA_LEÑADOR
                    Call DoTalar(UserIndex)
                    
                Case TIJERAS
                    Call DoBotanica(UserIndex)
            End Select
        End If
    End With
End Function
Public Function PuedeCasarse(ByVal UserIndex As Integer, ByVal Pareja As Integer)
 
    'Mermas sub
    
    On Error GoTo ErrorHandler
 
1    If DeadCheck(UserIndex) Then Exit Function
    
2    If UserList(Pareja).flags.Muerto = 1 Then
3        Call WriteLocaleMsg(UserIndex, 7) 'El usuario está muerto
4        Exit Function
5    End If
    
6    If UserIndex = Pareja Then
7        Call WriteLocaleMsg(UserIndex, 492) '¡No puedes casarte contigo mismo!
8        Exit Function
9    End If
    
10    If UserList(Pareja).Genero = UserList(UserIndex).Genero Then
11        Call WriteLocaleMsg(UserIndex, 117) 'No podes casarte con un usuario de tu mismo género.
12        Exit Function
13    End If
    
14    If Distancia(UserList(Pareja).Pos, UserList(UserIndex).Pos) > 1 Then
16        Call WriteLocaleMsg(UserIndex, 495) ' El objetivo está muy lejos.
15        Exit Function
17    End If

19    If UserList(UserIndex).Casamiento.Pareja = UserList(UserIndex).Name Then
18        Call WriteLocaleMsg(UserIndex, 116) 'Ya estas casado
20        Exit Function
21    End If
    
22    If UserList(Pareja).Casamiento.Casado = 1 Then
23        Call WriteLocaleMsg(UserIndex, 120) '¡El usuario ya esta casado!
24        Exit Function
25    End If
    
27    If UserList(Pareja).Casamiento.Candidato = UserIndex Then
    
28        UserList(Pareja).Casamiento.Casado = 1
29        UserList(Pareja).Casamiento.Pareja = UserList(UserIndex).Name
                   
30        UserList(UserIndex).Casamiento.Casado = 1
31        UserList(UserIndex).Casamiento.Pareja = UserList(Pareja).Name
                
32        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(236, Trim(MapInfo(UserList(UserIndex).Pos.Map).Name) & "%" & UserList(UserIndex).Name & "%" & UserList(Pareja).Name))
34        Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessagePlayWave(snd_casamiento, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
33        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SANAR, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
 
35    Else
36        'Se asegura que el target es un npc
37        If UserList(UserIndex).flags.TargetNPC = 0 Then
38            Call WriteLocaleMsg(UserIndex, 493) 'Primero tienes que seleccionar al sacerdote, haz click izquierdo sobre él
39            Exit Function
40        End If
        
        'Validate NPC and make sure player is dead
41        If (Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor) Then Exit Function
        
        'Make sure it's close enough
42        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 5 Then
43            Call WriteLocaleMsg(UserIndex, 494) 'Estas muy lejos del sacerdote.
45            Exit Function
44        End If
 
        
46        Call WriteLocaleMsg(UserIndex, 495, UserList(Pareja).Name) 'Le has mandado una propuesta de casamiento a #1, espera su respuesta...|
        
47        Call WriteEjecutarAccion(Pareja, 3, UserList(UserIndex).Name)

          UserList(UserIndex).Casamiento.Candidato = Pareja
          
50    End If

    Exit Function

ErrorHandler:
116     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeCasarse", Erl)
118     Resume Next
        
End Function
