Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '17/09/02
    'Agregue que la función se asegure que el objeto no es un barco

    On Error Resume Next

    Dim i        As Integer
    Dim ObjIndex As Integer

    For i = 1 To MAX_INVENTORY_SLOTS

        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).Newbie = 0) Then
                TieneObjetosRobables = True
                Exit Function

            End If
    
        End If

    Next i

End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            Optional ByRef sMotivo As Integer) As Boolean
                            
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo manejador
    If ObjIndex = 0 Then Exit Function

    Dim flag As Boolean
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
            Dim i As Integer

            For i = 1 To NUMCLASES
                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                sMotivo = 265
                ClasePuedeUsarItem = False
                Exit Function
                End If
            Next i
        End If
    End If
    
    ClasePuedeUsarItem = True
    
    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")

End Function
 Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

End Sub
 
 
Sub LimpiarInventario(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer

    With UserList(UserIndex)

      For j = 1 To MAX_INVENTORY_SLOTS
        .Invent.Object(j).ObjIndex = 0
        .Invent.Object(j).Amount = 0
        .Invent.Object(j).Equipped = 0
    Next j
    
        .Invent.NroItems = 0
    
        .Invent.ArmourEqpObjIndex = 0
        .Invent.ArmourEqpSlot = 0
    
        .Invent.NudiEqpObjIndex = 0
        .Invent.NudiEqpSlot = 0
        
        .Invent.WeaponEqpObjIndex = 0
        .Invent.WeaponEqpSlot = 0
    
        .Invent.CascoEqpObjIndex = 0
        .Invent.CascoEqpSlot = 0
    
        .Invent.EscudoEqpObjIndex = 0
        .Invent.EscudoEqpSlot = 0
    
        .Invent.AnilloEqpObjIndex = 0
        .Invent.AnilloEqpSlot = 0
        
        
        .Invent.MunicionEqpObjIndex = 0
        .Invent.MunicionEqpSlot = 0
    
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0

        .Invent.MonturaObjIndex = 0
        .Invent.MonturaSlot = 0
        
        .Invent.MagicIndex = 0
        .Invent.MagicSlot = 0
    End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
    '***************************************************
    On Error GoTo ErrHandler

    'If Cantidad > 100000 Then Exit Sub

    With UserList(UserIndex)

        'SI EL Pjta TIENE ORO LO TIRAMOS
        If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then
            Dim i     As Byte
            Dim MiObj As Obj
            'info debug
            Dim Loops As Integer
            
            'Seguridad Alkon (guardo el oro tirado si supera los 50k)
            If Cantidad > 50000 Then
                Dim j        As Integer
                Dim K        As Integer
                Dim M        As Integer
                Dim Cercanos As String
                M = .Pos.Map

                For j = .Pos.X - 10 To .Pos.X + 10
                    For K = .Pos.Y - 10 To .Pos.Y + 10

                        If InMapBounds(M, j, K) Then
                            If MapData(M, j, K).UserIndex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(M, j, K).UserIndex).Name & ","
                                  
                            End If

                        End If

                    Next K
                Next j
           

                Call LogDesarrollo(.Name & " tira oro. Cercanos: " & Cercanos)
               
            End If

            '/Seguridad
            Dim Extra    As Long
            Dim TeniaOro As Long
            TeniaOro = .Stats.GLD

            If Cantidad > 500000 Then 'Para evitar explotar demasiado
                Extra = Cantidad - 500000
                Cantidad = 500000

            End If
            
            Do While (Cantidad > 0)
                
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount

                End If
    
                MiObj.ObjIndex = iORO
                If EsGm(UserIndex) Then Call LogGM(.Name, "Tiró cantidad:" & MiObj.Amount & " Objeto:" & ObjData( _
                        MiObj.ObjIndex).Name)
                Dim AuxPos As WorldPos
                
                If .Clase = eClass.Sastre And .Invent.BarcoObjIndex = 476 Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)

                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount

                    End If

                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)

                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount

                    End If

                End If
                
                'info debug
                Loops = Loops + 1

                If Loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub

                End If
                
            Loop

            If TeniaOro = .Stats.GLD Then Extra = 0
            If Extra > 0 Then
                .Stats.GLD = .Stats.GLD - Extra

            End If
        
        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)

    On Error GoTo ErrHandler

    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub

    
    With UserList(UserIndex).Invent.Object(slot)

        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, slot)
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad

        '¿Quedan mas?
        If .Amount <= 0 Then
             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                  ByVal UserIndex As Integer, _
                  ByVal slot As Byte, Optional ByVal SubioNivel As Boolean = False)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim NullObj As UserObj
    Dim loopc   As Long

    Dim Obj      As ObjData
    Dim ObjIndex As Integer
    
    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then
    
            'Actualiza el inventario
            If .Invent.Object(slot).ObjIndex > 0 Then
                Call ChangeUserInv(UserIndex, slot, .Invent.Object(slot))
            Else
                Call ChangeUserInv(UserIndex, slot, NullObj)

            End If
    
        Else
    
            'Actualiza todos los slots
            For loopc = 1 To MAX_INVENTORY_SLOTS
        
             If SubioNivel = True Then
                    
                ObjIndex = .Invent.Object(loopc).ObjIndex
                Obj = ObjData(ObjIndex)
    
                If .Stats.ELV >= Obj.levelItem Then Call ChangeUserInv(UserIndex, loopc, .Invent.Object(loopc))
                
             Else
    
                'Actualiza el inventario
                If .Invent.Object(loopc).ObjIndex > 0 Then
                    Call ChangeUserInv(UserIndex, loopc, .Invent.Object(loopc))
                Else
                    Call ChangeUserInv(UserIndex, loopc, NullObj)

                End If

             End If
            Next loopc

        End If
    
        Exit Sub

    End With

ErrHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)

End Sub

Sub UpdateUserInventario(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Accion As Byte, ByVal Valor As Integer)

    On Error GoTo ErrHandler

    Dim NullObj As UserObj

    With UserList(UserIndex)

    'Actualiza el inventario
    If .Invent.Object(slot).ObjIndex > 0 Then

        .Invent.Object(slot) = .Invent.Object(slot)
        Call writeChangeInventorySlotUser(UserIndex, slot, Accion, Valor)
        
    Else
        .Invent.Object(slot) = NullObj
        Call writeChangeInventorySlotUser(UserIndex, slot, Accion, Valor)

    End If
        
    Exit Sub

    End With

ErrHandler:
    Call LogError("Error en UpdateInv. Error " & Err.Number & " : " & Err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Obj As Obj
    Dim ObjDestroy As ObjData
    Dim ObjIndex As Integer
    
    With UserList(UserIndex)
   
      ObjIndex = .Invent.Object(slot).ObjIndex
      ObjDestroy = ObjData(ObjIndex)
            
        If num > 0 Then
    
            If num > .Invent.Object(slot).Amount Then num = .Invent.Object(slot).Amount
        
            Obj.ObjIndex = .Invent.Object(slot).ObjIndex
            
           'Shermie80
           If .flags.Privilegios And PlayerType.User Then

               If ObjDestroy.Real Or ObjDestroy.Caos Or ObjDestroy.Milicia Then
                   Call WriteLocaleMsg(UserIndex, 260)
                  Exit Sub

               End If
               
            End If
            'Fin
            
            If (ItemNewbie(Obj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
                Call WriteLocaleMsg(UserIndex, 260)
                Exit Sub

            End If
            
                        
            If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 And MapData(Map, X, Y).ObjInfo.Amount > 9999 Then
                Call WriteLocaleMsg(UserIndex, 262)
                Exit Sub
            End If


            'Check objeto en el suelo
            If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Or MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
                If num + MapData(.Pos.Map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
                    num = MAX_INVENTORY_OBJS - MapData(.Pos.Map, X, Y).ObjInfo.Amount

                End If
             
                Obj.Amount = num
            
                Call MakeObj(Obj, Map, X, Y)
                Call QuitarUserInvItem(UserIndex, slot, num)
                Call UpdateUserInv(False, UserIndex, slot)
            Else
                Call WriteLocaleMsg(UserIndex, 262)

            End If

        End If

    End With

End Sub

Sub EraseObj(ByVal num As Integer, _
             ByVal Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With MapData(Map, X, Y)
        .ObjInfo.Amount = .ObjInfo.Amount - num
    
        If .ObjInfo.Amount <= 0 Then

            .ObjInfo.ObjIndex = 0
            .ObjInfo.Amount = 0
        
            Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

    End With

End Sub

Sub MakeObj(ByRef Obj As Obj, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then
    
        With MapData(Map, X, Y)

            If .ObjInfo.ObjIndex = Obj.ObjIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(X, Y, Obj.ObjIndex, .ObjInfo.Amount))

            Else
 
                .ObjInfo = Obj
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(X, Y, Obj.ObjIndex, Obj.Amount))

            End If

        End With

    End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim X    As Integer
    Dim Y    As Integer
    Dim slot As Byte
    
    With UserList(UserIndex)
        '¿el user ya tiene un objeto del mismo tipo?
        slot = 1

        Do Until .Invent.Object(slot).ObjIndex = MiObj.ObjIndex And .Invent.Object(slot).Amount + MiObj.Amount <= _
                MAX_INVENTORY_OBJS
            slot = slot + 1

            If slot > MAX_INVENTORY_SLOTS Then
                 Exit Do
           End If
        Loop
           
        'Sino busca un slot vacio
        If slot > MAX_INVENTORY_SLOTS Then
           slot = 1
           Do Until .Invent.Object(slot).ObjIndex = 0
               slot = slot + 1
               If slot > MAX_INVENTORY_SLOTS Then
                    Call WriteLocaleMsg(UserIndex, 25)
                    MeterItemEnInventario = False
                    Exit Function

                End If

            Loop
            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        

        'Mete el objeto
        If .Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(slot).ObjIndex = MiObj.ObjIndex
            .Invent.Object(slot).Amount = .Invent.Object(slot).Amount + MiObj.Amount
        Else
            .Invent.Object(slot).Amount = MAX_INVENTORY_OBJS

        End If

    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, UserIndex, slot)
    
    Exit Function
ErrHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.description)

End Function

Sub GetObj(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 18/12/2009
    '18/12/2009: ZaMa - Oro directo a la billetera.
    '***************************************************

    Dim Obj    As ObjData
    Dim MiObj  As Obj
    Dim ObjPos As String
    
    With UserList(UserIndex)

        '¿Hay algun obj?
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
                Dim X    As Integer
                Dim Y    As Integer
                Dim slot As Byte
                
                X = .Pos.X
                Y = .Pos.Y
                
                Obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
                MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                ' Oro directo a la billetera!
                If Obj.OBJType = otGuita Then
                
                
                    
                        If Not OroLleno(UserIndex, UserList(UserIndex).Stats.GLD, CLng(MiObj.Amount)) Then
                            .Stats.GLD = .Stats.GLD + MiObj.Amount
                            'Quitamos el objeto
                            Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                        
                            Call WriteUpdateGold(UserIndex)
                            
                        Else
                            Call WriteConsoleMsg(UserIndex, "No puedes juntar este oro por que tendrias mas del maximo disponible (2147483647)", FontTypeNames.FONTTYPE_INFO)
                        End If
                    
                Else

                    If MeterItemEnInventario(UserIndex, MiObj) Then
                    
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)



                        End If

                End If

            End If

        Else
           Call WriteLocaleMsg(UserIndex, 261)

        End If

    End With

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal slot As Byte)

    On Error GoTo ErrHandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(UserIndex)
    
        With .Invent

            If (slot < LBound(.Object)) Or (slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(slot).ObjIndex = 0 Then
                Exit Sub
            End If
            
            Obj = ObjData(.Object(slot).ObjIndex)

        End With
        
        Select Case Obj.OBJType
        
        Case eOBJType.otMonturas
        
            Call DoEquita(UserIndex, Obj, slot)
            
        Case eOBJType.otWeapon
            
            .Invent.Object(slot).Equipped = 0
            .Invent.WeaponEqpObjIndex = 0
            .Invent.WeaponEqpSlot = 0
            
            .Char.WeaponAnim = NingunArma
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 4))
            
            If .Char.Arma_Aura <> 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 1))
                .Char.Arma_Aura = 0
            End If
                
                    
        Case eOBJType.otNudillos
            
            .Invent.Object(slot).Equipped = 0
            .Invent.NudiEqpObjIndex = 0
            .Invent.NudiEqpSlot = 0
            
            .Char.WeaponAnim = NingunArma
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 8))
            
            If .Char.Arma_Aura <> 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 1))
                .Char.Arma_Aura = 0
            End If
            
        Case eOBJType.otItemsMagicos
    
            .Invent.MagicIndex = 0
            .Invent.MagicSlot = 0
            .Invent.Object(slot).Equipped = 0
            
            If .Char.Anillo_Aura <> 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 6))
                .Char.Anillo_Aura = 0
            End If
            
            Select Case Obj.EfectoMagico
                
                Case eMagicType.ModificaAtributo
                    If Obj.QueAtributo <> 0 Then
                        .Stats.UserAtributos(Obj.QueAtributo) = .Stats.UserAtributos(Obj.QueAtributo) - Obj.CuantoAumento
                    End If
                    
                Case eMagicType.ModificaSkill
                
                    If Obj.QueSkill <> 0 Then
                        .Stats.UserSkills(Obj.QueSkill) = .Stats.UserSkills(Obj.QueSkill) - Obj.CuantoAumento
                    End If
                    
            End Select
            
        
        Case eOBJType.otFlechas
            
            .Invent.Object(slot).Equipped = 0
            .Invent.MunicionEqpObjIndex = 0
            .Invent.MunicionEqpSlot = 0
            
            
        Case eOBJType.otAnillo

            
            .Invent.Object(slot).Equipped = 0
            .Invent.AnilloEqpObjIndex = 0
            .Invent.AnilloEqpSlot = 0

            If .flags.Trabajando = True Then
                .flags.Trabajando = False
                Call WriteLocaleMsg(UserIndex, 391, vbNullString, 1)
            End If
 
        Case eOBJType.otArmadura  ' Puede ser un escudo, casco , o vestimenta
                        
             .Invent.Object(slot).Equipped = 0
             
                Select Case Obj.SubTipo
                
                Case 0 'Armadura
                
                    .Invent.ArmourEqpObjIndex = 0
                    .Invent.ArmourEqpSlot = 0
                    
                    If .Char.Body_Aura <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 2))
                        .Char.Body_Aura = 0
                    End If
                    
                    Call DarCuerpoDesnudo(UserIndex)
                    
                    If .flags.Montando = 0 And .flags.Navegando = 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
                    End If
                    

                Case 1 'Casco
                
                    .Invent.CascoEqpObjIndex = 0
                    .Invent.CascoEqpSlot = 0
                
                    If .Char.Head_Aura <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 4))
                        .Char.Head_Aura = 0
                    End If
            
                    .Char.CascoAnim = NingunCasco
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.CascoAnim, 6))
 
                Case 2 'Escudo
                    
                    .Invent.EscudoEqpObjIndex = 0
                    .Invent.EscudoEqpSlot = 0
                    
                    If .Char.Escudo_Aura <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 3))
                        .Char.Escudo_Aura = 0
                    End If
                    
                    .Char.ShieldAnim = NingunEscudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.ShieldAnim, 5))
                     
                End Select 'End Obj.Subtipo
                
                                
        End Select
        
        
    Call UpdateUserInventario(UserIndex, slot, 2, 0) ' Accion 2, Desequipped
             
             
    End With
 
    Exit Sub

ErrHandler:
    Call RegistrarError(Err.Number, Err.description, "InvUsuario.Desequipar", Erl)
    Resume Next
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           Optional ByRef sMotivo As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo ErrHandler
    If ObjIndex = 0 Then Exit Function
    
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = 267
    
    Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer, _
                              Optional ByRef sMotivo As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************
    If ObjIndex = 0 Then Exit Function
    If ObjData(ObjIndex).Real = 1 Then
        If esArmada(UserIndex) Then
            FaccionPuedeUsarItem = esArmada(UserIndex)
        Else
            FaccionPuedeUsarItem = False

        End If

    ElseIf ObjData(ObjIndex).Caos = 1 Then

        If esCaos(UserIndex) Then
            FaccionPuedeUsarItem = esCaos(UserIndex)
        Else
            FaccionPuedeUsarItem = False

        End If
    ElseIf ObjData(ObjIndex).Milicia = 1 Then

        If esMili(UserIndex) Then
            FaccionPuedeUsarItem = esMili(UserIndex)
        Else
            FaccionPuedeUsarItem = False

        End If
    Else
        FaccionPuedeUsarItem = True

    End If
    
    If Not FaccionPuedeUsarItem Then sMotivo = 439

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)


    On Error GoTo ErrHandler

    'Equipa un item del inventario
    Dim Obj      As ObjData
    Dim ObjIndex As Integer
    Dim sMotivo  As Integer
    
    With UserList(UserIndex)
    
        ObjIndex = .Invent.Object(slot).ObjIndex
        Obj = ObjData(ObjIndex)
               
        If .flags.Privilegios And PlayerType.User Then
            If Obj.levelItem > 0 Then
                If .Stats.ELV < Obj.levelItem Then
                     Call WriteLocaleMsg(UserIndex, 268, Obj.levelItem)
                     Exit Sub
               End If
            End If
            
        End If
        
        Select Case Obj.OBJType
        
            Case eOBJType.otMonturas
            
                Call DoEquita(UserIndex, Obj, slot)
                
            Case eOBJType.otItemsMagicos
                
                If .Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, slot)
                    Exit Sub
                End If
            
                If .Invent.MagicIndex <> 0 Then
                    Call Desequipar(UserIndex, .Invent.MagicSlot)
                End If
            
                .Invent.MagicIndex = .Invent.Object(slot).ObjIndex
                .Invent.MagicSlot = slot
                .Invent.Object(slot).Equipped = 1
                
                Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
        
                If Obj.Aura <> 0 Then
                    .Char.Anillo_Aura = Obj.Aura
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Anillo_Aura, 6))
                End If
                        
                Select Case Obj.EfectoMagico
            
                    Case eMagicType.ModificaAtributo
                    
                        If Obj.QueAtributo <> 0 Then
                            .Stats.UserAtributos(Obj.QueAtributo) = .Stats.UserAtributos(Obj.QueAtributo) + Obj.CuantoAumento
                        End If
                         
                    Case eMagicType.ModificaSkill
                    
                        If Obj.QueSkill <> 0 Then
                            .Stats.UserSkills(Obj.QueSkill) = .Stats.UserSkills(Obj.QueSkill) + Obj.CuantoAumento
                        End If
                
                End Select
            
        
            Case eOBJType.otNudillos
        
                    If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                
                        'Si esta equipado lo quita
                        If .Invent.Object(slot).Equipped Then
                             Call Desequipar(UserIndex, slot)
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.WeaponEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                        End If
                    
                        If .Invent.NudiEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.NudiEqpSlot)
                        End If
                        
                        .Invent.Object(slot).Equipped = 1
                        
                        Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                            
                        .Invent.NudiEqpObjIndex = .Invent.Object(slot).ObjIndex
                        .Invent.NudiEqpSlot = slot
                        .Char.WeaponAnim = Obj.WeaponAnim
                            
                        'Mermas// actualizamos solo lo que ocupamos :P
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 8))
                             
                        If Obj.Aura <> 0 Then
                            .Char.Arma_Aura = Obj.Aura
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, 1))
                                            
                            If ObjData(.Invent.Object(slot).ObjIndex).SndEspecial > 0 Then 'Sonido de auras
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(.Invent.Object(slot).ObjIndex).SndEspecial, .Pos.X, .Pos.Y))
                            End If
                            
                        End If
                    
                    Else
            
                        Call WriteLocaleMsg(UserIndex, sMotivo)
            
                    End If
 
            Case eOBJType.otWeapon
 
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                
                    If .Invent.EscudoEqpObjIndex <> 0 Then
                    
                      If ObjData(.Invent.Object(slot).ObjIndex).DosManos = 1 Then
                      
                          If ObjData(.Invent.EscudoEqpObjIndex).DosManos = 0 Then
                              Call WriteLocaleMsg(UserIndex, 263)
                              Exit Sub
                        End If
                        
                      End If
                      
                    End If
                     
                    'Si esta equipado lo quita
                    If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
            
                    'Quitamos el elemento anterior
                    If .Invent.NudiEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.NudiEqpSlot)
                    End If
               
                    .Invent.Object(slot).Equipped = 1
                    Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                    
                    .Invent.WeaponEqpObjIndex = .Invent.Object(slot).ObjIndex
                    .Invent.WeaponEqpSlot = slot
                    .Char.WeaponAnim = Obj.WeaponAnim
                    
                    
                    If .flags.Montando = 0 And .flags.Navegando = 0 Then
                        'Mermas// actualizamos solo lo que ocupamos :P
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.WeaponAnim, 4))
                    End If
                     
                    If Obj.Aura <> 0 Then
                        .Char.Arma_Aura = Obj.Aura
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, 1))
                                        
                        If ObjData(.Invent.Object(slot).ObjIndex).SndEspecial > 0 Then 'Sonido de auras
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(.Invent.Object(slot).ObjIndex).SndEspecial, .Pos.X, .Pos.Y))
                        End If
                        
                    End If
                    
                Else
                
                    Call WriteLocaleMsg(UserIndex, sMotivo)
 
                End If
  
  
            Case eOBJType.otAnillo
  
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                
                    'Si esta equipado lo quita
                    If .Invent.Object(slot).Equipped Then
                        Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.AnilloEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
                    End If
                
                    .Invent.Object(slot).Equipped = 1
                    Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                    
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = slot

                Else
                
                    Call WriteLocaleMsg(UserIndex, sMotivo)
                    
                End If
 
    
            Case eOBJType.otFlechas

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                        
                    'Si esta equipado lo quita
                    If .Invent.Object(slot).Equipped Then
                        Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                    End If
                
                    .Invent.Object(slot).Equipped = 1
                    Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                     
                    .Invent.MunicionEqpObjIndex = ObjIndex
                    .Invent.MunicionEqpSlot = slot
                        
                Else
                
                    Call WriteLocaleMsg(UserIndex, sMotivo)
                    
                End If
 
            Case eOBJType.otArmadura
                
                Select Case Obj.SubTipo
                
                Case 0 'Armadura
 
                    If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And SexoPuedeUsarItem(UserIndex, ObjIndex, sMotivo) And CheckRazaUsaRopa(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    
                        If .Invent.Object(slot).Equipped Then
                            Call Desequipar(UserIndex, slot)
                            Exit Sub
                        End If
                        
 
                        If .Invent.ArmourEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                        End If
       
                        'Lo equipa
                        .Invent.Object(slot).Equipped = 1
                        Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                        
                        .Invent.ArmourEqpObjIndex = ObjIndex
                        .Invent.ArmourEqpSlot = slot
                        
                        .Char.body = Obj.Ropaje
                        .flags.Desnudo = 0
                    
                        If .flags.Montando = 0 And .flags.Navegando = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.body, 1))
                        End If
                        
                        If Obj.Aura <> 0 Then
                            .Char.Body_Aura = Obj.Aura
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, 2))
                                            
                            If ObjData(.Invent.Object(slot).ObjIndex).SndEspecial > 0 Then 'Sonido de auras
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(.Invent.Object(slot).ObjIndex).SndEspecial, .Pos.X, .Pos.Y))
                            End If
                            
                        End If
                        
                    Else
                    
                        Call WriteLocaleMsg(UserIndex, sMotivo)
 
                    End If

                Case 1
                
                    If ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, sMotivo) Then
                    
                        If .Invent.Object(slot).Equipped Then
                            Call Desequipar(UserIndex, slot)
                            Exit Sub
                        End If

                        If .Invent.CascoEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                        End If
       
                        .Invent.Object(slot).Equipped = 1
                        Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                        
                        .Invent.CascoEqpObjIndex = .Invent.Object(slot).ObjIndex
                        .Invent.CascoEqpSlot = slot
                        
                        .Char.CascoAnim = Obj.CascoAnim
                        
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.CascoAnim, 6))
                        
                        If Obj.Aura <> 0 Then
                            .Char.Head_Aura = Obj.Aura
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, 4))
                        End If
                        
                    Else
                    
                        Call WriteLocaleMsg(UserIndex, 265)

                    End If

                Case 2

                    If ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, sMotivo) Then
                        
                        If .Invent.WeaponEqpObjIndex <> 0 Then
                        
                          If ObjData(.Invent.WeaponEqpObjIndex).DosManos = 1 Then
                          
                            If ObjData(.Invent.Object(slot).ObjIndex).DosManos = 0 Then
                                Call WriteLocaleMsg(UserIndex, 428)
                                Exit Sub
                            End If
                                
                          End If
                          
                        End If
                        
                        'Si esta equipado lo quita
                        If .Invent.Object(slot).Equipped Then
                            Call Desequipar(UserIndex, slot)
                            Exit Sub
                        End If
                        
                         'Quita el anterior
                        If .Invent.EscudoEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                        End If
     
                        'Lo equipa
                        .Invent.Object(slot).Equipped = 1
                        Call UpdateUserInventario(UserIndex, slot, 1, 0) ' Accion 1, Equipped
                        
                        .Invent.EscudoEqpObjIndex = .Invent.Object(slot).ObjIndex
                        .Invent.EscudoEqpSlot = slot
                        
                        .Char.ShieldAnim = Obj.ShieldAnim
                        
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.ShieldAnim, 5))
                        
                        If Obj.Aura <> 0 Then
                            .Char.Escudo_Aura = Obj.Aura
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, 3))
                                            
                            If ObjData(.Invent.Object(slot).ObjIndex).SndEspecial > 0 Then 'Sonido de auras
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(.Invent.Object(slot).ObjIndex).SndEspecial, .Pos.X, .Pos.Y))
                            End If
                            
                        End If
                        
                    Else
                    
                      Call WriteLocaleMsg(UserIndex, 265)

                    End If
                    
                End Select
                
                
        End Select
        
    End With
    
 
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.description, "InvUsuario.EquiparInvItem", Erl)
    Resume Next
End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, _
                                  ItemIndex As Integer, _
                                  Optional ByRef sMotivo As Integer) As Boolean
    
    On Error GoTo ErrHandler
    
    If ItemIndex = 0 Then Exit Function

    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ItemIndex).RazaProhibida(1) <> 0 Then
            Dim i As Integer
            For i = 1 To NUMRAZAS
                If ObjData(ItemIndex).RazaProhibida(i) = UserList(UserIndex).raza Then
                    sMotivo = 266
                    CheckRazaUsaRopa = False
                    Exit Function
                End If
            Next i
        End If
    End If
    
    CheckRazaUsaRopa = True
    
Exit Function
ErrHandler:
    Call RegistrarError(Err.Number, Err.description, "InvUsuario.CheckRazaUsaRopa", Erl)
    Resume Next
End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)
    
    
    On Error GoTo ErrHandler
    
1    Dim Obj      As ObjData
2    Dim ObjIndex As Integer
3    Dim TargObj  As ObjData
4    Dim MiObj    As Obj
    
    With UserList(UserIndex)
    
5        If .Invent.Object(slot).Amount = 0 Then Exit Sub
        
6        Obj = ObjData(.Invent.Object(slot).ObjIndex)

         If .flags.Privilegios And PlayerType.User Then
7            If Obj.levelItem > 0 Then
8                If .Stats.ELV < Obj.levelItem Then
9                     Call WriteLocaleMsg(UserIndex, 268, Obj.levelItem)
10                    Exit Sub
11               End If
12            End If
         End If
        
13         ObjIndex = .Invent.Object(slot).ObjIndex
15        .flags.TargetObjInvIndex = ObjIndex
14        .flags.TargetObjInvSlot = slot
        
16        Select Case CInt(Obj.OBJType)

            Case eOBJType.otRuna
            
17                If .flags.Muerto Then
18                    Call WarpUserChar(UserIndex, Ciudades(.Hogar).Dead_Map, Ciudades(.Hogar).Dead_X, Ciudades(.Hogar).Dead_Y, True)

100                Else
                       If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 6 Or MapInfo(.Pos.Map).Pk = False Then
19                        Call WarpUserChar(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)

101                    Else

20                        Call WriteLocaleMsg(UserIndex, 468)

102                    End If
                    
                    
103                End If
                
                Exit Sub
            
            Case eOBJType.otBarcos
                Call DoNavega(UserIndex, Obj, slot)
                
            Case eOBJType.otPasajes
                        
21                If DeadCheck(UserIndex) Or .flags.Meditando Then Exit Sub
                
                'Se asegura que el target es un npc
22                If .flags.TargetNPC = 0 Then
23                    Call WriteLocaleMsg(UserIndex, 22)
24                    Exit Sub
25                End If
                
                'Si no es pirata no seguimos JUMP!
26                If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.transportadores Then
27                     Call WriteLocaleMsg(UserIndex, 22)
33                     Exit Sub
28                End If
                
                'Distancia ^^
29                If Distancia(Npclist(.flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
30                    Call WriteLocaleMsg(UserIndex, 8)
32                    Exit Sub
31                End If

                'No es de aca :$
34                If Not .Pos.Map = Obj.DesdeMap Then
35                      Call WriteLocaleMsg(UserIndex, 30)
36                      Exit Sub
37                End If
38
39                'Mapa inválido
40                If Not MapaValido(Obj.HastaMap) Then
41                  Call WriteLocaleMsg(UserIndex, 30)
42                    Exit Sub
43                End If
44
45                Call WarpUserChar(UserIndex, Obj.HastaMap, Obj.HastaX, Obj.HastaY, True)
46                Call WriteLocaleMsg(UserIndex, 31)
47
48                .Stats.MinAGU = 0
49                .Stats.MinHam = 0
50                Call WriteUpdateHungerAndThirst(UserIndex)
51
52                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributosBackUP(eAtributos.Agilidad)
53                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributosBackUP(eAtributos.Fuerza)
54                Call WriteUpdateStrenght(UserIndex)
55                Call WriteUpdateDexterity(UserIndex)
        
        
57                Call QuitarUserInvItem(UserIndex, slot, 1)
56                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
                
                 
              Case eOBJType.otUseOnce

59                If DeadCheck(UserIndex) Then Exit Sub
        
60                'Usa el item
61                .Stats.MinHam = .Stats.MinHam + Obj.MinHam

62                If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
63                .flags.Hambre = 0

64                Call WriteUpdateHambre(UserIndex)
                
                'Sonido
65                If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
66                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
67                Else
68                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
69                End If
                
                'Quitamos del inv el item
70                Call QuitarUserInvItem(UserIndex, slot, 1)
71                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
        
            Case eOBJType.otGuita

72                If DeadCheck(UserIndex) Then Exit Sub

74                .Stats.GLD = .Stats.GLD + .Invent.Object(slot).Amount
73                .Invent.Object(slot).Amount = 0
75                .Invent.Object(slot).ObjIndex = 0
76                .Invent.NroItems = .Invent.NroItems - 1
                
77                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
78                Call WriteUpdateGold(UserIndex)
                
            Case eOBJType.otWeapon

79                If DeadCheck(UserIndex) Then Exit Sub
                
80                If Not .Stats.MinSta > 0 Then
81                    Call WriteLocaleMsg(UserIndex, 93)
82                    Exit Sub
83                End If
                
                  Select Case Obj.proyectil
                  
                      Case 1 'Arco
                        
                        If .Invent.Object(slot).Equipped = 0 Then
84                            Call WriteLocaleMsg(UserIndex, 391)
87                            Exit Sub
88                        End If
    
                        If SeguroCheck(UserIndex, 1) Then Exit Sub
                        Call WriteWorkRequestTarget(UserIndex, eSkill.Proyectiles)
                        
                        Case 2 'Daga arrojadiza, arpon, shuriken, etc
                        
91                        If .Invent.Object(slot).Equipped = 0 Then
92                            Call WriteLocaleMsg(UserIndex, 391)
93                            Exit Sub
105                        End If

                        If SeguroCheck(UserIndex, 1) Then Exit Sub
                        Call WriteWorkRequestTarget(UserIndex, eSkill.ArmasArrojadizas)
                        
                        Case Else
                        
                            If .flags.TargetObj = Leña Then
                                If .Invent.Object(slot).ObjIndex = DAGA Then
                                    If .Invent.Object(slot).Equipped = 0 Then
                                        Call WriteLocaleMsg(UserIndex, 391)
                                        Exit Sub
                                    End If
                                    
                                    Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                                End If
                            Else
                                Call WriteLocaleMsg(UserIndex, 422)
                            End If
                            
                  End Select
                  
            Case eOBJType.otAnillo
                 
                If UserList(UserIndex).Invent.Object(slot).Equipped <> 0 Then
      
                    Select Case ObjIndex
                    
                    Case COSTURERO
                    
                        If .Invent.AnilloEqpObjIndex = COSTURERO Then
                        
                            Call EnivarObjTejibles(UserIndex)
                            Call WriteAbrirFormularios(UserIndex, 3)
                        Else
                            Call WriteLocaleMsg(UserIndex, 391)
                        End If
                    
                        
                    Case OLLA
                    
                        If .Invent.AnilloEqpObjIndex = OLLA Then
                        
                            Call EnivarObjalquimia(UserIndex)
                            Call WriteAbrirFormularios(UserIndex, 4)
                        Else
                            Call WriteLocaleMsg(UserIndex, 391)
                        End If
                    
                    Case SERRUCHO_CARPINTERO
                    
                        If .Invent.AnilloEqpObjIndex = SERRUCHO_CARPINTERO Then
                           Call EnivarObjConstruibles(UserIndex)
                           Call WriteAbrirFormularios(UserIndex, 2)
                        Else
                           Call WriteLocaleMsg(UserIndex, 391)
                        End If
    
                    End Select
            
                End If
 
            Case eOBJType.otPociones

                If DeadCheck(UserIndex) Then Exit Sub
                
                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                    Call WriteLocaleMsg(UserIndex, 469)
                    Exit Sub
                End If
                
                .flags.TomoPocion = True
                .flags.TipoPocion = Obj.SubTipo
                        
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                        Call WriteUpdateDexterity(UserIndex)

                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If

                         
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                        
                        Call WriteUpdateStrenght(UserIndex)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If

                         
                        
                    Case 3 'Pocion roja, restaura HP
                        'Usa el item
                        .Stats.MinHP = .Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.MinHP > .Stats.MaxHP Then .Stats.MinHP = .Stats.MaxHP
                        
                        Call WriteUpdateHP(UserIndex)

                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case 4 'Pocion azul, restaura MANA
                        'Usa el item
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV

                        If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                        Call WriteUpdateMana(UserIndex)

                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If

 
                     Case 5 'Lanzan hechizo
                        If Obj.LanzaHechizo > 0 Then
                             Dim Casteo As Boolean
                            .flags.TargetUser = UserIndex
                            Call HechizoEstadoUsuario(UserIndex, Casteo, Obj.LanzaHechizo)
                            .flags.TargetUser = 0
                            If Casteo = False Then Exit Sub
                        End If
                        
                    Case 6 'Scroll Intermundia
                        
                        If .Counters.TiempoDeMapeo = 0 Then
                            Call WarpUserChar(UserIndex, Ciudades(eCiudad.cIntermundia).Map, Ciudades(eCiudad.cIntermundia).X, Ciudades(eCiudad.cIntermundia).Y, True)
                        End If
                        
                    Case 7 ' Cambio de cara
                        
                        If .flags.Navegando = 1 Then
                            Call WriteLocaleMsg(UserIndex, 20)
                            Exit Sub
                        End If
                        
                        If .flags.Montando = 1 Then
                            Call WriteLocaleMsg(UserIndex, 21)
                            Exit Sub
                        End If
                        
                        Call ChangeHead(UserIndex)
                        
                    Case 8 ' Cambio de sexo
                        
                        If .flags.Navegando = 1 Then
                            Call WriteLocaleMsg(UserIndex, 20)
                            Exit Sub
                        End If
                        
                        If .flags.Montando = 1 Then
                            Call WriteLocaleMsg(UserIndex, 21)
                            Exit Sub
                        End If
                        
                        Call DarCuerpoNuevo(UserIndex)
                        
                    Case 9 ' Nareth
                        If UserList(CharList(UserIndex)).Char.ParticulaFx <> 0 Then Exit Sub
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, 23, -1, False, True))
                    
                    
                    Case 13 ' Adquirir creditos
                    
                        If Obj.CuantoAumento > 0 Then
                            .Donador.CreditoDonador = .Donador.CreditoDonador + Obj.CuantoAumento
                             Call WriteLocaleMsg(UserIndex, 470, Obj.CuantoAumento & "%" & .Donador.CreditoDonador)
                        End If
                        
                End Select
                              
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, slot, 1)
                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
            
            Case eOBJType.otbebidas
    
                If DeadCheck(UserIndex) Then Exit Sub
    
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
    
                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                
                .flags.Sed = 0
                
                Call WriteUpdateSed(UserIndex)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, slot, 1)
                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
                
      
            Case eOBJType.otLlaves

                If DeadCheck(UserIndex) Then Exit Sub
                
                If .flags.TargetObj = 0 Then Exit Sub
                
                TargObj = ObjData(.flags.TargetObj)

                '¿El objeto clickeado es una puerta?
                If TargObj.OBJType = eOBJType.otPuertas Then

                    '¿Esta cerrada?
                    If TargObj.Cerrada = 1 Then

                        '¿Cerrada con llave?
                        If TargObj.Llave > 0 Then
                            If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                 Call WriteLocaleMsg(UserIndex, 472)
                                 Exit Sub
                            Else
                                 Call WriteLocaleMsg(UserIndex, 473)
                                 Exit Sub
                            End If

                        Else

                            If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                Call WriteLocaleMsg(UserIndex, 474)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Exit Sub
                            Else
                                Call WriteLocaleMsg(UserIndex, 475)
                                Exit Sub
                            End If

                        End If

                    Else
                    
                        Call WriteLocaleMsg(UserIndex, 476)
                        
                        Exit Sub

                    End If

                End If
            
            Case eOBJType.otBotellaVacia

                If DeadCheck(UserIndex) Then Exit Sub

                If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteLocaleMsg(UserIndex, 392)
                    Exit Sub
                End If

                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(slot).ObjIndex).IndexAbierta
                
                Call QuitarUserInvItem(UserIndex, slot, 1)

                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                
                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
            
            Case eOBJType.otBotellaLlena

                If DeadCheck(UserIndex) Then Exit Sub

                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                
                .flags.Sed = 0
                
                Call WriteUpdateSed(UserIndex)
                
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(slot).ObjIndex).IndexCerrada
                
                Call QuitarUserInvItem(UserIndex, slot, 1)

                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                
                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
                
                If .flags.AdminInvisible = 1 Then
                     Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                 Else
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                 End If
            
            Case eOBJType.otPergaminos

                If DeadCheck(UserIndex) Then Exit Sub
                
                If .Stats.MaxMAN > 0 Then
                    Call AgregarHechizo(UserIndex, slot)
                    Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
                Else
                    Call WriteLocaleMsg(UserIndex, 393)
                End If

            Case eOBJType.otMinerales

                If DeadCheck(UserIndex) Then Exit Sub
                  
                Call WriteWorkRequestTarget(UserIndex, FundirMetal)
                
                .flags.Lingoteando = slot
 
            Case eOBJType.otBolsas
            
                If DeadCheck(UserIndex) Then Exit Sub
       
                .Stats.GLD = .Stats.GLD + ObjData(ObjIndex).Valor
                Call WriteLocaleMsg(UserIndex, 471, Obj.Name & "%" & Obj.Valor)
                Call WriteUpdateGold(UserIndex)

                Call QuitarUserInvItem(UserIndex, slot, 1)
                Call UpdateUserInventario(UserIndex, slot, 3, .Invent.Object(slot).Amount) ' Accion 3, Cantidad
        
            Case eOBJType.otInstrumentos

                If DeadCheck(UserIndex) Then Exit Sub
                
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
                
                'REGALO1!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Case eOBJType.otRegalos
    Dim Miobj1 As Obj
    Dim MiObj2 As Obj
    
    MiObj.ObjIndex = 1081
    Miobj1.ObjIndex = 1280
    MiObj2.ObjIndex = 1604
    
    MiObj.Amount = 1
    Miobj1.Amount = 1
    MiObj2.Amount = 1
    
        If UserList(UserIndex).flags.Muerto = 1 Then
        
               Call WriteConsoleMsg(UserIndex, "Clanes Deshabilitados temporalmente.", FontTypeNames.FONTTYPE_INFO)
        
        Exit Sub
End If

If MeterItemEnInventario(UserIndex, MiObj) Then

  If MeterItemEnInventario(UserIndex, Miobj1) Then
  
   If MeterItemEnInventario(UserIndex, MiObj2) Then
   
   End If
   
  End If
  
End If

Call QuitarUserInvItem(UserIndex, slot, 1)
Call UpdateUserInv(False, UserIndex, slot)
                
                
                
                
 
            End Select
    
                
    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.description, "InvUsuario.UseInvItem", Erl)
    Resume Next
End Sub

Private Function ChangeHead(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler

    Dim NewHead As Integer
     
    Dim UserGenero As Integer
    Dim UserRaza As Integer
    
    With UserList(UserIndex)
    
    UserGenero = .Genero
    UserRaza = .raza

    Select Case UserGenero
    
       Case eGenero.Hombre
   
            Select Case UserRaza
            
                Case eRaza.Humano
                    NewHead = RandomNumber(1, 30)
                Case eRaza.Elfo
                    NewHead = RandomNumber(101, 120)
                Case eRaza.Drow
                    NewHead = RandomNumber(201, 213)
                Case eRaza.enano
                    NewHead = RandomNumber(301, 313)
                Case eRaza.gnomo
                    NewHead = RandomNumber(401, 410)
                Case eRaza.Orco
                    NewHead = RandomNumber(501, 514)
                
            End Select
        
       Case eGenero.Mujer
       
            Select Case UserRaza
            
                Case eRaza.Humano
                    NewHead = RandomNumber(70, 80)
                Case eRaza.Elfo
                    NewHead = RandomNumber(170, 189)
                Case eRaza.Drow
                    NewHead = RandomNumber(270, 278)
                Case eRaza.gnomo
                    NewHead = RandomNumber(470, 481)
                Case eRaza.enano
                    NewHead = RandomNumber(370, 373)
                Case eRaza.Orco
                    NewHead = RandomNumber(570, 573)
                    
            End Select
            
    End Select

    .Char.Head = NewHead
    .OrigChar.Head = NewHead
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeSlot(.Char.CharIndex, .Char.Head, 2))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 30, 0))
    
    End With
    
    Exit Function

ErrorHandler:
    Call RegistrarError(Err.Number, Err.description, "InvUsuario.ChangeHead", Erl)
    Resume Next
End Function
 
Private Function DarCuerpoNuevo(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler
    
    With UserList(UserIndex)
 
    Dim NewHead As Integer
    Dim NewBody As Integer
    Dim UserRaza As Integer
    Dim UserGenero As Integer
    
    UserGenero = .Genero
    UserRaza = .raza
 
    If .Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        .Char.Body_Aura = 0
    End If
    
    'desequipar arma
    If .Invent.NudiEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.NudiEqpSlot)
        .Char.Arma_Aura = 0
    End If
        
    'desequipar nudillo
    If .Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, 1))
        .Char.Arma_Aura = 0
    End If
    
    'desequipar casco
    If .Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
        .Char.Head_Aura = 0
    End If
    
    'desequipar herramienta
    If .Invent.AnilloEqpSlot > 0 Then
        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
    End If
    
    'desequipar municiones
    If .Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
    End If
    
    'desequipamos items macigos
    If .Invent.MagicIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.MagicSlot)
        .Char.Anillo_Aura = 0
    End If
    
    'desequipar escudo
    If .Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        .Char.Escudo_Aura = 0
    End If
 
    Select Case UserGenero
      
       Case eGenero.Hombre
    
         .Genero = eGenero.Mujer
        
            Select Case UserRaza
                Case eRaza.Humano
                    NewHead = RandomNumber(70, 80)
                    NewBody = 1
                Case eRaza.Elfo
                    NewHead = RandomNumber(170, 189)
                    NewBody = 2
                Case eRaza.Drow
                    NewHead = RandomNumber(270, 278)
                    NewBody = 3
                Case eRaza.gnomo
                    NewHead = RandomNumber(470, 481)
                    NewBody = 138
                Case eRaza.enano
                    NewHead = RandomNumber(370, 373)
                    NewBody = 138
                Case eRaza.Orco
                    NewHead = RandomNumber(570, 573)
                    NewBody = 253
            End Select
            
            
        Case eGenero.Mujer
        
            .Genero = eGenero.Hombre
            
            Select Case UserRaza
                Case eRaza.Humano
                    NewHead = RandomNumber(1, 30)
                    NewBody = 1
                Case eRaza.Elfo
                    NewHead = RandomNumber(101, 120)
                    NewBody = 2
                Case eRaza.Drow
                    NewHead = RandomNumber(201, 213)
                    NewBody = 3
                Case eRaza.enano
                    NewHead = RandomNumber(301, 313)
                    NewBody = 52
                Case eRaza.gnomo
                    NewHead = RandomNumber(401, 410)
                    NewBody = 52
                Case eRaza.Orco
                    NewHead = RandomNumber(501, 514)
                    NewBody = 252
            End Select
            
        End Select
        
    .OrigChar.body = NewBody
    
    .Char.ShieldAnim = NingunEscudo
    .Char.WeaponAnim = NingunArma
    .Char.CascoAnim = NingunCasco
    .Char.Head = NewHead
    .OrigChar.Head = NewHead
    
    Call DarCuerpoDesnudo(UserIndex)
    
    Call ChangeUserCharTodo(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 30, 0))
    
    End With
    
    Exit Function

ErrorHandler:
    Call RegistrarError(Err.Number, Err.description, "InvUsuario.DarCuerpoNuevo", Erl)
    Resume Next
End Function



Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteBlacksmithWeapons(UserIndex)

End Sub

Sub EnivarCascosConstruibles(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Shermie80
    '
    '***************************************************

    Call WriteBlacksmithHelmet(UserIndex)

End Sub
 
Sub EnivarEscudosConstruibles(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Shermie80
    '
    '***************************************************

    Call WriteBlacksmithShield(UserIndex)

End Sub

Sub EnivarObjTejibles(ByVal UserIndex As Integer)

    Call WriteTejiblesObjects(UserIndex)

End Sub
 
Sub EnivarObjalquimia(ByVal UserIndex As Integer)

  Call WriteAlquimiaObjects(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 6 Or MapInfo(.Pos.Map).Pk = False Then Exit Sub  'UserList(userindex).Pos.map = 1 Or UserList(userindex).Pos.map = 34 _
        Or UserList(userindex).Pos.map = 184 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 185 _
        Or UserList(userindex).Pos.map = 49 Or UserList(userindex).Pos.map = 194 Or UserList(userindex).Pos.map = 179 _
        Or UserList(userindex).Pos.map = 62 Or UserList(userindex).Pos.map = 64 Or UserList(userindex).Pos.map = 63 _
        Or UserList(userindex).Pos.map = 181 Or UserList(userindex).Pos.map = 180 Or UserList(userindex).Pos.map = 112 _
        Or UserList(userindex).Pos.map = 61 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 111 _
        Or UserList(userindex).Pos.map = 59 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 60 _
        Or UserList(userindex).Pos.map = 58 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 364 _
        Or UserList(userindex).Pos.map = 217 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 218 _
        Or UserList(userindex).Pos.map = 20 Or UserList(userindex).Pos.map = 208 Or UserList(userindex).Pos.map = 37 _
        Then Exit Sub
 
        Call TirarTodosLosItems(UserIndex)

    End With

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And (.Caos <> 1 Or .NoSeCae = 0) And .OBJType <> eOBJType.otLlaves _
                And .OBJType <> eOBJType.otBarcos And .OBJType <> eOBJType.otMonturas And .NoSeCae = 0

    End With

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
    '***************************************************

    Dim i         As Byte
    Dim NuevaPos  As WorldPos
    Dim MiObj     As Obj
    Dim ItemIndex As Integer
    Dim DropAgua  As Boolean
 
    
With UserList(UserIndex)

    Dim Carro As Byte
    Dim Minerales As Integer
    Dim Porc As Byte
    Dim Hierro As Integer, Plata As Integer, oro As Integer
    
    Dim Rycanfricio As Obj
    Rycanfricio.ObjIndex = RYKAN
    Rycanfricio.Amount = 1

        If TieneObjetos(1601, 1, UserIndex) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, Rycanfricio)
            Call QuitarObjetos(RYKAN, 1, UserIndex)
            Exit Sub
        End If
        
    Dim sacri3 As Obj
    Dim sacri2 As Obj
    Dim sacri1 As Obj
    
    sacri3.ObjIndex = 1081
    sacri2.ObjIndex = 1498
    sacri2.Amount = 1
    sacri1.ObjIndex = 1499
    sacri1.Amount = 1
        If TieneObjetos(1081, 1, UserIndex) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, sacri2)
            Call QuitarObjetos(1081, 1, UserIndex)
            Exit Sub
        ElseIf TieneObjetos(1489, 1, UserIndex) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, sacri1)
            Call QuitarObjetos(1498, 1, UserIndex)
            Exit Sub
        ElseIf TieneObjetos(1499, 1, UserIndex) Then
            Call QuitarObjetos(1499, 1, UserIndex)
            Exit Sub
        End If

    Carro = Have_Obj_Slot(CARROMINERO, UserIndex)
    If Carro > 0 Then
        Hierro = Have_Obj_To_Slot(iMinerales.HierroCrudo, Carro, UserIndex)
        Plata = Have_Obj_To_Slot(iMinerales.PlataCruda, Carro, UserIndex)
        oro = Have_Obj_To_Slot(iMinerales.OroCrudo, Carro, UserIndex)
        
        If Hierro > 0 Then
            Porc = Porc + 1
        End If
        
        If Plata > 0 Then
            Porc = Porc + 1
        End If
        
        If oro > 0 Then
            Porc = Porc + 1
        End If
        
        If Hierro > 0 Then Hierro = Porcentaje(Hierro, (100 - ObjData(UserList(UserIndex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        If Plata > 0 Then Plata = Porcentaje(Plata, (100 - ObjData(UserList(UserIndex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        If oro > 0 Then oro = Porcentaje(oro, (100 - ObjData(UserList(UserIndex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        
        If Porc > 0 Then
            For i = 1 To Carro
                If UserList(UserIndex).Invent.Object(i).ObjIndex = iMinerales.HierroCrudo Then
                    If Hierro > 0 Then
                        TirarObjeto UserIndex, i, Hierro
                        Hierro = Hierro - IIf(UserList(UserIndex).Invent.Object(i).Amount > Hierro, Hierro, UserList(UserIndex).Invent.Object(i).Amount)
                    End If
                ElseIf UserList(UserIndex).Invent.Object(i).ObjIndex = iMinerales.PlataCruda Then
                    If Plata > 0 Then
                        TirarObjeto UserIndex, i, Plata
                        Plata = Plata - IIf(UserList(UserIndex).Invent.Object(i).Amount > Plata, Plata, UserList(UserIndex).Invent.Object(i).Amount)
                    End If
                ElseIf UserList(UserIndex).Invent.Object(i).ObjIndex = iMinerales.OroCrudo Then
                    If oro > 0 Then
                        TirarObjeto UserIndex, i, oro
                        oro = oro - IIf(UserList(UserIndex).Invent.Object(i).Amount > oro, oro, UserList(UserIndex).Invent.Object(i).Amount)
                    End If
                End If
            Next i
        End If
    End If

        For i = 1 To MAX_INVENTORY_SLOTS
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex

                    DropAgua = True

                    ' Es pirata?
                    If .Clase = eClass.Sastre Then

                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 476 Then

                            ' Limitación por nivel, después dropea normalmente
                            If .Stats.ELV >= 20 And .Stats.ELV <= 25 Then
                                ' No dropea en agua
                                DropAgua = False

                            End If

                        End If

                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    If TieneObjetos(880, 1, UserIndex) And ObjData(.Invent.Object(i).ObjIndex).OBJType = eOBJType.otMinerales Then
                     Dim Cantidad As Integer
                    Cantidad = (MiObj.Amount * 30) / 100
                    Call DropObj(UserIndex, i, Cantidad, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    Else
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    
                      End If
                    
                    End If

                End If

            End If

        Next i

    End With

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 23/11/2009
    '07/11/09: Pato - Fix bug #2819911
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim i         As Byte
    Dim NuevaPos  As WorldPos
    Dim MiObj     As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 6 Or MapInfo(.Pos.Map).Pk = False Then Exit Sub   ' UserList(userindex).Pos.map = 1 Or UserList(userindex).Pos.map = 34 _
        Or UserList(userindex).Pos.map = 184 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 185 _
        Or UserList(userindex).Pos.map = 49 Or UserList(userindex).Pos.map = 194 Or UserList(userindex).Pos.map = 179 _
        Or UserList(userindex).Pos.map = 62 Or UserList(userindex).Pos.map = 64 Or UserList(userindex).Pos.map = 63 _
        Or UserList(userindex).Pos.map = 181 Or UserList(userindex).Pos.map = 180 Or UserList(userindex).Pos.map = 112 _
        Or UserList(userindex).Pos.map = 61 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 111 _
        Or UserList(userindex).Pos.map = 59 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 60 _
        Or UserList(userindex).Pos.map = 58 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 364 _
        Or UserList(userindex).Pos.map = 217 Or UserList(userindex).Pos.map = 183 Or UserList(userindex).Pos.map = 218 _
        Or UserList(userindex).Pos.map = 20 Or UserList(userindex).Pos.map = 208 Or UserList(userindex).Pos.map = 37 _
        Then Exit Sub
        
         
               For i = 1 To MAX_INVENTORY_SLOTS

            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True

                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

        Next i

    End With

End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If ObjIndex > 0 Then
        getObjType = ObjData(ObjIndex).OBJType

    End If
    
End Function

Function Have_Obj_To_Slot(ByVal ItemIndex As Integer, ByVal slot As Byte, ByVal UserIndex As Integer) As Long
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long

    For i = 1 To slot
        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Have_Obj_To_Slot = Have_Obj_To_Slot + UserList(UserIndex).Invent.Object(i).Amount
        End If
    Next i
        
End Function
Function Have_Obj_Slot(ByVal ItemIndex As Integer, ByVal UserIndex As Integer) As Integer
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long

    For i = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Have_Obj_Slot = i
        End If
    Next i
        
End Function

Sub TirarObjeto(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal cant As Integer)
    Dim MiObj As Obj
    Dim NuevaPos As WorldPos
    
  
    
    If cant > UserList(UserIndex).Invent.Object(slot).Amount Then _
        cant = UserList(UserIndex).Invent.Object(slot).Amount
    'Creo el Obj
    MiObj.Amount = cant
    MiObj.ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
    
    If UserList(UserIndex).Clase = eClass.Mercenario And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
        Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
    Else
        Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
    End If
                
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call DropObj(UserIndex, slot, cant, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
    End If
End Sub
Public Function DeadCheck(ByVal UserIndex As Integer) As Boolean
    
    On Error GoTo ErrorHandler
    
    If UserList(UserIndex).flags.Muerto Then
        Call WriteLocaleMsg(UserIndex, 77, "", 0, 12)
        DeadCheck = True
        Exit Function
    End If

    DeadCheck = False
    
    Exit Function
     
ErrorHandler:

    Call RegistrarError(Err.Number, Err.description, "General.DeadCheck", Erl)
    Resume Next
 
End Function
Public Function OnlineCheck(ByVal UserIndex As Integer, ByVal tUser As Integer) As Boolean
    
    On Error GoTo ErrorHandler

1    If tUser <= 0 Then  'Si está online
2        Call WriteLocaleMsg(UserIndex, 75)
3        OnlineCheck = False
4        Exit Function
5    End If
    
6    OnlineCheck = True
     
     Exit Function
     
ErrorHandler:

    Call RegistrarError(Err.Number, Err.description, "General.OnlineCheck", Erl)
    Resume Next
End Function

Public Function SeguroCheck(ByVal UserIndex As Integer, ByVal TipoSeguro As Byte) As Boolean
    
    On Error GoTo ErrorHandler
    
    Select Case TipoSeguro
    
        Case 1 'Seguro de ataque
    
            If Not UserList(UserIndex).flags.ModoCombate Then 'Si lo tiene desactivado damos un True y mandamos msj
               Call WriteLocaleMsg(UserIndex, 102)
               SeguroCheck = True
               Exit Function
            End If
    
    End Select

    SeguroCheck = False
    
    Exit Function
     
ErrorHandler:

    Call RegistrarError(Err.Number, Err.description, "General.SeguroCheck", Erl)
    Resume Next
End Function
