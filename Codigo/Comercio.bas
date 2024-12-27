Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio

    Compra = 1
    Venta = 2

End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, _
                    ByVal UserIndex As Integer, _
                    ByVal npcindex As Integer, _
                    ByVal slot As Integer, _
                    ByVal Cantidad As Integer)
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
    '  - 06/13/08 (NicoNZ)
    '*************************************************
    Dim Precio As Long
    Dim Objeto As Obj
    
    If Cantidad < 1 Or slot < 1 Then Exit Sub
    
        If UserList(UserIndex).flags.Montando = 1 Then
            If UserList(UserIndex).Invent.MonturaSlot = slot Then
                Call WriteConsoleMsg(UserIndex, "No podes vender tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
 
    
    Select Case Modo
    
    Case eModoComercio.Compra
    
    If slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    If Cantidad > MAX_INVENTORY_OBJS Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
    Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados ítems:" & Cantidad)
    UserList(UserIndex).flags.Ban = 1
    Call WriteSendMsgBox(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
    Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
    End If
    If Not Npclist(npcindex).Invent.Object(slot).Amount > 0 Then Exit Sub

        If Cantidad > Npclist(npcindex).Invent.Object(slot).Amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(slot).Amount
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = Npclist(npcindex).Invent.Object(slot).ObjIndex
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        Precio = CLng((ObjData(Npclist(npcindex).Invent.Object(slot).ObjIndex).Valor / Descuento(UserIndex) * Cantidad) + 0.5)

        If UserList(UserIndex).Stats.GLD < Precio Then
            Call WriteLocaleMsg(UserIndex, 26)
            Exit Sub
        End If
        
        If Not MeterItemEnInventario(UserIndex, Objeto) Then Exit Sub
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
        Call WriteUpdateGold(UserIndex)
        
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(slot), Cantidad)
        
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(npcindex).Numero, "obj" & slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(UserIndex).Name & " compró " & ObjData(Objeto.ObjIndex).Name)
        End If
    
    Case eModoComercio.Venta
    
        If Cantidad > UserList(UserIndex).Invent.Object(slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
        If Objeto.ObjIndex = 0 Then Exit Sub
        
        If ObjData(Objeto.ObjIndex).Newbie = 1 Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf (Npclist(npcindex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(npcindex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then

            Call WriteLocaleMsg(UserIndex, 148)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Shop > 0 Then
            Call WriteLocaleMsg(UserIndex, 148)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then
            If Npclist(npcindex).Name <> "SR" Then
                Call WriteLocaleMsg(UserIndex, 148)
                Exit Sub
            End If
        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

            If Npclist(npcindex).Name <> "SC" Then
                Call WriteLocaleMsg(UserIndex, 148)
                Exit Sub

            End If
                ElseIf ObjData(Objeto.ObjIndex).Milicia > 0 Then
            If Npclist(npcindex).Name <> "SM" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras de la Milicia solo pueden ser vendidas a los sastres milicianos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf UserList(UserIndex).Invent.Object(slot).Amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf slot < LBound(UserList(UserIndex).Invent.Object()) Or slot > UBound(UserList(UserIndex).Invent.Object( _
                )) Then
            Exit Sub
        End If
        
        Call QuitarUserInvItem(UserIndex, slot, Cantidad)
        Call UpdateUserInv(False, UserIndex, slot)
        
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
 
        
        If OroLleno(UserIndex, UserList(UserIndex).Stats.GLD, Precio) Then
        Call WriteConsoleMsg(UserIndex, "Tienes la cantidad máxima de oro que puedes tener. No has obtenido oro", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
        End If
 
              
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
        If UserList(UserIndex).Stats.GLD > MAXORO Then UserList(UserIndex).Stats.GLD = MAXORO
        
        Call WriteUpdateGold(UserIndex)
        
        Dim NpcSlot As Integer
        NpcSlot = SlotEnNPCInv(npcindex, Objeto.ObjIndex, Objeto.Amount)
        
        If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            Npclist(npcindex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            Npclist(npcindex).Invent.Object(NpcSlot).Amount = Npclist(npcindex).Invent.Object(NpcSlot).Amount + _
                    Objeto.Amount

            If Npclist(npcindex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                Npclist(npcindex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS

            End If
        Call UpdateNpcInvToAll(False, npcindex, NpcSlot)
        End If

        
    End Select
 
    Call SubirSkill(UserIndex, eSkill.comerciar)

End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    Call UpdateNpcInv(True, UserIndex, UserList(UserIndex).flags.TargetNPC, 0)
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex)

End Sub

Private Function SlotEnNPCInv(ByVal npcindex As Integer, _
                              ByVal Objeto As Integer, _
                              ByVal Cantidad As Integer) As Integer
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    SlotEnNPCInv = 1

    Do Until Npclist(npcindex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto And Npclist(npcindex).Invent.Object( _
            SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1

        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until Npclist(npcindex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
            SlotEnNPCInv = SlotEnNPCInv + 1

            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(npcindex).Invent.NroItems = Npclist( _
                npcindex).Invent.NroItems + 1
    
    End If
    
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.comerciar) / 100

End Function

''
' Update the inventory of the Npc to the user
'
' @param updateAll if is needed to update all
' @param userIndex The index of the User
' @param npcIndex The index of the NPC
' @param slot The slot to update

Private Sub UpdateNpcInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal npcindex As Integer, ByVal slot As Byte)
'***************************************************
    Dim Obj As Obj
    Dim loopc As Byte
    Dim desc As Single
    Dim val As Single
    
    desc = Descuento(UserIndex)
    
    'Actualiza un solo slot
    If Not UpdateAll Then
        With Npclist(npcindex).Invent.Object(slot)
            Obj.ObjIndex = .ObjIndex
            Obj.Amount = .Amount
            
            If .ObjIndex > 0 Then
                val = (ObjData(.ObjIndex).Valor) / desc
            End If
            
            Call WriteChangeNPCInventorySlot(UserIndex, slot, Obj, val)
        End With
    Else
    'Actualiza todos los slots
        For loopc = 1 To MAX_INVENTORY_SLOTS
            With Npclist(npcindex).Invent.Object(loopc)
                Obj.ObjIndex = .ObjIndex
                Obj.Amount = .Amount
                
                If .ObjIndex > 0 Then
                    val = (ObjData(.ObjIndex).Valor) / desc
                End If
                
                Call WriteChangeNPCInventorySlot(UserIndex, loopc, Obj, val)
            End With
        Next loopc
    End If
End Sub


''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El número de objeto al cual le calculamos el precio de venta

Public Function SalePrice(ByVal ObjIndex As Integer) As Single

    '*************************************************
    'Author: Nicolás (NicoNZ)
    '
    '*************************************************
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    
    SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA

End Function

' Update the inventory of the Npc to all users trading with him
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC
' @param slot The slot to update

Public Sub UpdateNpcInvToAll(ByVal UpdateAll As Boolean, ByVal npcindex As Integer, ByVal slot As Byte)
'***************************************************
    Dim loopc As Byte
    
    ' Recorremos todos los usuarios
    For loopc = 1 To LastUser
        With UserList(loopc)
            ' Si esta comerciando
            If .flags.Comerciando Then
                ' Si el ultimo NPC que cliqueo es el que hay que actualizar
                If .flags.TargetNPC = npcindex Then
                    ' Actualizamos el inventario del NPC
                    Call UpdateNpcInv(UpdateAll, loopc, npcindex, slot)
                End If
            End If
        End With
    Next
End Sub
