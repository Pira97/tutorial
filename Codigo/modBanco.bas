Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
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

Option Explicit

Sub IniciarDeposito(ByVal UserIndex As Integer, ByVal goliath As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler
    If goliath Then
        Call WriteBankInit(UserIndex, 1)
    Else
    'Hacemos un Update del inventario del usuario
    Call UpdateBanUserInv(True, UserIndex, 0)
    
    'Actualizamos el dinero
    Call WriteUpdateGold(UserIndex)
    
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    Call WriteBankInit(UserIndex, 0)
     
    UserList(UserIndex).flags.Comerciando = True
    End If
ErrHandler:

End Sub

Sub SendBanObj(UserIndex As Integer, slot As Byte, Object As UserObj)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    UserList(UserIndex).BancoInvent.Object(slot) = Object

    Call WriteChangeBankSlot(UserIndex, slot)

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, _
                     ByVal UserIndex As Integer, _
                     ByVal slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NullObj As UserObj
    Dim loopc   As Byte

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .BancoInvent.Object(slot).ObjIndex > 0 Then
                Call SendBanObj(UserIndex, slot, .BancoInvent.Object(slot))
            Else
                Call SendBanObj(UserIndex, slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
            For loopc = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
                If .BancoInvent.Object(loopc).ObjIndex > 0 Then
                    Call SendBanObj(UserIndex, loopc, .BancoInvent.Object(loopc))
                Else
                    Call SendBanObj(UserIndex, loopc, NullObj)

                End If

            Next loopc

        End If

    End With

End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, _
                   ByVal i As Integer, _
                   ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    If Cantidad < 1 Then Exit Sub
    
    If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
        If Cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then Cantidad = UserList( _
                UserIndex).BancoInvent.Object(i).Amount
        'Agregamos el obj que compro al inventario
        Call UserReciveObj(UserIndex, CInt(i), Cantidad)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        'Actualizamos el banco
        Call UpdateBanUserInv(True, UserIndex, 0)

    End If
       
    'Actualizamos la ventana de comercio
    Call UpdateVentanaBanco(UserIndex)

ErrHandler:

End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, _
                  ByVal ObjIndex As Integer, _
                  ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim slot As Integer
    Dim obji As Integer

    With UserList(UserIndex)

        If .BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub
    
        obji = .BancoInvent.Object(ObjIndex).ObjIndex
    
        '¿Ya tiene un objeto de este tipo?
        slot = 1

        Do Until .Invent.Object(slot).ObjIndex = obji And .Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
            slot = slot + 1
        If slot > MAX_INVENTORY_SLOTS Then
            Exit Do
        End If
    Loop
   
    'Sino se fija por un slot vacio
    If slot > MAX_INVENTORY_SLOTS Then
        slot = 1
        Do Until .Invent.Object(slot).ObjIndex = 0
            slot = slot + 1
 
            If slot > MAX_INVENTORY_SLOTS Then
                Call WriteLocaleMsg(UserIndex, 25)
                Exit Sub
            End If
        Loop
        .Invent.NroItems = .Invent.NroItems + 1
    End If
        'Mete el obj en el slot
        If .Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(slot).ObjIndex = obji
            .Invent.Object(slot).Amount = .Invent.Object(slot).Amount + Cantidad
        
            Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Else
            Call WriteLocaleMsg(UserIndex, 25)

        End If

    End With

End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, _
                       ByVal slot As Byte, _
                       ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ObjIndex As Integer

    With UserList(UserIndex)
        ObjIndex = .BancoInvent.Object(slot).ObjIndex

        'Quita un Obj

        .BancoInvent.Object(slot).Amount = .BancoInvent.Object(slot).Amount - Cantidad
    
        If .BancoInvent.Object(slot).Amount <= 0 Then
            .BancoInvent.NroItems = .BancoInvent.NroItems - 1
            .BancoInvent.Object(slot).ObjIndex = 0
            .BancoInvent.Object(slot).Amount = 0

        End If

    End With
    
End Sub

Sub UpdateVentanaBanco(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteBankOK(UserIndex)

End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, _
                     ByVal Item As Integer, _
                     ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
        If Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList( _
                UserIndex).Invent.Object(Item).Amount
        
        'Agregamos el obj que deposita al banco
        Call UserDejaObj(UserIndex, CInt(Item), Cantidad)
        
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        
        'Actualizamos el inventario del banco
        Call UpdateBanUserInv(True, UserIndex, 0)

    End If
    
    'Actualizamos la ventana del banco
    Call UpdateVentanaBanco(UserIndex)
ErrHandler:

End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, _
                ByVal ObjIndex As Integer, _
                ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim slot As Integer
    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Sub
    
    With UserList(UserIndex)
        obji = .Invent.Object(ObjIndex).ObjIndex
        
        '¿Ya tiene un objeto de este tipo?
        slot = 1

        Do Until .BancoInvent.Object(slot).ObjIndex = obji And .BancoInvent.Object(slot).Amount + Cantidad <= _
                MAX_INVENTORY_OBJS
            slot = slot + 1
            
            If slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do

            End If

        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If slot > MAX_BANCOINVENTORY_SLOTS Then
            slot = 1

            Do Until .BancoInvent.Object(slot).ObjIndex = 0
                slot = slot + 1
                
                If slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteLocaleMsg(UserIndex, 149)
                    Exit Sub

                End If

            Loop
            
            .BancoInvent.NroItems = .BancoInvent.NroItems + 1

        End If
        
        If slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
            If .BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
                
                'Menor que MAX_INV_OBJS
                .BancoInvent.Object(slot).ObjIndex = obji
                .BancoInvent.Object(slot).Amount = .BancoInvent.Object(slot).Amount + Cantidad
                
                Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
            Else
                Call WriteLocaleMsg(UserIndex, 149)

            End If

        End If

    End With

End Sub

Sub SendUserBovedaTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Integer

    Call WriteConsoleMsg(SendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(SendIndex, "Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", _
            FontTypeNames.FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object( _
                    j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, _
                    FontTypeNames.FONTTYPE_INFO)

        End If

    Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j        As Integer
    Dim CharFile As String, Tmp As String
    Dim ObjInd   As Long, ObjCant As Long

    CharFile = CharPath & charName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos.", _
                FontTypeNames.FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, _
                        FontTypeNames.FONTTYPE_INFO)

            End If

        Next
    Else
        Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

