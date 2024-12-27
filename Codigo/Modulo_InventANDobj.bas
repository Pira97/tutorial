Attribute VB_Name = "InvNpc"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, _
                                Obj As Obj, _
                                Optional NotPirata As Boolean = True) As WorldPos
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, Obj, NotPirata, True

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)

    End If

    TirarItemAlPiso = NuevaPos

    Exit Function
ErrHandler:

End Function

'AyudandOh
Public Sub NPC_TIRAR_ITEMS(ByVal UserIndex As Integer, _
                           ByRef npc As npc)
 
    On Error Resume Next
 
    With npc
    
    
    If npc.Invent.NroItems > 0 Then
        Dim i As Byte
        Dim MiObj As Obj
        For i = 1 To MAX_INVENTORY_SLOTS
            If npc.Invent.Object(i).ObjIndex > 0 Then
                MiObj.Amount = npc.Invent.Object(i).Amount
                MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
                If Not npc.NPCtype = Dragon Then
                    If npc.Invent.Object(i).ProbTirar = 100 Then
                        Call TirarItemAlPiso(npc.Pos, MiObj)
                    ElseIf RandomNumber(1, 50) = npc.Invent.Object(i).ProbTirar Then
                        Call TirarItemAlPiso(npc.Pos, MiObj)
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_DROP, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    'Exit Sub
                    End If
            Else
                    If RandomNumber(1, 50) <= npc.Invent.Object(i).ProbTirar Then
                    Call TirarItemAlPiso(npc.Pos, MiObj)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_DROP, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    End If
                End If
            End If
        Next i
    End If

 
    End With
 
End Sub
Function QuedanItems(ByVal npcindex As Integer, ByVal ObjIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next
'
    Dim i As Integer

    If Npclist(npcindex).Invent.NroItems > 0 Then

        For i = 1 To MAX_INVENTORY_SLOTS

            If Npclist(npcindex).Invent.Object(i).ObjIndex = ObjIndex Then
                QuedanItems = True
                Exit Function

            End If

        Next

    End If

    QuedanItems = False

End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal npcindex As Integer, ByVal ObjIndex As Integer) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/09/08
    'Last Modification By: Marco Vanotti (Marco)
    ' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
    '***************************************************
    On Error Resume Next

    'Devuelve la cantidad original del obj de un npc

    Dim ln As String, npcfile As String
    Dim i  As Integer
    
    npcfile = DatPath & "NPCs.dat"
     
    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(npcindex).Numero, "Obj" & i)

        If ObjIndex = val(ReadField(1, ln, 45)) Then
            EncontrarCant = val(ReadField(2, ln, 45))
            Exit Function

        End If

    Next
                       
    EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim i As Integer
    
    With Npclist(npcindex)
        .Invent.NroItems = 0
        
        For i = 1 To MAX_INVENTORY_SLOTS
            .Invent.Object(i).ObjIndex = 0
            .Invent.Object(i).Amount = 0
        Next i
        
        .InvReSpawn = 0

    End With

End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal npcindex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 23/11/2009
    'Last Modification By: Marco Vanotti (Marco)
    ' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim ObjIndex As Integer
    Dim iCant    As Integer
    
    With Npclist(npcindex)
        ObjIndex = .Invent.Object(slot).ObjIndex
    
        'Quita un Obj
        If ObjData(.Invent.Object(slot).ObjIndex).Crucial = 0 Then
            .Invent.Object(slot).Amount = .Invent.Object(slot).Amount - Cantidad
            
            If .Invent.Object(slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(slot).ObjIndex = 0
                .Invent.Object(slot).Amount = 0

                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                    Call CargarInvent(npcindex) 'Reponemos el inventario

                End If

            End If

        Else
            .Invent.Object(slot).Amount = .Invent.Object(slot).Amount - Cantidad
            
            If .Invent.Object(slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(slot).ObjIndex = 0
                .Invent.Object(slot).Amount = 0
                
                If Not QuedanItems(npcindex, ObjIndex) Then
                    'Check if the item is in the npc's dat.
                    iCant = EncontrarCant(npcindex, ObjIndex)

                    If iCant Then
                        .Invent.Object(slot).ObjIndex = ObjIndex
                        .Invent.Object(slot).Amount = iCant
                        .Invent.NroItems = .Invent.NroItems + 1

                    End If

                End If
                
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                    Call CargarInvent(npcindex) 'Reponemos el inventario

                End If

            End If

        End If

    End With

End Sub

Sub CargarInvent(ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Vuelve a cargar el inventario del npc NpcIndex
    Dim loopc   As Integer
    Dim ln      As String
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    
    With Npclist(npcindex)
        .Invent.NroItems = val(GetVar(npcfile, "NPC" & .Numero, "NROITEMS"))
        
        For loopc = 1 To .Invent.NroItems
            ln = GetVar(npcfile, "NPC" & .Numero, "Obj" & loopc)
            .Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
            
        Next loopc

    End With

End Sub
