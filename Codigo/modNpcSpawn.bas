Attribute VB_Name = "modNpcSpawn"
Option Explicit

'Cantidad maxima de npcs
Private Const MAX_INVOCACIONES As Integer = 620

'Tiempo extra si el npc no se logro invocar(En minutos)
Private Const TIEMPO_EXTRA As Integer = 2

Private Type t_Npc
    Num As Integer    'Numero del npc
    Time As Long      'Tiempo del npc
    Pos As WorldPos   'Pos donde renace el npc
End Type

Private Cuantos As Integer

'Array con la info de los npcs
Private Npcs(1 To MAX_INVOCACIONES) As t_Npc
Private Function Slot_Libre() As Integer


    Dim i As Long

    For i = 1 To MAX_INVOCACIONES
        With Npcs(i)
            If .Num = 0 And .Time = 0 And .Pos.map = 0 Then
                Slot_Libre = i
                Exit Function
            End If
        End With
    Next i
End Function
Private Sub Reset_Slot(ByVal Slot As Integer)
'***************************************************
'Autor:Bateman
'
'***************************************************

    Npcs(Slot).Num = 0
    Npcs(Slot).Time = 0
    Npcs(Slot).Pos.map = 0
    Npcs(Slot).Pos.x = 0
    Npcs(Slot).Pos.y = 0

End Sub

Public Sub Agregar_Npc(ByVal Num As Integer, ByVal Tiempo As Long, Pos As WorldPos)
'***************************************************
'Autor:Bateman
'
'***************************************************
    Dim i As Integer
    Dim MSG As String
    'Esto esta de mas pero lo pongo por las dudas...
    If InMapBounds(Pos.map, Pos.x, Pos.y) = False Or Num <= 0 Then
        ' Call LogInvocaciones("No se logro invocar al npc " & Num & " en el mapa " & Pos.map)
        Exit Sub
    End If

 
    'Buscamos un slot libre
    i = Slot_Libre()

    'Agregamos el npc
    Npcs(i).Num = Num
    Npcs(i).Time = Tiempo
    Npcs(i).Pos.map = Pos.map
    Npcs(i).Pos.x = Pos.x
    Npcs(i).Pos.y = Pos.y

    Cuantos = Cuantos + 1


    'MSG = "Se agrego el npc " & Num & " Tiempo " & Tiempo & " Pos " & Pos.map & " " & Pos.X & " " & Pos.Y
   ' Call LogInvocaciones(MSG)
End Sub
Public Sub Comprobar_Tiempo_Npc()
'***************************************************
'Autor:Bateman
'
'***************************************************
    Dim i As Long

    For i = 1 To MAX_INVOCACIONES
        If Npcs(i).Num > 0 And Npcs(i).Time > 0 And Npcs(i).Pos.map > 0 Then
            Npcs(i).Time = Npcs(i).Time - 1
            If Npcs(i).Time < 1 Then
                Call Invocar_Npc(i)
            End If
        End If
    Next i
End Sub
Private Sub Invocar_Npc(ByVal Slot As Integer)
'***************************************************
'Autor:Bateman
'
'***************************************************
    Dim NpcIndex As Integer

    NpcIndex = SpawnNpc(Npcs(Slot).Num, Npcs(Slot).Pos, False)

    If NpcIndex <> 0 Then
        With Npclist(NpcIndex)
            .InvocacionPos.map = Npcs(Slot).Pos.map
            .InvocacionPos.x = Npcs(Slot).Pos.x
            .InvocacionPos.y = Npcs(Slot).Pos.y
            Call Reset_Slot(Slot)
            Cuantos = Cuantos - 1
            ' Call LogInvocaciones("Se invoco al npc " & .Name)
        End With
    Else  'No se logro invocarlo!
        Npcs(Slot).Time = TIEMPO_EXTRA
    End If


End Sub
