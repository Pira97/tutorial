Attribute VB_Name = "mod_Auras"
Option Explicit

Public Sub SetAura(ByVal UserIndex As Integer, _
                   ByVal Aura As Byte, _
                   ByVal slot As Byte, _
                   Optional ByVal Refresh As Boolean = False)

    If slot <= 0 Or slot >= 7 Then Exit Sub
    
    UserList(UserIndex).Char.Aura(slot) = Aura
    
     If Refresh Then
        SendSpecificAura UserIndex, slot
    End If

    
End Sub

Public Sub ResetAuras(ByVal UserIndex As Integer)

    Dim i As Long

    For i = 1 To 6
        UserList(UserIndex).Char.Aura(i) = 0
    Next i

    
End Sub

Public Sub SendSpecificAura(ByVal UserIndex As Integer, ByVal slot As Byte)

    If UserList(UserIndex).Char.Aura(slot) <> 0 Then
        Call modSendData.SendToMap(UserList(UserIndex).Pos.map, PrepareMessageSendAura(UserIndex, slot))

    End If

End Sub

Public Sub SendAuras(ByVal UserIndex As Integer)
    Dim i As Long

    For i = 1 To 6

        If UserList(UserIndex).Char.Aura(i) <> 0 Then
            Call SendToMap(UserList(UserIndex).Pos.map, PrepareMessageSendAura(UserIndex, CByte(i)))

        End If

    Next i

End Sub

Public Sub KickAuras(ByVal UserIndex As Integer)
    Dim i As Long

    For i = 1 To 6
        UserList(UserIndex).Char.Aura(i) = 0
        Call SendToMap(UserList(UserIndex).Pos.map, PrepareMessageSendAura(UserIndex, CByte(i)))
    Next i
    
    
End Sub

Public Function FindSlotFreeAura(ByVal UserIndex As Integer) As Byte

    Dim i As Long

    For i = 1 To 6

        If UserList(UserIndex).Char.Aura(i) = 0 Or UserList(UserIndex).Char.Aura(i) = 1 Then
            FindSlotFreeAura = CByte(i)
            Exit Function

        End If

    Next i

End Function

Public Function TieneEstaAura(ByVal UserIndex As Integer, ByVal Aura As Byte) As Byte
    Dim i As Long

    For i = 1 To 6

        If UserList(UserIndex).Char.Aura(i) = Aura Then
            TieneEstaAura = CByte(i)
            Exit Function

        End If

    Next i

    TieneEstaAura = 0

End Function



