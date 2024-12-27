Attribute VB_Name = "mod_BindKeys"
Public Enum eMacros
    aComando = 1
    aLanzar
    aEquipar
    aUsar
End Enum

Public Type tMacros
    mTipe As Byte
    Grh As Integer
    Nombre As String
    Slot As Byte
    OBJIndex As Integer
    SpellSlot As Byte
End Type
Public MacroIndex As Integer

Public MacroList(1 To 11) As tMacros
Public Sub LoadMacros(ByVal Nombre As String)
    Dim MacroPatch As String
    Dim i As Integer
    MacroPatch = App.Path & "\Init\Macros\" & Nombre & ".Mac"
    If FileExist(MacroPatch, vbNormal) Then
        For i = 1 To 11
            With MacroList(i)
                .Nombre = GetVar(MacroPatch, "Macro" & i, "Nombre")
                .Grh = val(GetVar(MacroPatch, "Macro" & i, "Grh"))
                .mTipe = val(GetVar(MacroPatch, "Macro" & i, "Tipo"))
                .Slot = val(GetVar(MacroPatch, "Macro" & i, "Slot"))
                .SpellSlot = val(GetVar(MacroPatch, "Macro" & i, "SlotSpell"))
                .OBJIndex = val(GetVar(MacroPatch, "Macro" & i, "ObjIndex"))
            End With
        Next i
    Else
        For i = 1 To 11
            With MacroList(i)
                .Nombre = vbNullString
                .Grh = 0
                .mTipe = 0
                .Slot = 0
                .SpellSlot = 0
                .OBJIndex = 0
            End With
         Next i
         Call SaveMacros(Nombre)
    End If
End Sub
Public Sub SaveMacros(ByVal Nombre As String)
    Dim MacroPatch As String
    Dim i As Integer
    MacroPatch = App.Path & "\init\Macros\" & Nombre & ".Mac"

        For i = 1 To 11
            With MacroList(i)
                Call WriteVar(MacroPatch, "Macro" & i, "Nombre", .Nombre)
                Call WriteVar(MacroPatch, "Macro" & i, "Grh", .Grh)
                Call WriteVar(MacroPatch, "Macro" & i, "Tipo", .mTipe)
                Call WriteVar(MacroPatch, "Macro" & i, "Slot", .Slot)
                Call WriteVar(MacroPatch, "Macro" & i, "SlotSpell", .SpellSlot)
                Call WriteVar(MacroPatch, "Macro" & i, "ObjIndex", .OBJIndex)
            End With
        Next i
End Sub
Public Function CheckMacrosSpells(ByVal SlotSpells As Byte, ByVal NameSpell As String, ByVal MacroIndex As Byte) As Byte
    Dim i As Integer
    
    If SlotSpells < 0 Or SlotSpells > MAXHECHI - 1 Or _
       NameSpell = "" Then Exit Function
       
    If frmMain.hlst.list(SlotSpells) = NameSpell Then
        CheckMacrosSpells = SlotSpells
        Exit Function
    Else
        For i = 0 To 34
            If frmMain.hlst.list(i) = NameSpell Then
                Exit For
            End If
        Next i

        CheckMacrosSpells = i
        MacroList(MacroIndex).SpellSlot = i
        Call SaveMacros(UserName)
        Exit Function
    End If
    
    'ERROR!!
    CheckMacrosSpells = -1
    MacroList(MacroIndex).mTipe = 0

End Function

Public Function CheckMacrosUsarItem(ByVal Slot As Byte, ByVal OBJIndex As Integer, ByVal MacroIndex As Byte) As Byte
    Dim i As Byte

    If Slot = 0 Or Slot > MAX_INVENTORY_SLOTS Then Exit Function

    If Inventario.OBJIndex(Slot) = OBJIndex Then
        CheckMacrosUsarItem = Slot
        Exit Function
    Else
        For i = 1 To MAX_INVENTORY_SLOTS - 1
            If Inventario.OBJIndex(i) = OBJIndex Then
                Exit For
            End If
        Next i

        If Inventario.OBJIndex(i) = OBJIndex Then
            CheckMacrosUsarItem = i
            MacroList(MacroIndex).Slot = i
            Call SaveMacros(UserName)
            Exit Function
        Else
            CheckMacrosUsarItem = 0
        End If


        Exit Function
    End If
End Function

Public Sub UsarMacro(ByVal Index As Byte)
    Dim Slot As Byte
    
    Select Case MacroList(Index).mTipe

    Case eMacros.aLanzar
        If CurrentUser.Muerto Then
 
            Exit Sub
        End If
        If Slot = 34 Then
           Slot = Slot - 1
        End If
        Slot = CheckMacrosSpells(MacroList(Index).SpellSlot, MacroList(Index).Nombre, Index)
        Call WriteCastSpell(Slot + 1)
        
        UsaMacro = True

    Case eMacros.aUsar
        If CurrentUser.Muerto Then
   
            Exit Sub
        End If
        Slot = CheckMacrosUsarItem(MacroList(Index).Slot, MacroList(Index).OBJIndex, Index)
        If Slot = 0 Then Exit Sub
        If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
        If MainTimer.Check(TimersIndex.UseItemWithU) Then _
           Call WriteUseItem(Slot)
    
    Case eMacros.aEquipar
        If CurrentUser.Muerto Then
 
        Exit Sub
        End If
        Slot = CheckMacrosUsarItem(MacroList(Index).Slot, MacroList(Index).OBJIndex, Index)
        If Slot = 0 Then Exit Sub
        
        If Inventario.ObjType(MacroList(Index).Slot) = eObjType.otpociones And MainTimer.Check(TimersIndex.UseItemWithU) Then
        Call WriteUseItem(Slot)
        Exit Sub
        End If
        
        If Comerciando Then Exit Sub
 
        Call frmMain.EquiparObjeto(Slot)
        
    Case eMacros.aComando
    'If LenB(MacroList(Index).Nombre) > 0 Then _
  '  Call Clienttcp.ParseUserCommand("/" & MacroList(Index).Nombre)
    Case Else
        MacroIndex = Index
        FrmBindKey.lblTecla = Locale_GUI_Frase(205) & ": F" & Index
        FrmBindKey.Show vbModeless, frmMain
    End Select

End Sub
