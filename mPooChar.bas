Attribute VB_Name = "mPooChar"
 Option Explicit
 
Sub EraseChar(ByVal charindex As Integer)

    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************
    
    On Error GoTo error_Err
    With charlist(charindex)
    
    If (charindex = 0) Then Exit Sub
    If (charindex > LastChar) Then Exit Sub
    

    If InMapBounds(.Pos.X, .Pos.Y) Then  '// Posicion valida
        MapData(.Pos.X, .Pos.Y).charindex = 0  '// Borramos el user
    End If
    
 
        'Update lastchar
        If charindex = LastChar Then
 
            Do Until charlist(LastChar).Heading > 0
               
                LastChar = LastChar - 1
 
                If LastChar = 0 Then
                                
                    NumChars = 0

                    Exit Sub

                End If
                       
            Loop
 
        End If
    
    Call ResetCharInfo(charindex)
    Call Char_Dialog_Remove(charindex)
 
        
    'Update NumChars
    NumChars = NumChars - 1
    End With
    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.EraseChar", Erl)
    Resume Next
    
End Sub

Public Sub Char_MapPosGet(ByVal charindex As Long, ByRef X As Byte, ByRef Y As Byte)
                                
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 13/12/2013
    '// By Miqueas150
    '
    '*****************************************************************
        
    'Make sure it's a legal char_index
    On Error GoTo error_Err
    
    With charlist(charindex)
                  
        'Get map pos
        X = .Pos.X
        Y = .Pos.Y
        
    End With
 
    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_MapPosGet", Erl)
    Resume Next
    
End Sub
 
Public Sub Char_MapPosSet(ByVal X As Byte, ByVal Y As Byte)

    'Sets the user postion
    On Error GoTo error_Err
    
    If (InMapBounds(X, Y)) Then  '// Posicion valida
        
        UserPos.X = X
        UserPos.Y = Y
                        
        'Set char
        MapData(UserPos.X, UserPos.Y).charindex = CurrentUser.UserCharIndex
        charlist(CurrentUser.UserCharIndex).Pos = UserPos
        
        Exit Sub
 
    End If

    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_MapPosSet", Erl)
    Resume Next
    
End Sub
Public Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
     On Error GoTo error_Err
    
    '*****************************************************************
    'Checks to see if a tile position has a char_index and return it
    '*****************************************************************
   
    If (InMapBounds(X, Y)) Then
        Char_MapPosExits = MapData(X, Y).charindex
    Else
        Char_MapPosExits = 0
    End If
  
    Exit Function

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_MapPosExits", Erl)
    Resume Next
    
End Function
Public Function Char_Techo() As Boolean

    '// Autor : Marcos Zeni
    '// Nueva forma de establecer si el usuario esta bajo un techo
    On Error GoTo error_Err
    
    Char_Techo = False
 
    With charlist(CurrentUser.UserCharIndex)
      
        If (InMapBounds(.Pos.X, .Pos.Y)) Then '// Posicion valida
                       
            If (MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.trigger_2) Then
                Char_Techo = True

            End If
                               
        End If
   
    End With
    Exit Function

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_Techo", Erl)
    Resume Next
    
End Function

Sub ResetCharInfo(ByVal charindex As Integer)
    On Error GoTo error_Err
    
    Call Char_Particle_Group_Remove_All(charindex)
    Call Delete_All_Auras(charindex)

    With charlist(charindex)
        .active = 0
        .priv = 0
        .FxIndex = 0
        .Invisible = False
        .Moving = 0
        .Muerto = False
        .LastStep = 0
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
        .Nombre = vbNullString
        .OffSetNombre = 0
        .Clan = vbNullString
        .OffSetClan = 0
        .color(0) = 0
        .color(1) = 0
        .color(2) = 0
        .color(3) = 0
        .particle_count = 0
        .Particula = 0
        .ParticulaTime = 0
        .Arma_Aura = 0
        .Body_Aura = 0
        .Escudo_Aura = 0
        .Head_Aura = 0
        .Anillo_Aura = 0
        .Otra_Aura = 0
        .EsGM = False
    End With

    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.ResetCharInfo", Erl)
    Resume Next
    
End Sub
 Public Sub Map_RemoveOldUser()
    On Error GoTo error_Err
    
      With MapData(UserPos.X, UserPos.Y)

            If (.charindex = CurrentUser.UserCharIndex) Then
                  .charindex = 0
            End If

      End With

    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Map_RemoveOldUser", Erl)
    Resume Next
    
End Sub

Public Sub Char_UserPos()

    '// Author Miqueas
    '// Actualizamo el lbl de la posicion del usuario
    On Error GoTo error_Err
    
    Dim X As Byte

    Dim Y As Byte
     
    If Char_Check(CurrentUser.UserCharIndex) Then
        
        bTecho = Char_Techo '// Pos : Techo :P
 
        Call Char_Refresh(CurrentUser.UserCharIndex)
        
        Call ActualizarMiniMapa
        
        Call Char_MapPosGet(CurrentUser.UserCharIndex, X, Y)
        
        If frmMain.UltPos = 0 Then
            If VerLugar = 1 Then
                frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & X & ", " & Y
            End If
        ElseIf VerLugar = 0 Then
            frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.UserMap & ", " & X & ", " & Y
        End If
        
    End If

    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_UserPos", Erl)
    Resume Next
    
End Sub
Public Sub Char_UserIndexSet(ByVal charindex As Integer)
    On Error GoTo error_Err
    
    CurrentUser.UserCharIndex = charindex
 
    With charlist(CurrentUser.UserCharIndex)
 
        'Nueva posicion para el usuario.
        UserPos = .Pos
         
        Exit Sub
 
    End With
         
    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_UserIndexSet", Erl)
    Resume Next
    
End Sub
Public Function Char_Check(ByVal charindex As Integer) As Boolean
'mkermas ok
    On Error GoTo error_Err
    'check CharIndex
    If charindex > 0 And charindex <= LastChar Then
     Char_Check = (charlist(charindex).Heading > 0)

    End If
  
    Exit Function

error_Err:
    Call RegistrarError(Err.number, Err.Description, "mPooChar.Char_Check", Erl)
    Resume Next
    
End Function

Public Function NickIgnorado(ByVal Nick As String) As Boolean

Dim i As Long

If Nick <> vbNullString Then
    Nick = UCase$(Nick)
    For i = 0 To frmOpciones.lstIgnore.ListCount
        If Nick = UCase$(frmOpciones.lstIgnore.list(i)) Then
            NickIgnorado = True
            Exit Function
        End If
    Next i
End If

End Function
