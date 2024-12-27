Attribute VB_Name = "mod_Correos"

Public Sub EnviarCorreo(ByVal UserIndex As Integer, ByVal Destinatario As String, ByVal Mensaje As String, ByVal ObjIndex As Integer, ByVal AmountIndex As Integer, Optional ByVal EnvioDirecto As Byte = 0)
'---------------------------------------------
'Last Modification: Shermie80
'18/08/15
'---------------------------------------------

Dim slot        As Byte
Dim tObject     As Obj
Dim Obj         As ObjData
Dim pjexiste    As Boolean
Dim tUser As Integer
Dim CantidadCorreos As Byte
tUser = NameIndex(Destinatario)
pjexiste = FileExist(CharPath & Destinatario & ".chr")

If ObjIndex <> 0 And AmountIndex <> 0 Then
tObject.Amount = AmountIndex
tObject.ObjIndex = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
Obj = ObjData(tObject.ObjIndex)
End If

If EnvioDirecto = 0 Then
'Si no existe el personaje.
If Not pjexiste Then
  Call WriteConsoleMsg(UserIndex, "No existe el personaje [" & Destinatario & "]", FontTypeNames.FONTTYPE_INFO)
  Exit Sub
  End If
    
'Descuento el oro.
If UserList(UserIndex).Stats.GLD < 1750 Then
 Call WriteLocaleMsg(UserIndex, 26) 'Oro insuficiente.|12|1
 Exit Sub
End If

'Saco el item
If TieneObjetos(tObject.ObjIndex, tObject.Amount, UserIndex) Then
  Call QuitarUserInvItem(UserIndex, ObjIndex, tObject.Amount)
  Call InvUsuario.UpdateUserInv(False, UserIndex, ObjIndex)
Else
  Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
 Exit Sub
End If


 UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1750
 Call WriteUpdateGold(UserIndex)

End If
       
        If tUser > 0 Then
         slot = UserList(tUser).flags.CantidadCorreos + 1

         If EnvioDirecto = 1 Then
         UserList(tUser).Correos(slot).Emisor = "LinkAO Staff"
         Else
         UserList(tUser).Correos(slot).Emisor = UserList(UserIndex).Name
         End If
         UserList(tUser).Correos(slot).Leida = 0
         UserList(tUser).Correos(slot).Carta = Mensaje
         If ObjIndex <> 0 And AmountIndex <> 0 Then
         UserList(tUser).Correos(slot).ObjetoIndex = tObject.ObjIndex
         UserList(tUser).Correos(slot).ObjetoCantidad = AmountIndex
         End If
         UserList(tUser).flags.RecibioCorreo = 1
         UserList(tUser).flags.CantidadCorreos = (UserList(tUser).flags.CantidadCorreos + 1)
         Call WriteMensajeSigno(tUser, 1)
         
          
         If EnvioDirecto = 1 Then
         Call WriteConsoleMsg(tUser, "Has recibido una respuesta de [LinkAO Staff], ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFO)
         Else
         Call WriteConsoleMsg(tUser, "Has recibido un nuevo mensaje de " & "[" & UserList(UserIndex).Name & "]" & ", ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFO)
         End If
        
        Else
        
        
        slot = GetVar(CharPath & UCase$(Destinatario) & ".chr", "FLAGS", "CORREOS") + 1
        If EnvioDirecto = 1 Then
        Call WriteVar(App.Path & "\charfile\" & Destinatario & ".chr", "CORREO", "EMISOR" & slot, "LinkAO Staff")
        Else
        Call WriteVar(App.Path & "\charfile\" & Destinatario & ".chr", "CORREO", "EMISOR" & slot, UserList(UserIndex).Name)
        End If
        Call WriteVar(App.Path & "\charfile\" & Destinatario & ".chr", "CORREO", "LEIDA" & slot, 0)
        Call WriteVar(App.Path & "\charfile\" & Destinatario & ".chr", "CORREO", "CARTA" & slot, Mensaje)
                 If ObjIndex <> 0 And AmountIndex <> 0 Then
        Call WriteVar(App.Path & "\charfile\" & Destinatario & ".chr", "CORREO", "OBJETO" & slot, tObject.ObjIndex & "-" & AmountIndex & "-" & Obj.GrhIndex & "-" & Obj.Name)
                 End If
        Call WriteVar(CharPath & UCase$(Destinatario) & ".chr", "FLAGS", "RecibioCorreo", "1")
        CantidadCorreos = GetVar(CharPath & UCase$(Destinatario) & ".chr", "FLAGS", "CORREOS")
        Call WriteVar(CharPath & UCase$(Destinatario) & ".chr", "FLAGS", "CORREOS", CantidadCorreos + 1)
 End If
        Call WriteConsoleMsg(UserIndex, "El mensaje fue enviado correctamente.", FontTypeNames.FONTTYPE_INFO)
 
        
End Sub
 
 
Public Function CantSendCorreo(ByVal UserIndex As Integer, _
                               ByVal Destinatario As String) As Boolean
    
        Dim tUser As Integer
        tUser = NameIndex(Destinatario)
        CantSendCorreo = False
      
      If tUser > 0 Then
        If (getFreeSlotCorreo(Destinatario, 1) = -1) Then
            ' @@ El target no tiene suficiente espacio en la lista de correos :P
            Call WriteConsoleMsg(UserIndex, Destinatario & " no tiene más espacio en su casilla de mensajes.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
        End If
      Else
            If (getFreeSlotCorreo(Destinatario, 0) = -1) Then
            ' @@ El target no tiene suficiente espacio en la lista de correos :P
            Call WriteConsoleMsg(UserIndex, Destinatario & " no tiene más espacio en su casilla de mensajes.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
            End If
      End If

      CantSendCorreo = True
End Function
 
Public Function getFreeSlotCorreo(ByVal UserName As String, Optional ByVal Estado As Byte) As Integer
      ' @@ damos un slot del correo para generar un nuevo correo
 
      Dim index As Long
      Dim CantidadCorreos As String
      Dim tUser As Integer
   
   If Estado = 0 Then
        CantidadCorreos = GetVar(CharPath & UserName & ".chr", "FLAGS", "CORREOS")
            If CantidadCorreos < Max_Correos Then
                getFreeSlotCorreo = index + 1
                Exit Function
            End If
    Else
         tUser = NameIndex(UserName)
        CantidadCorreos = UserList(tUser).flags.CantidadCorreos
       
           If CantidadCorreos < Max_Correos Then
                getFreeSlotCorreo = index + 1
                Exit Function
            End If
    End If
      getFreeSlotCorreo = -1
End Function
 
Sub ResetCorreos(ByVal UserIndex As Integer, ByVal index As Byte)
      '//Shak - Sistema de correos
      'Last Modification: Shermie80
 
      With UserList(UserIndex)
        Dim CantidadCorreos As Byte
        CantidadCorreos = .flags.CantidadCorreos
        If index = 0 Then Exit Sub 'podria pasar xD //mermas
        
        
        
       If index = CantidadCorreos Then
            With .Correos(index)
                .Carta = vbNullString
                .Emisor = vbNullString 'Nombre de quien manda la carta.
                .Leida = 0
                .ObjetoCantidad = 0
                .ObjetoIndex = 0
            End With
            UserList(UserIndex).flags.CantidadCorreos = CantidadCorreos - 1
            Else
            UserList(UserIndex).flags.CantidadCorreos = CantidadCorreos - 1
            For slot = index To (CantidadCorreos)
                    If slot = CantidadCorreos Then Exit For
                    UserList(UserIndex).Correos(slot).Carta = UserList(UserIndex).Correos(slot + 1).Carta
                    UserList(UserIndex).Correos(slot + 1).Carta = vbNullString
                                        
                                        
                    UserList(UserIndex).Correos(slot).Emisor = UserList(UserIndex).Correos(slot + 1).Emisor
                    UserList(UserIndex).Correos(slot + 1).Emisor = vbNullString
                    
                    UserList(UserIndex).Correos(slot).Leida = UserList(UserIndex).Correos(slot + 1).Leida
                    UserList(UserIndex).Correos(slot + 1).Leida = 0
                    
                    UserList(UserIndex).Correos(slot).ObjetoCantidad = UserList(UserIndex).Correos(slot + 1).ObjetoCantidad
                    UserList(UserIndex).Correos(slot + 1).ObjetoCantidad = 0

                    UserList(UserIndex).Correos(slot).ObjetoIndex = UserList(UserIndex).Correos(slot + 1).ObjetoIndex
                    UserList(UserIndex).Correos(slot + 1).ObjetoIndex = 0
                     
                    Next slot
        
         End If
      End With
       
                     
 
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : RetirarItemCorreo
' Author    : Shermie80
' Date      : 18/08/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub RetirarItemCorreo(ByVal UserIndex As Integer, ByVal index As Byte)

Dim Leer As New clsIniReader
Dim Item As Integer
Dim cant As Integer
Dim slot As Byte
Dim tObject As Obj
    Leer.Initialize CharPath & UserList(UserIndex).Name & ".chr"
    Item = CInt(ReadField(1, Leer.GetValue("CORREO", "Objeto" & index), 45))
    cant = CInt(ReadField(2, Leer.GetValue("CORREO", "Objeto" & index), 45))
     Item = UserList(UserIndex).Correos(index).ObjetoIndex
    cant = UserList(UserIndex).Correos(index).ObjetoCantidad
     tObject.Amount = cant
    tObject.ObjIndex = Item

    If Not MeterItemEnInventario(UserIndex, tObject) Then
      Call TirarItemAlPiso(UserList(UserIndex).Pos, tObject)
    End If
   
    With UserList(UserIndex)

        If index <> 0 Then
        
            With .Correos(index)
                .Leida = 1
                .ObjetoCantidad = 0
                .ObjetoIndex = 0
 
            End With
            
       End If
       
   End With
   
   Call UpdateUserInv(True, UserIndex, 0) 'Actualizo en inventario
End Sub
