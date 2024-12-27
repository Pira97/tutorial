Attribute VB_Name = "modCentinela"
'*****************************************************************
'modCentinela.bas - ImperiumAO - v1.2
'
'Funci�nes de control para usuarios que se encuentran trabajando
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
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

'*****************************************************************
'Augusto Rando(barrin@imperiumao.com.ar)
'   ImperiumAO 1.2
'   - First Relase
'
'Juan Mart�n Sotuyo Dodero (juansotuyo@gmail.com)
'   Alkon AO 0.11.5
'   - Small improvements and added logs to detect possible cheaters
'
'Juan Mart�n Sotuyo Dodero (juansotuyo@gmail.com)
'   Alkon AO 0.12.0
'   - Added several messages to spam users until they reply
'*****************************************************************

Option Explicit

Private Const NPC_CENTINELA_TIERRA As Integer = 622  '�ndice del NPC en el .dat
Private Const NPC_CENTINELA_AGUA   As Integer = 623    '�dem anterior, pero en mapas de agua

Public CentinelaNPCIndex           As Integer                '�ndice del NPC en el servidor

Private Const TIEMPO_INICIAL       As Byte = 2 'Tiempo inicial en minutos. No reducir sin antes revisar el timer que maneja estos datos.

Private Type tCentinela

    RevisandoUserIndex As Integer   '�Qu� �ndice revisamos?
    TiempoRestante As Integer       '�Cu�ntos minutos le quedan al usuario?
    clave As Integer                'Clave que debe escribir
    spawnTime As Long

End Type

Public centinelaActivado As Boolean

Public Centinela         As tCentinela

Public Sub CallUserAttention()

    '############################################################
    'Makes noise and FX to call the user's attention.
    '############################################################
    If (GetTickCount() And &H7FFFFFFF) - Centinela.spawnTime >= 5000 Then
        If Centinela.RevisandoUserIndex <> 0 And centinelaActivado Then
            If Not UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK Then
                Call WritePlayWave(Centinela.RevisandoUserIndex, SND_WARP, Npclist(CentinelaNPCIndex).Pos.X, Npclist( _
                        CentinelaNPCIndex).Pos.Y)
                Call WriteCreateFX(Centinela.RevisandoUserIndex, Npclist(CentinelaNPCIndex).Char.CharIndex, _
                        FXIDs.FXWARP, 0)
                
                'Resend the key
                Call CentinelaSendClave(Centinela.RevisandoUserIndex)
                
                Call FlushBuffer(Centinela.RevisandoUserIndex)

            End If

        End If

    End If

End Sub

Private Sub GoToNextWorkingChar()
    '############################################################
    'Va al siguiente usuario que se encuentre trabajando
    '############################################################
    Dim loopc As Long
    
    For loopc = 1 To LastUser

        If UserList(loopc).flags.UserLogged And UserList(loopc).Counters.Trabajando > 0 And (UserList( _
                loopc).flags.Privilegios And PlayerType.User) Then
            If Not UserList(loopc).flags.CentinelaOK Then
                'Inicializamos
                Centinela.RevisandoUserIndex = loopc
                Centinela.TiempoRestante = TIEMPO_INICIAL
                Centinela.clave = RandomNumber(1, 32000)
                Centinela.spawnTime = GetTickCount() And &H7FFFFFFF
                
                'Ponemos al centinela en posici�n
                Call WarpCentinela(loopc)
                
                If CentinelaNPCIndex Then
                    'Mandamos el mensaje (el centinela habla y aparece en consola para que no haya dudas)
                    Call WriteChatOverHead(loopc, "Saludos " & UserList(loopc).Name & _
                            ", soy el Centinela de estas tierras. Me gustar�a que escribas /CENTINELA " & _
                            Centinela.clave & " en no m�s de dos minutos.", CStr(Npclist( _
                            CentinelaNPCIndex).Char.CharIndex))
                 '   Call WriteConsoleMsg( LoopC, "El centinela intenta llamar tu atenci�n. �Resp�ndele r�pido!", _
                            FontTypeNames.FONTTYPE_CENTINELA)
                    Call FlushBuffer(loopc)

                End If

                Exit Sub

            End If

        End If

    Next loopc
    
    'No hay chars trabajando, eliminamos el NPC si todav�a estaba en alg�n lado y esperamos otro minuto
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0

    End If
    
    'No estamos revisando a nadie
    Centinela.RevisandoUserIndex = 0

End Sub
Private Sub CentinelaFinalCheck()

    '############################################################
    'Al finalizar el tiempo, se retira y realiza la acci�n
    'pertinente dependiendo del caso
    '############################################################
    On Error GoTo Error_Handler

    Dim Name     As String
    Dim numPenas As Integer
    
    If Not UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK Then
        'Logueamos el evento
        Call LogCentinela("El centinela encarcel� a " & UserList(Centinela.RevisandoUserIndex).Name & _
                " por uso de macro inasistido.")
        
        'Ponemos la carcel
        Call Encarcelar(Centinela.RevisandoUserIndex, 15)
        Name = UserList(Centinela.RevisandoUserIndex).Name
        
        'Avisamos a los admins
        Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg("Servidor> El centinela ha encarcelado a " & Name, _
                FontTypeNames.FONTTYPE_SERVER))
        
 
        numPenas = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", numPenas + 1)
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & numPenas + 1, "Centinela : Macro inasistido." _
                & Date & " " & Time)
        
        'Evitamos loguear el logout
        Dim index As Integer
        index = Centinela.RevisandoUserIndex
        Centinela.RevisandoUserIndex = 0
        
        Call CloseSocket(index)

    End If
    
    Centinela.clave = 0
    Centinela.TiempoRestante = 0
    Centinela.RevisandoUserIndex = 0
    
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0

    End If

    Exit Sub

Error_Handler:
    Centinela.clave = 0
    Centinela.TiempoRestante = 0
    Centinela.RevisandoUserIndex = 0
    
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0

    End If
    
    Call LogError("Error en el checkeo del centinela: " & Err.description)

End Sub
    
Public Sub CentinelaCheckClave(ByVal UserIndex As Integer, ByVal clave As Integer)

    '############################################################
    'Corrobora la clave que le envia el usuario
    '############################################################
    If clave = Centinela.clave And UserIndex = Centinela.RevisandoUserIndex Then
        UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK = True
        Call WriteChatOverHead(UserIndex, "�Muchas gracias " & UserList(Centinela.RevisandoUserIndex).Name & _
                "! Espero no haber sido una molestia.", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
        Centinela.RevisandoUserIndex = 0
        Call FlushBuffer(UserIndex)
    Else
        Call CentinelaSendClave(UserIndex)
        
        'Logueamos el evento
        If UserIndex <> Centinela.RevisandoUserIndex Then
            Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondi� aunque no se le hablaba a �l.")
        Else
            Call LogCentinela("El usuario " & UserList(UserIndex).Name & " respondi� una clave incorrecta: " & clave _
                    & " - Se esperaba : " & Centinela.clave)

        End If

    End If

End Sub

Public Sub ResetCentinelaInfo()
    '############################################################
    'Cada determinada cantidad de tiempo, volvemos a revisar
    '############################################################
    Dim loopc As Long
    
    For loopc = 1 To LastUser

        If (LenB(UserList(loopc).Name) <> 0 And loopc <> Centinela.RevisandoUserIndex) Then
            UserList(loopc).flags.CentinelaOK = False

        End If

    Next loopc

End Sub

Public Sub CentinelaSendClave(ByVal UserIndex As Integer)

    '############################################################
    'Enviamos al usuario la clave v�a el personaje centinela
    '############################################################
    If CentinelaNPCIndex = 0 Then Exit Sub
    
    If UserIndex = Centinela.RevisandoUserIndex Then
        If Not UserList(UserIndex).flags.CentinelaOK Then
            Call WriteChatOverHead(UserIndex, "�La clave que te he dicho es /CENTINELA " & Centinela.clave & _
                    ", escr�belo r�pido!", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex))
        Else
            'Logueamos el evento
            Call LogCentinela("El usuario " & UserList(Centinela.RevisandoUserIndex).Name & _
                    " respondi� m�s de una vez la contrase�a correcta.")
            Call WriteChatOverHead(UserIndex, "Te agradezco, pero ya me has respondido. Me retirar� pronto.", CStr( _
                    Npclist(CentinelaNPCIndex).Char.CharIndex))

        End If

    Else
        Call WriteChatOverHead(UserIndex, "No es a ti a quien estoy hablando, �No ves?", CStr(Npclist( _
                CentinelaNPCIndex).Char.CharIndex))

    End If

End Sub

Public Sub PasarMinutoCentinela()

    '############################################################
    'Control del timer. Llamado cada un minuto.
    '############################################################
    If Not centinelaActivado Then Exit Sub
    
    If Centinela.RevisandoUserIndex = 0 Then
        Call GoToNextWorkingChar
    Else
        Centinela.TiempoRestante = Centinela.TiempoRestante - 1
        
        If Centinela.TiempoRestante = 0 Then
            Call CentinelaFinalCheck
            Call GoToNextWorkingChar
        Else

            'Recordamos al user que debe escribir
            If Matematicas.Distancia(Npclist(CentinelaNPCIndex).Pos, UserList(Centinela.RevisandoUserIndex).Pos) > 5 _
                    Then
                Call WarpCentinela(Centinela.RevisandoUserIndex)

            End If
            
            'El centinela habla y se manda a consola para que no quepan dudas
            Call WriteChatOverHead(Centinela.RevisandoUserIndex, "�" & UserList(Centinela.RevisandoUserIndex).Name & _
                    ", tienes un minuto m�s para responder! Debes escribir /CENTINELA " & Centinela.clave & ".", CStr( _
                    Npclist(CentinelaNPCIndex).Char.CharIndex))
            Call WriteConsoleMsg(Centinela.RevisandoUserIndex, "�" & UserList(Centinela.RevisandoUserIndex).Name & _
                    ", tienes un minuto m�s para responder!", FontTypeNames.FONTTYPE_CENTINELA)
            Call FlushBuffer(Centinela.RevisandoUserIndex)

        End If

    End If

End Sub

Private Sub WarpCentinela(ByVal UserIndex As Integer)

    '############################################################
    'Inciamos la revisi�n del usuario UserIndex
    '############################################################
    'Evitamos conflictos de �ndices
    If CentinelaNPCIndex Then
        Call QuitarNPC(CentinelaNPCIndex)
        CentinelaNPCIndex = 0

    End If
    
    If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
        CentinelaNPCIndex = SpawnNpc(NPC_CENTINELA_AGUA, UserList(UserIndex).Pos, True, False)
    Else
        CentinelaNPCIndex = SpawnNpc(NPC_CENTINELA_TIERRA, UserList(UserIndex).Pos, True, False)

    End If
    
    'Si no pudimos crear el NPC, seguimos esperando a poder hacerlo
    If CentinelaNPCIndex = 0 Then Centinela.RevisandoUserIndex = 0

End Sub

Public Sub CentinelaUserLogout()

    '############################################################
    'El usuario al que revisabamos se desconect�
    '############################################################
    If Centinela.RevisandoUserIndex Then
        'Logueamos el evento
        Call LogCentinela("El usuario " & UserList(Centinela.RevisandoUserIndex).Name & _
                " se desolgue� al pedirsele la contrase�a.")
        
        'Reseteamos y esperamos a otro PasarMinuto para ir al siguiente user
        Centinela.clave = 0
        Centinela.TiempoRestante = 0
        Centinela.RevisandoUserIndex = 0
        
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0

        End If

    End If

End Sub

Private Sub LogCentinela(ByVal texto As String)

    '*************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last modified: 03/15/2006
    'Loguea un evento del centinela
    '*************************************************
    On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Centinela.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile
    Exit Sub

ErrHandler:

End Sub
