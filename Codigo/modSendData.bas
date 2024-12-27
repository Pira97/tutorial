Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget

    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToADMINS
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToClanArea
    ToDeadArea
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, _
                    ByVal sndIndex As Integer, _
                    ByVal sndData As String)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
    'Last Modify Date: 01/08/2007
    'Last modified by: (liquid)
    '**************************************************************
    On Error Resume Next

    Dim loopc As Long
    Dim Map   As Integer
    
    Select Case sndRoute

        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToADMINS

            For loopc = 1 To LastUser

                If UserList(loopc).ConnID <> -1 Then
                    If UserList(loopc).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or _
                            PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call EnviarDatosASlot(loopc, sndData)

                    End If

                End If

            Next loopc

            Exit Sub
        
        Case SendTarget.ToAll

            For loopc = 1 To LastUser

                If UserList(loopc).ConnID <> -1 Then
                    If UserList(loopc).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopc, sndData)

                    End If

                End If

            Next loopc

            Exit Sub
        
        Case SendTarget.ToAllButIndex

            For loopc = 1 To LastUser

                If (UserList(loopc).ConnID <> -1) And (loopc <> sndIndex) Then
                    If UserList(loopc).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopc, sndData)

                    End If

                End If

            Next loopc

            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While loopc > 0

                If (UserList(loopc).ConnID <> -1) Then
                    Call EnviarDatosASlot(loopc, sndData)

                End If

                loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub

            
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYclan
            loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While loopc > 0

                If (UserList(loopc).ConnID <> -1) Then
                    Call EnviarDatosASlot(loopc, sndData)

                End If

                loopc = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            loopc = modGuilds.Iterador_ProximoGM(sndIndex)

            While loopc > 0

                If (UserList(loopc).ConnID <> -1) Then
                    Call EnviarDatosASlot(loopc, sndData)

                End If

                loopc = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
 

        Case SendTarget.ToHigherAdmins

            For loopc = 1 To LastUser

                If UserList(loopc).ConnID <> -1 Then
                    If UserList(loopc).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call EnviarDatosASlot(loopc, sndData)

                    End If

                End If

            Next loopc

            Exit Sub
            
        Case SendTarget.ToGMsAreaButRmsOrCounselors
            Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)
            Exit Sub

        Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
            Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
            Exit Sub
            
            

    End Select

End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim TempInt   As Integer
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If tempIndex <> UserIndex Then
                    If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

        End If

    Next loopc

End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

                'Dead and admins read
                If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList( _
                        tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios _
                        Or PlayerType.Consejero)) <> 0) Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).GuildIndex = UserList( _
                        UserIndex).GuildIndex Or ((UserList(tempIndex).flags.Privilegios And PlayerType.Dios) And ( _
                        UserList(tempIndex).flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, _
                                          ByVal sdData As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or _
                            PlayerType.Admin) Then Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Private Sub SendToNpcArea(ByVal npcindex As Long, ByVal sdData As String)
    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim TempInt   As Integer
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = Npclist(npcindex).Pos.Map
    AreaX = Npclist(npcindex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(npcindex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, _
                           ByVal AreaX As Integer, _
                           ByVal AreaY As Integer, _
                           ByVal sdData As String)
    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim loopc     As Long
    Dim TempInt   As Integer
    Dim tempIndex As Integer
    
    AreaX = 2 ^ (AreaX \ 9)
    AreaY = 2 ^ (AreaY \ 9)
    
    If Not MapaValido(Map) Then Exit Sub

    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next loopc

End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    If Not MapaValido(Map) Then Exit Sub

    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)

        End If

    Next loopc

End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim loopc     As Long
    Dim Map       As Integer
    Dim tempIndex As Integer
    
    Map = UserList(UserIndex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub

    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)

        End If

    Next loopc

End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, _
                                            ByVal sdData As String)
    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 12/02/2010
    '12/02/2010: ZaMa - Restrinjo solo a dioses, admins y gms.
    '15/02/2010: ZaMa - Cambio el nombre de la funcion (viejo: ToGmsArea, nuevo: ToGmsAreaButRMsOrCounselors)
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        With UserList(tempIndex)

            If .AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If .AreasInfo.AreaReciveY And AreaY Then
                    If .ConnIDValida Then

                        ' Exclusivo para dioses, admins y gms
                        If (.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero And Not _
                                PlayerType.RoleMaster) = .flags.Privilegios Then
                            Call EnviarDatosASlot(tempIndex, sdData)

                        End If

                    End If

                End If

            End If

        End With

    Next loopc

End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)
    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 10/17/2009
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And PlayerType.User Then
                        Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

        End If

    Next loopc

End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, _
                                                     ByVal sdData As String)
    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 10/17/2009
    '
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    
    Dim Map       As Integer
    Dim AreaX     As Integer
    Dim AreaY     As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or _
                            PlayerType.RoleMaster) Then
                        Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

        End If

    Next loopc

End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Alerta a los faccionarios, dandoles una orientacion
    '**************************************************************
    Dim loopc     As Long
    Dim tempIndex As Integer
    Dim Map       As Integer
    Dim Font      As FontTypeNames
    
    If esCaos(UserIndex) Then
        Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
    Else
        Font = FontTypeNames.FONTTYPE_CONSEJO

    End If
    
    Map = UserList(UserIndex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub

    For loopc = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(loopc)
        
        If UserList(tempIndex).ConnIDValida Then
            If tempIndex <> UserIndex Then

                ' Solo se envia a los de la misma faccion
                If SameFaccion(UserIndex, tempIndex) Then
                    Call EnviarDatosASlot(tempIndex, PrepareMessageConsoleMsg( _
                            "Escuchas el llamado de un compañero que proviene del " & GetDireccion(UserIndex, _
                            tempIndex), Font))

                End If

            End If

        End If

    Next loopc

End Sub


