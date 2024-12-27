Attribute VB_Name = "Admin"
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

Public Type tAPuestas

    Ganancias As Long
    Perdidas As Long
    Jugadas As Long

End Type

Public Apuestas                     As tAPuestas

Public MinutosGuardarUsuarios As Long

Public tInicioServer                As Long

'INTERVALOS
'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed                 As Integer
Public IntervaloHambre              As Integer
Public IntervaloVeneno              As Integer
Public IntervaloIncinerado          As Integer
Public IntervaloParalizado          As Integer

Public Const IntervaloParalizadoReducido As Integer = 37

Public IntervaloInvisible           As Integer
Public IntervaloFrio                As Integer
Public IntervaloWavFx               As Integer
Public IntervaloLanzaHechizo        As Integer

Public IntervaloNPCAI               As Integer
Public IntervaloInvocacion          As Integer
Public IntervaloOculto              As Integer '[Nacho]
Public IntervaloUserPuedeAtacar     As Long
Public IntervaloGolpeUsar           As Long
Public IntervaloMagiaGolpe          As Long
Public IntervaloGolpeMagia          As Long
Public IntervaloUserPuedeCastear    As Long
Public IntervaloUserPuedeTrabajar   As Long
Public IntervaloParaConexion        As Long
Public IntervaloCerrarConexion      As Long '[Gonzalo]
Public IntervaloUserPuedeUsar       As Long
Public IntervaloFlechasCazadores    As Long
Public IntervaloPuedeSerAtacado     As Long
Public IntervaloAtacable            As Long
Public IntervaloOwnedNpc            As Long
Public IntervaloMensajeAutomatico1           As Long
Public IntervaloMensajeAutomatico2           As Long
'BALANCE

Public PorcentajeRecuperoMana       As Integer

Public Puerto                       As Integer

Public BootDelBackUp                As Byte
Public Lloviendo                    As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    VersionOK = (Ver = ULTIMAVERSION)

End Function

Sub ReSpawnOrigPosNpcs()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next
    If frmmain.Visible Then frmmain.AgregarConsola "Haciendo ReSpawn de NPCS en posicion original"

    Dim i     As Integer
    Dim MiNPC As npc
       
    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
            
            If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = _
                    Guardias Then
                MiNPC = Npclist(i)
 
                Call ReSpawnNpc(MiNPC)

            End If
            
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If
       
    Next i
    If frmmain.Visible Then frmmain.AgregarConsola "Respawn NPCS en posicion original finalizado."
End Sub

Public Sub PurgarPenas(ByVal i As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, tCiudades.Libertad.Map, tCiudades.Libertad.X, tCiudades.Libertad.Y, True)
                    Call WriteLocaleMsg((i), 121)
                    Call FlushBuffer(i)

                End If

            End If

        End If

End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, _
                      ByVal Minutos As Long, _
                      Optional ByVal GmName As String = vbNullString)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    UserList(UserIndex).Counters.Pena = Minutos
    
    Call WarpUserChar(UserIndex, tCiudades.Prision.Map, tCiudades.Prision.X, tCiudades.Prision.Y, True)
    
    If LenB(GmName) = 0 Then
        Call WriteLocaleMsg(UserIndex, 146, Minutos)
    Else
        Call WriteLocaleMsg(UserIndex, 145, GmName)

    End If


End Sub

Public Sub BorrarUsuario(ByVal UserName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"

    End If

End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")

End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer
    
    If MD5ClientesActivado = 1 Then

        For i = 0 To UBound(MD5s)

            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function

            End If

        Next i

        MD5ok = False
    Else
        MD5ok = True

    End If

End Function

Public Sub MD5sCarga()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Integer
    
    MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
    
    If MD5ClientesActivado = 1 Then
        ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))

        For LoopC = 0 To UBound(MD5s)
            MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
            MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
        Next LoopC

    End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    BanIps.Add ip
    
    Call BanIpGuardar

End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Dale  As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1

    Do While LoopC <= BanIps.count And Dale
        Dale = (BanIps.Item(LoopC) <> ip)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1

    End If

End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim n As Long
    
    n = BanIpBuscar(ip)

    If n > 0 Then
        BanIps.Remove n
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False

    End If

End Function

Public Sub BanIpGuardar()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ArchivoBanIp As String
    Dim ArchN        As Long
    Dim LoopC        As Long
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For LoopC = 1 To BanIps.count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC
    
    Close #ArchN

End Sub

Public Sub BanIpCargar()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo BanIpCargar_Err
    
    Dim ArchN        As Long
    Dim Tmp          As String
    Dim ArchivoBanIp As String
    
        If frmmain.Visible Then frmmain.AgregarConsola "Cargando lista de ips baneadas."
100    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
102    Do While BanIps.count > 0
104        BanIps.Remove 1
       Loop
    
106    ArchN = FreeFile()
108    Open ArchivoBanIp For Input As #ArchN
    
110    Do While Not EOF(ArchN)
112        Line Input #ArchN, Tmp
114        BanIps.Add Tmp
        Loop
    
116    Close #ArchN
 If frmmain.Visible Then frmmain.AgregarConsola "Se cargo la lista de ip baneadas. Operacion Realizada con exito."
        Exit Sub

BanIpCargar_Err:
118     Call RegistrarError(Err.Number, Err.description, "Admin.BanIpCargar", Erl)
120     Resume Next
        
End Sub
Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
    '***************************************************
    'Author: Unknown
    'Last Modification: 03/02/07
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    '***************************************************

    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User

    End If

End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String)
    
    On Error GoTo BanCharacter_Err
    
    Dim tuser     As Integer
    Dim userPriv  As Byte
    Dim cantPenas As Byte
    Dim rank      As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tuser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        If tuser <= 0 Then
        
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteLocaleMsg(bannerUserIndex, 484) ' No puedes banear a un personaje con mayor rango que el tuyo.
                Else

                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteLocaleMsg(bannerUserIndex, 390) ' El personaje ya está baneado
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex)
                        Call WriteLocaleMsg(bannerUserIndex, 481, UserName) ' El personaje #1 esta offline, pero ha sido baneado.
                        Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(482, .Name & "%" & UserName, 0, 8)) ' #1 ha baneado a #2.
                        
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        
                        'ponemos la pena
                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": " & Date & " " & Time)
                        
                        
                        If (userPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            
                            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(483)) '#1 ha sido baneado por banear a un Administrador.
                            
                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName)

                    End If

                End If

            Else
                Call WriteLocaleMsg(bannerUserIndex, 80) ' El personaje no existe.|

            End If

        Else

            If (UserList(tuser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteLocaleMsg(bannerUserIndex, 484) ' No puedes banear a un personaje con mayor rango que el tuyo.
            End If
            
            Call LogBan(tuser, bannerUserIndex)
            
            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(482, .Name & "%" & UserName, 0, 8)) ' #1 ha baneado a #2.
            
            'Ponemos el flag de ban a 1
            UserList(tuser).flags.Ban = 1
            
            If (UserList(tuser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                
                Call SendData(SendTarget.ToADMINS, 0, PrepareMessageLocaleMsg(483)) '#1 ha sido baneado por banear a un Administrador.
                
                Call CloseSocket(bannerUserIndex)

            End If
            
            Call LogGM(.Name, "BAN a " & UserName)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": " & Date & " " & Time)
            
            Call CloseSocket(tuser)

        End If

    End With

    Exit Sub

BanCharacter_Err:
    Call RegistrarError(Err.Number, Err.description, "Admin.BanCharacter", Erl)
    Resume Next
    
End Sub

Public Sub UnBanCharacter(ByVal UserIndex As Integer, ByVal UserName As String)
    
    On Error GoTo UnBanCharacter_Err
    
    Dim cantPenas As Byte
    
    If (InStrB(UserName, "\") <> 0) Then
        UserName = Replace(UserName, "\", "")
    End If

    If (InStrB(UserName, "/") <> 0) Then
        UserName = Replace(UserName, "/", "")
    End If
    

    If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
        Call WriteLocaleMsg(UserIndex, 80)
    Else

        If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
        
            Call UnBan(UserName)
        
            'penas
            cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(UserList(UserIndex).Name) & ": UNBAN. " & Date & " " & Time)
        
            Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & UserName)
            
            Call WriteLocaleMsg(UserIndex, 485, UserName)
        Else
        
            Call WriteLocaleMsg(UserIndex, 486)

        End If

    End If
    
    Exit Sub

UnBanCharacter_Err:
    Call RegistrarError(Err.Number, Err.description, "Admin.UnBanCharacter", Erl)
    Resume Next
    
End Sub

