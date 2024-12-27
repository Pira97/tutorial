Attribute VB_Name = "ModFacciones"
'---------------------------------------------------------------------------------------
' LinkAO - V1.0                                                                       '
' Fecha     : 01/08/2020                                                               '
' Fecha 2   : 16/06/2021 módulo finalizado
' Module    : Modulo para manejar las facciones                                        '
'---------------------------------------------------------------------------------------
'Reformas, el admin puede hacer todo igual
Option Explicit
Private Const PERDON   As Long = 100000
Public Sub EntrarRepublica(ByVal UserIndex As Integer)
On Error GoTo EntrarRepublica_Err

       With UserList(UserIndex)
        
        If esRepu(UserIndex) Then
        Call WriteChatOverHead(UserIndex, "¡No puedo hacerte republicano porque ya lo eres!.", str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Exit Sub
        End If
        
        If .Faccion.Status <> 1 Then
            Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de facciones enemigas, ¡largo de aqui", str(Npclist(.flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
 
        If UserList(UserIndex).GuildIndex > 0 Then
        Call WriteChatOverHead(UserIndex, "¡Para realizar esta acción no debes pertenecer a ningún clan!.", str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Exit Sub
        End If
    
        If .flags.Meditando Then
        Dim Estado As Byte
        Estado = ParticleToLevel(UserIndex)
        End If
     
        If esRene(UserIndex) Then
            If Not .Stats.GLD > PERDON Then
                   Call WriteLocaleMsg(UserIndex, 383, PERDON)
               Exit Sub
            End If
           .Stats.GLD = .Stats.GLD - PERDON
           .Faccion.Status = 3
           .Hogar = cIlliandor
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, .Faccion.Status))
            Call WriteUpdateGold(UserIndex)
            Call WriteChatOverHead(UserIndex, "Bienvenido a la República.", str(Npclist(.flags.TargetNPC).Char.CharIndex))
        End If
                            
        If .flags.Meditando Then
        If Estado <> ParticleToLevel(UserIndex) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Estado, 0, True, True))
        End If
        End If

    End With
    
    Exit Sub

EntrarRepublica_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.EntrarRepublica", Erl)
       Resume Next
End Sub
Public Sub EntrarImperial(ByVal UserIndex As Integer)
On Error GoTo EntrarImperial_Err

       With UserList(UserIndex)
       
        If esCiuda(UserIndex) Then
        Call WriteChatOverHead(UserIndex, "¡No puedo hacerte un ciudadano imperial porque ya lo eres!.", str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Exit Sub
        End If
        
        If .Faccion.Status <> 1 Then
            Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de facciones enemigas, ¡largo de aqui", str(Npclist(.flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
 
        If UserList(UserIndex).GuildIndex > 0 Then
        Call WriteChatOverHead(UserIndex, "¡Para realizar esta acción no debes pertenecer a ningún clan!.", str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Exit Sub
        End If

        If .flags.Meditando Then
        Dim Estado As Byte
        Estado = ParticleToLevel(UserIndex)
        End If
     
        If esRene(UserIndex) Then
            If Not .Stats.GLD > PERDON Then
               Call WriteLocaleMsg(UserIndex, 383, PERDON)
               Exit Sub
            End If
           .Stats.GLD = .Stats.GLD - PERDON
           .Faccion.Status = 2
           .Hogar = cNix
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, .Faccion.Status))
            
            Call WriteUpdateGold(UserIndex)
            Call WriteChatOverHead(UserIndex, "Bienvenido al Imperio.", str(Npclist(.flags.TargetNPC).Char.CharIndex))
        End If
                            
        If .flags.Meditando Then
        If Estado <> ParticleToLevel(UserIndex) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Estado, 0, True, True))
        End If
        End If

    End With
    
    Exit Sub

EntrarImperial_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.EntrarImperial", Erl)
       Resume Next
End Sub

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
On Error GoTo EnlistarCaos_Err

       With UserList(UserIndex)

       Dim Matados As Long
       
Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.RepublicanosMatados)
        Matados = UserList(UserIndex).Stats.NPCsMuertos
 
        If esCaos(UserIndex) Then
         Call WriteLocaleMsg(UserIndex, 1)
            Exit Sub

        End If
        
         If .Faccion.Status <> 1 Then
            Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de facciones enemigas, ¡largo de aqui", str(Npclist(.flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
        
        If Matados < 1 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 40 criaturas, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
 
        If UserList(UserIndex).Stats.ELV < 40 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos nivel 40.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
        
        If .flags.Meditando Then
        Dim Estado As Byte
        Estado = ParticleToLevel(UserIndex)
        End If

        UserList(UserIndex).Faccion.Status = 4
        UserList(UserIndex).Faccion.Rango = 1
        
        '------- Ropa -------
        Dim MiObj As Obj
        MiObj.Amount = 1
    
        If UserList(UserIndex).raza = enano Or UserList(UserIndex).raza = gnomo Then
            Dim bajos As Byte
            bajos = 1
        End If
    
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo
                MiObj.ObjIndex = 1500 + bajos
            Case eClass.Mago
                MiObj.ObjIndex = 1502 + bajos
            Case eClass.Guerrero
                MiObj.ObjIndex = 1504 + bajos
            Case eClass.Asesino
                MiObj.ObjIndex = 1506 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1508 + bajos
            Case eClass.Druida
                MiObj.ObjIndex = 1510 + bajos
            Case eClass.Gladiador
                MiObj.ObjIndex = 1512 + bajos
            Case eClass.Paladin
                MiObj.ObjIndex = 1514 + bajos
            Case eClass.Cazador
                MiObj.ObjIndex = 1516 + bajos
            Case eClass.Mercenario
                MiObj.ObjIndex = 1518 + bajos
            Case eClass.Nigromante
                MiObj.ObjIndex = 1520 + bajos
        End Select
 
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
         If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
         End If
        End If
 
    
       If UserList(UserIndex).GuildIndex > 0 Then
           Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
           Call WriteLocaleMsg(UserIndex, 1, UserList(UserIndex).Name)
       End If
    
        Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a la Hordas del Caos!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, .Faccion.Status))
        
        If .flags.Meditando Then
        If Estado <> ParticleToLevel(UserIndex) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Estado, 0, True, True))
        End If
        End If

    End With
    
    Exit Sub

EnlistarCaos_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.EnlistarCaos", Erl)
       Resume Next
End Sub

Public Sub EnlistarMilicia(ByVal UserIndex As Integer)
On Error GoTo EnlistarMilicia_Err

       With UserList(UserIndex)
    
       Dim Matados As Long
       
       'Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados)
        Matados = UserList(UserIndex).Stats.NPCsMuertos
 
        If esMili(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "Ya perteneces a las tropas milicianas, ve a combatir enemigos", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
        
         If .Faccion.Status <> 3 Then
            Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de facciones enemigas, ¡largo de aqui", str(Npclist(.flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
        
        If Matados < 1 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 1 criaturas, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
 
        If UserList(UserIndex).Stats.ELV < 25 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos nivel 25.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If

        If .flags.Meditando Then
        Dim Estado As Byte
        Estado = ParticleToLevel(UserIndex)
        End If

        UserList(UserIndex).Faccion.Status = 6
        UserList(UserIndex).Faccion.Rango = 1
        
        '------- Ropa -------
        Dim MiObj As Obj
        MiObj.Amount = 1
    
        If UserList(UserIndex).raza = enano Or UserList(UserIndex).raza = gnomo Then
        MiObj.ObjIndex = 1587
        Else
        MiObj.ObjIndex = 1588
        End If
 
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
         If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
         End If
        End If
 
    
       If UserList(UserIndex).GuildIndex > 0 Then
           Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
           Call WriteLocaleMsg(UserIndex, 1, UserList(UserIndex).Name)
       End If
    
       Call WriteChatOverHead(UserIndex, "Bienvenido a la Milicia Republicana, aqui tienes tu Armadura. Cumple bien tu labor exterminando criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, .Faccion.Status))
        
        If .flags.Meditando Then
        If Estado <> ParticleToLevel(UserIndex) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Estado, 0, True, True))
        End If
        End If
        
    End With
    
    Exit Sub

EnlistarMilicia_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.EnlistarMilicia", Erl)
       Resume Next
End Sub

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
On Error GoTo EnlistarArmadaReal_Err

       With UserList(UserIndex)
  
       Dim Matados As Long
       
       'Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.RepublicanosMatados + UserList(UserIndex).Faccion.MilicianosMatados)
        Matados = UserList(UserIndex).Stats.NPCsMuertos
 
        If esArmada(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "Ya perteneces a las tropas reales, ve a combatir enemigos", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
        
         If .Faccion.Status <> 2 Then
            Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de facciones enemigas, ¡largo de aqui", str(Npclist(.flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
        
        If Matados < 1 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 1 criaturas, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
 
        If UserList(UserIndex).Stats.ELV < 25 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos nivel 25.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If

        If .flags.Meditando Then
        Dim Estado As Byte
        Estado = ParticleToLevel(UserIndex)
        End If

        UserList(UserIndex).Faccion.Status = 5
        UserList(UserIndex).Faccion.Rango = 1
        
        '------- Ropa -------
        Dim MiObj As Obj
        MiObj.Amount = 1
    
        If UserList(UserIndex).raza = enano Or UserList(UserIndex).raza = gnomo Then
        Dim bajos As Byte
        bajos = 1
        End If
 
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo
                MiObj.ObjIndex = 1544 + bajos
            Case eClass.Mago
                MiObj.ObjIndex = 1546 + bajos
            Case eClass.Guerrero
                MiObj.ObjIndex = 1548 + bajos
            Case eClass.Asesino
                MiObj.ObjIndex = 1550 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1552 + bajos
            Case eClass.Druida
                MiObj.ObjIndex = 1554 + bajos
            Case eClass.Gladiador
                MiObj.ObjIndex = 1556 + bajos
            Case eClass.Paladin
                MiObj.ObjIndex = 1558 + bajos
            Case eClass.Cazador
                MiObj.ObjIndex = 1560 + bajos
            Case eClass.Mercenario
                MiObj.ObjIndex = 1562 + bajos
            Case eClass.Nigromante
                MiObj.ObjIndex = 1564 + bajos
        End Select

        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
         If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
         End If
        End If
 
    
       If UserList(UserIndex).GuildIndex > 0 Then
           Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
           Call WriteLocaleMsg(UserIndex, 1, UserList(UserIndex).Name)
       End If
    
       Call WriteChatOverHead(UserIndex, "Bienvenido al Ejército Imperial, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, .Faccion.Status))
        
        If .flags.Meditando Then
        If Estado <> ParticleToLevel(UserIndex) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1, False, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, Estado, 0, True, True))
        End If
        End If

    End With
    
    Exit Sub

EnlistarArmadaReal_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.EnlistarArmadaReal", Erl)
       Resume Next
End Sub

 
Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

On Error GoTo RecompensaArmadaReal_Err

    If UserList(UserIndex).Faccion.Rango = 10 Then
    Call WriteChatOverHead(UserIndex, "Ya no tengo trabajo para darte, has alcanzado el rango más alto aquí.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
    End If
    
    Dim Matados As Long
    
    'Matados = UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.RepublicanosMatados
    Matados = UserList(UserIndex).Stats.NPCsMuertos
    
    If Matados < matadosArmada(UserList(UserIndex).Faccion.Rango) Then
    Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).Faccion.Rango) - Matados & " criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
    End If
 
    UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
    
    Call WriteChatOverHead(UserIndex, "Felicidades, has subido de rango.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    
    If UserList(UserIndex).Faccion.Rango >= 6 Then ' Segunda jeraquia
    
        Dim MiObj As Obj
        MiObj.Amount = 1
         
        If UserList(UserIndex).raza = enano Or UserList(UserIndex).raza = gnomo Then
        Dim bajos As Byte
        bajos = 1
        End If
        
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo
                MiObj.ObjIndex = 1566 + bajos
            Case eClass.Mago
                MiObj.ObjIndex = 1568 + bajos
            Case eClass.Guerrero
                MiObj.ObjIndex = 1570 + bajos
            Case eClass.Asesino
                MiObj.ObjIndex = 1572 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1574 + bajos
            Case eClass.Druida
                MiObj.ObjIndex = 1576 + bajos
            Case eClass.Gladiador
                MiObj.ObjIndex = 1578 + bajos
            Case eClass.Paladin
                MiObj.ObjIndex = 1580 + bajos
            Case eClass.Cazador
                MiObj.ObjIndex = 1582 + bajos
            Case eClass.Mercenario
                MiObj.ObjIndex = 1584 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1586 + bajos
        End Select

        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
          If Not MeterItemEnInventario(UserIndex, MiObj) Then
           Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
         End If
        End If
        
    End If
    
    Exit Sub

RecompensaArmadaReal_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.RecompensaArmadaReal", Erl)
       Resume Next
End Sub
Public Sub RecompensaMilicia(ByVal UserIndex As Integer)

    On Error GoTo RecompensaMilicia_Err
    
    If UserList(UserIndex).Faccion.Rango = 7 Then
    Call WriteChatOverHead(UserIndex, "Ya no tengo trabajo para darte, has alcanzado el rango más alto aquí.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
    End If
    
    Dim Matados As Long
 
    'Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados)
    Matados = UserList(UserIndex).Stats.NPCsMuertos
 
     If Matados < matadosArmada(UserList(UserIndex).Faccion.Rango) Then
        Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).Faccion.Rango) - Matados & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
        Exit Sub
    End If
 
    UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
    Call WriteChatOverHead(UserIndex, "Felicidades, has subido de rango.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
        
    If UserList(UserIndex).Faccion.Rango >= 4 Then
    
        Dim MiObj As Obj
        
        MiObj.Amount = 1
        
        If UserList(UserIndex).raza = enano Or UserList(UserIndex).raza = gnomo Then
        Dim bajos As Byte
        bajos = 1
        End If
            
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo, eClass.Mago, eClass.Bardo, eClass.Druida, eClass.Nigromante
                MiObj.ObjIndex = 1592 + bajos
            Case eClass.Gladiador, eClass.Guerrero, eClass.Cazador, eClass.Mercenario, eClass.Paladin, eClass.Asesino
                MiObj.ObjIndex = 1590 + bajos
        End Select
     
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
          If Not MeterItemEnInventario(UserIndex, MiObj) Then
           Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
         End If
        End If
    End If
    
    Exit Sub

RecompensaMilicia_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.RecompensaMilicia", Erl)
       Resume Next
End Sub
Public Sub RecompensaCaos(ByVal UserIndex As Integer)

    On Error GoTo RecompensaCaos_Err

        If UserList(UserIndex).Faccion.Rango = 10 Then
        Call WriteChatOverHead(UserIndex, "Ya no tengo trabajo para darte, has alcanzado el rango más alto aquí.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
        Exit Sub
        End If
 
       Dim Matados As Long
        Matados = UserList(UserIndex).Stats.NPCsMuertos

       'Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.RepublicanosMatados)
 
        If Matados < matadosCaos(UserList(UserIndex).Faccion.Rango) Then
            Call WriteChatOverHead(UserIndex, "Mata " & matadosCaos(UserList(UserIndex).Faccion.Rango) - Matados & " enemigos más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
 
        UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
        Call WriteChatOverHead(UserIndex, "Felicidades, has subido de rango.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    If UserList(UserIndex).Faccion.Rango >= 6 Then ' Segunda jeraquia
        
        Dim MiObj As Obj
        MiObj.Amount = 1
         
        If UserList(UserIndex).raza = enano Or UserList(UserIndex).raza = gnomo Then
        Dim bajos As Byte
        bajos = 1
        End If
        
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo
                MiObj.ObjIndex = 1522 + bajos
            Case eClass.Mago
                MiObj.ObjIndex = 1524 + bajos
            Case eClass.Guerrero
                MiObj.ObjIndex = 1526 + bajos
            Case eClass.Asesino
                MiObj.ObjIndex = 1528 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1530 + bajos
            Case eClass.Druida
                MiObj.ObjIndex = 1532 + bajos
            Case eClass.Gladiador
                MiObj.ObjIndex = 1534 + bajos
            Case eClass.Paladin
                MiObj.ObjIndex = 1536 + bajos
            Case eClass.Cazador
                MiObj.ObjIndex = 1538 + bajos
            Case eClass.Mercenario
                MiObj.ObjIndex = 1540 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1542 + bajos
        End Select
 
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
               If Not MeterItemEnInventario(UserIndex, MiObj) Then
                   Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
               End If
         End If
    End If
    Exit Sub

RecompensaCaos_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.RecompensaCaos", Erl)
       Resume Next
       
End Sub
Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)

    On Error GoTo ExpulsarFaccionReal_Err:

    If UserList(UserIndex).Faccion.Status <> 5 Then Exit Sub
    
    UserList(UserIndex).Faccion.Status = 2
    UserList(UserIndex).Faccion.Rango = 0
    
    Call WriteLocaleMsg(UserIndex, 384)
 
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
  
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Faccion.Status))
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

    Exit Sub

ExpulsarFaccionReal_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.ExpulsarFaccionReal", Erl)
       Resume Next
       
End Sub
Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

    On Error GoTo ExpulsarFaccionCaos_Err:
    
    If UserList(UserIndex).Faccion.Status <> 4 Then Exit Sub
    
    UserList(UserIndex).Faccion.Status = 1
    UserList(UserIndex).Faccion.Rango = 0
    

    
    Call WriteLocaleMsg(UserIndex, 384)
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Faccion.Status))
       If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)


    Exit Sub

ExpulsarFaccionCaos_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.ExpulsarFaccionCaos", Erl)
       Resume Next
End Sub
Public Sub ExpulsarFaccionMilicia(ByVal UserIndex As Integer)

    
    On Error GoTo ExpulsarFaccionMilicia_Err:
    
    If UserList(UserIndex).Faccion.Status <> 6 Then Exit Sub
    
    UserList(UserIndex).Faccion.Status = 1
    UserList(UserIndex).Faccion.Rango = 0
    

    
    Call WriteLocaleMsg(UserIndex, 384)
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Milicia = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Milicia = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Faccion.Status))
       If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)


    Exit Sub


ExpulsarFaccionMilicia_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.ExpulsarFaccionMilicia", Erl)
       Resume Next
End Sub
Public Function TituloCaos(ByVal UserIndex As Integer) As String
    
    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloCaos = "Miembro de las Hordas"
        Case 2
            TituloCaos = "Guerrero del Caos"
        Case 3
            TituloCaos = "Teniente del Caos"
        Case 4
            TituloCaos = "Comandante del Caos"
        Case 5
            TituloCaos = "General del Caos"
        Case 6
            TituloCaos = "Elite del Caos"
        Case 7
            TituloCaos = "Asolador de las Sombras"
        Case 8
            TituloCaos = "Caballero Negro"
        Case 9
            TituloCaos = "Emisario de las Sombras"
        Case 10
            TituloCaos = "Avatar del Apocalipsis"
    End Select

       
End Function
Public Function TituloReal(ByVal UserIndex As Integer) As String

    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloReal = "Legionario"
        Case 2
            TituloReal = "Soldado Real"
        Case 3
            TituloReal = "Teniente Real"
        Case 4
            TituloReal = "Comandante Real"
        Case 5
            TituloReal = "General Real"
        Case 6
            TituloReal = "Elite Real"
        Case 7
            TituloReal = "Guardian del Bien"
        Case 8
            TituloReal = "Caballero Imperial"
        Case 9
            TituloReal = "Justiciero"
        Case 10
            TituloReal = "Guardia Imperial"
    End Select
    
End Function
Public Function TituloMilicia(ByVal UserIndex As Integer) As String

    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloMilicia = "Milicia de Reserva"
        Case 2
            TituloMilicia = "Miliciano"
        Case 3
            TituloMilicia = "Miliciano Elite"
        Case 4
            TituloMilicia = "Soldado de la República"
        Case 5
            TituloMilicia = "Soldado Raso"
        Case 6
            TituloMilicia = "Soldado Elite"
        Case 7
            TituloMilicia = "Comandante de la República"
    End Select
    

       
End Function
Public Function matadosArmada(ByVal Rango As Byte) As Integer

    Select Case Rango
        Case 1
            matadosArmada = 1300
        Case 2
            matadosArmada = 1400
        Case 3
            matadosArmada = 1500
        Case 4
            matadosArmada = 1600
        Case 5
            matadosArmada = 1700
        Case 6
            matadosArmada = 1800
        Case 7
            matadosArmada = 1900
        Case 8
            matadosArmada = 2000
        Case 9
            matadosArmada = 2200
    End Select
       
End Function
Public Function matadosCaos(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosCaos = 1
        Case 2
            matadosCaos = 2
        Case 3
            matadosCaos = 3
        Case 4
            matadosCaos = 4
        Case 5
            matadosCaos = 5
        Case 6
            matadosCaos = 6
        Case 7
            matadosCaos = 7
        Case 8
            matadosCaos = 8
        Case 9
            matadosCaos = 9
    End Select

End Function
Public Function matadosMilicia(ByVal Rango As Byte) As Integer

    Select Case Rango
        Case 1
            matadosMilicia = 1
        Case 2
            matadosMilicia = 2
        Case 3
            matadosMilicia = 3
        Case 4
            matadosMilicia = 4
        Case 5
            matadosMilicia = 1700
        Case 6
            matadosMilicia = 1800
    End Select
    

End Function
 
Public Sub QuitarItemsFaccionarios(ByVal UserIndex As Integer)

    On Error GoTo QuitarItemsFaccionarios_Err:
    
    Dim i As Byte
    Dim ObjIndex As Integer
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ObjIndex <> 1 Then
            If ObjData(ObjIndex).Caos = 1 Or ObjData(ObjIndex).Real = 1 Or ObjData(ObjIndex).Milicia = 1 Then
                QuitarUserInvItem UserIndex, i, UserList(UserIndex).Invent.Object(i).Amount
                UpdateUserInv False, UserIndex, i
            End If
        End If
    Next i

    Exit Sub

QuitarItemsFaccionarios_Err:
       Call RegistrarError(Err.Number, Err.description, "ModFacciones.QuitarItemsFaccionarios", Erl)
       Resume Next
       
End Sub
