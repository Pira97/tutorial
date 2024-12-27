Attribute VB_Name = "SistemaCombate"

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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat

Option Explicit
Public Const NPC_DEMONIO As Integer = 1
Public Const ARCO_DEMONIO As Integer = 666
Public Const MAXDISTANCIAARCO  As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a

    End If

End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b

    End If

End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList( _
            UserIndex).Clase).Escudo) / 2

End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim lTemp As Long

    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos( _
                eAtributos.Agilidad)) * ModClase(.Clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)

        If .Stats.UserSkills(eSkill.armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.armas) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueArmas

        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)

        If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) _
                    * ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) _
                    * ModClase(.Clase).AtaqueProyectiles

        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, _
                               ByVal npcindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaque As Long
    Dim Arma        As Integer
    Dim Skill       As eSkill
    Dim ProbExito   As Long
 
     Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
  
    If UserList(UserIndex).Invent.NudiEqpObjIndex > 0 Then
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.Wrestling
    ElseIf Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles
        ElseIf ObjData(Arma).proyectil = 2 Then
            PoderAtaque = PoderAtaqueArpon(UserIndex)
            Skill = eSkill.ArmasArrojadizas
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.armas
        End If
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.Wrestling
 

    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(npcindex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill)
    End If

End Function

Public Function NpcImpacto(ByVal npcindex As Integer, _
                           ByVal UserIndex As Integer) As Boolean
    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Revisa si un NPC logra impactar a un user o no
    '03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
    '*************************************************
    Dim Rechazo           As Boolean
    Dim ProbRechazo       As Long
    Dim ProbExito         As Long
    Dim UserEvasion       As Long
    Dim NpcPoderAtaque    As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas     As Long
    Dim SkillDefensa      As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(npcindex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
 
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        
        If Not NpcImpacto Then
        
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                     Call WriteLocaleMsg(UserIndex, 73)
                     Call SubirSkill(UserIndex, eSkill.Defensa)
                    
                End If

            End If

        End If

    End If

End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal npcindex As Integer = 0) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: 01/04/2010 (ZaMa)
    '01/04/2010: ZaMa - Modifico el daño de wrestling.
    '01/04/2010: ZaMa - Agrego bonificadores de wrestling para los guantes.
    '***************************************************
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma       As ObjData
    Dim ModifClase As Single
    Dim proyectil  As ObjData
    Dim DañoMaxArma As Long
    Dim DañoMinArma As Long
    Dim ObjIndex   As Integer
    Dim nudis As ObjData
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(UserIndex)

         If .Invent.WeaponEqpObjIndex > 0 And .Invent.NudiEqpSlot = 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If npcindex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                     If .Invent.WeaponEqpObjIndex = ARCO_DEMONIO Then ' Usa la arco mata Demonios?
                        If Npclist(npcindex).Numero = NPC_DEMONIO Then
                            DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            DañoMaxArma = Arma.MaxHIT
                        Else
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                     End If
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)

                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If

               If Arma.proyectil = 2 Then
                    ModifClase = ModClase(.Clase).DañoArpon
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                    DañoArma = DañoArma + RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                 End If
                Else
                    ModifClase = ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(npcindex).NPCtype = Dragon Then 'Ataca Dragon?
                            DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            DañoMaxArma = Arma.MaxHIT * 50
                            matoDragon = False ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon daño es 1
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT
                    End If

                End If

            Else ' Ataca usuario

                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                    If .Invent.WeaponEqpObjIndex = ARCO_DEMONIO Then
                        DañoArma = 1
                        DañoMaxArma = 1
                    Else
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                     End If
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)

                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If
                    If Arma.proyectil = 2 Then
                       ModifClase = ModClase(.Clase).DañoArpon
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    'DañoMaxArma = Arma.MaxHIT
                    End If
                Else
                    ModifClase = ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.Clase).DañoArmas
                        DañoArma = RandomNumber(22, 35)
                        DañoMaxArma = 35
                    Else
                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        DañoMaxArma = Arma.MaxHIT

                    End If

                End If

            End If
 


 
       Else
        
            ModifClase = ModClase(.Clase).DañoWrestling
             If .Invent.NudiEqpObjIndex > 0 Then
                Arma = ObjData(.Invent.NudiEqpObjIndex)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
            End If
 
            
              DañoArma = DañoArma + RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
            DañoMaxArma = DañoMaxArma + 3
        End If
        
        DañoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDaño = Npclist(npcindex).Stats.MinHP + Npclist(npcindex).Stats.def
        Else
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DañoUsuario) * ModifClase
                If .flags.Montando = 1 Then
                Dim Obj        As ObjData
                Dim hitmontura As Integer
                Obj = ObjData(.Invent.MonturaObjIndex)
                hitmontura = RandomNumber(Obj.MinHIT, Obj.MaxHIT)
                CalcularDaño = CalcularDaño + hitmontura
                End If

                If UserList(UserIndex).Invent.MagicIndex > 0 And npcindex <> 0 Then
                    If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.AumentaGolpe Then
                        CalcularDaño = CalcularDaño + ObjData(UserList(UserIndex).Invent.MagicIndex).CuantoAumento
                    End If
                End If
        
        End If

    End With

End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal npcindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 07/04/2010 (ZaMa)
    '25/01/2010: ZaMa - Agrego poder acuchillar npcs.
    '07/04/2010: ZaMa - Los asesinos apuñalan acorde al daño base sin descontar la defensa del npc.
    '***************************************************

    Dim daño As Long
    Dim DañoBase As Long
    
    DañoBase = CalcularDaño(UserIndex, npcindex)
    
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 Then
        If UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
            DañoBase = DañoBase + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData( _
                    UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

        End If

    End If
    
    With Npclist(npcindex)
        daño = DañoBase - .Stats.def
        
        Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, daño, 2)


        Call CalcularDarExp(UserIndex, npcindex, daño)
        
        .Stats.MinHP = .Stats.MinHP - daño
        
        If .Stats.MinHP > 0 Then

            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(UserIndex) Then
                Call DoApuñalar(UserIndex, npcindex, 0, DañoBase)

            End If
            
        End If
        
        If .Stats.MinHP <= 0 Then

            ' $ Shermie80 / Creamos un log de quien mato un dragón
            If .NPCtype = Dragon Then

                If .Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(UserIndex).Name & " mató un dragón")

            End If
            ' $ Fin
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer

            For j = 1 To MAXMASCOTAS

                If UserList(UserIndex).MascotasIndex(j) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = npcindex Then
                        Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo

                    End If

                End If

            Next j
            
            Call MuereNpc(npcindex, UserIndex)

        End If

    End With

End Sub

Public Sub NpcDaño(ByVal npcindex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim daño       As Integer
    Dim lugar      As Integer
    Dim absorbido  As Integer
    Dim defbarco   As Integer
    Dim defmontura As Integer
    Dim Obj        As ObjData
    Dim hitmontura As Integer
    daño = RandomNumber(Npclist(npcindex).Stats.MinHIT, Npclist(npcindex).Stats.MaxHIT)
    
    With UserList(UserIndex)

        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)

        End If
        
        If .flags.Montando = 1 Then
            Obj = ObjData(.Invent.MonturaObjIndex)
            defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case lugar

            Case PartesCuerpo.bCabeza

                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(.Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

            Case Else

                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Invent.ArmourEqpObjIndex)

                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)

                    End If

                End If

        End Select
        
        absorbido = absorbido + defbarco + defmontura
        daño = daño - absorbido
        
        
     
            
        If daño < 1 Then daño = 1

        Call WriteLocaleMsg(UserIndex, lugar, daño, 2)
        
        Call WriteChatOverHeadLocale(UserIndex, Npclist(npcindex).Char.CharIndex, daño, 2)
        
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHP = .Stats.MinHP - daño
        
        If .flags.Meditando Then
            If daño > Fix(.Stats.MinHP / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteLocaleMsg(UserIndex, 123)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectoCharParticula(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), 0, True, True))
            
            End If
        End If

        'Muere el usuario
        If .Stats.MinHP <= 0 Then
            Call WriteLocaleMsg(UserIndex, 72, "$" & Npclist(npcindex).Numero)
            
            If Npclist(npcindex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(npcindex).MaestroUser)
            Else

                'Al matarlo no lo sigue mas
                If Npclist(npcindex).Movement = 0 Then
                    Npclist(npcindex).Movement = Npclist(npcindex).flags.OldMovement
                    Npclist(npcindex).Hostile = Npclist(npcindex).flags.OldHostil
                    Npclist(npcindex).flags.AttackedBy = vbNullString

                End If

            End If
            
            Call UserDie(UserIndex)

        End If

    End With

End Sub

Public Sub CheckPets(ByVal npcindex As Integer, _
                     ByVal UserIndex As Integer, _
                     Optional ByVal CheckElementales As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/04/2010
    '15/04/2010: ZaMa - Las mascotas no se apropian de npcs.
    '***************************************************

    Dim j As Integer
    
    ' Si no tengo mascotas, para que cheaquear lo demas?
    If UserList(UserIndex).NroMascotas = 0 Then Exit Sub
    
    If Not PuedeAtacarNPC(UserIndex, npcindex, , True) Then Exit Sub
    
    With UserList(UserIndex)

        For j = 1 To MAXMASCOTAS

            If .MascotasIndex(j) > 0 Then
                If .MascotasIndex(j) <> npcindex Then
                    If CheckElementales Or (Npclist(.MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist( _
                            .MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                    
                        If Npclist(.MascotasIndex(j)).TargetNPC = 0 Then Npclist(.MascotasIndex(j)).TargetNPC = npcindex
                        Npclist(.MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc

                    End If

                End If

            End If

        Next j

    End With

End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))

        End If

    Next j

End Sub

Public Function NpcAtacaUser(ByVal npcindex As Integer, _
                             ByVal UserIndex As Integer) As Boolean
    '*************************************************
    'Author: Unknown
    'Last modified: -
    '
    '*************************************************

    With UserList(UserIndex)
    
        If .flags.AdminInvisible = 1 Then Exit Function
        
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
        

        If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Function
        
        If .flags.Muerto = 1 Then Exit Function
    End With
    
    With Npclist(npcindex)

        ' El npc puede atacar ???
        If IntervaloPermiteAtacarNpc(npcindex) Then
            NpcAtacaUser = True
            Call CheckPets(npcindex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = npcindex

            End If

        Else
            NpcAtacaUser = False
            Exit Function

        End If

        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))

        End If

    End With
    
    If NpcImpacto(npcindex, UserIndex) Then

        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 And .flags.Montando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDaño(npcindex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '¿Puede envenenar?
            If Npclist(npcindex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex)

        End With
    Else
        Call WriteChatOverHeadLocale(UserIndex, Npclist(npcindex).Char.CharIndex, 0, 2) 'Fallas
    End If
    
    Call SubirSkill(UserIndex, eSkill.Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)

End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, _
                               ByVal Victima As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtt  As Long
    Dim PoderEva  As Long
    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim daño As Integer
    
    With Npclist(Atacante)
        daño = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño
        
        If Npclist(Victima).Stats.MinHP < 1 Then
            .Movement = .flags.OldMovement
           If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil

            End If
            
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)

            End If
            
            Call MuereNpc(Victima, .MaestroUser)

        End If

    End With

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, _
                       ByVal Victima As Integer, _
                       Optional ByVal cambiarMOvimiento As Boolean = True)
    '*************************************************
    'Author: Unknown
    'Last modified: 01/03/2009
    '*************************************************
    
    With Npclist(Atacante)

        ' El npc puede atacar ???
        If IntervaloPermiteAtacarNpc(Atacante) Then
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc

            End If

        Else
            Exit Sub

        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))

        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, _
                        Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
                                        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist( _
                        Victima).Pos.X, Npclist(Victima).Pos.Y))
                        
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist( _
                        Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist( _
                        Victima).Pos.X, Npclist(Victima).Pos.Y))

            End If
            
            Call NpcDañoNpc(Atacante, Victima)
        Else

            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist( _
                        Victima).Pos.X, Npclist(Victima).Pos.Y))

            End If

        End If

    End With

End Sub
Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, _
                                ByVal npcindex As Integer) As Boolean


    On Error GoTo ErrHandler
 
    
    'If UserList(UserIndex).Invent.WeaponEqpObjIndex = 1608 And EsGm(UserIndex) Then
    'Call MuereNpc(npcindex, UserIndex)
    'Exit Function
    'End If
        
    If Not PuedeAtacarNPC(UserIndex, npcindex) Then Exit Function
    
    If UserList(UserIndex).flags.Oculto = 1 Or UserList(UserIndex).flags.Invisible = 1 Then
       UserList(UserIndex).flags.Oculto = 0
       UserList(UserIndex).flags.Invisible = 0
       UserList(UserIndex).Counters.Invisibilidad = 0
       Call WriteLocaleMsg(UserIndex, 307)
       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
    
    Call NPCAtacado(npcindex, UserIndex)
    
    If UserImpactoNpc(UserIndex, npcindex) Then
        If Npclist(npcindex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessagePlayWave(Npclist(npcindex).flags.Snd2, _
                    Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
        End If
   
            Select Case UserList(UserIndex).Invent.WeaponEqpObjIndex
            Case Is > 0
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 And UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessagePlayWave(ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).Snd1, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
                    Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessageCreateFX(Npclist(npcindex).Char.CharIndex, ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).Snd2, 0))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO3, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
                    ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 2 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(68, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
                    Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
                    End If
            Case Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
            End Select
            
                
            Call GolpeParalizaNPC(UserIndex, npcindex)

            Call UserDañoNpc(UserIndex, npcindex)
 
    Else
            Select Case UserList(UserIndex).Invent.WeaponEqpObjIndex
            Case Is > 0
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 And UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_FALLASFLECHA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 2 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(67, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    End If
            Case Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

            End Select
            
            Call WriteChatOverHeadLocale(UserIndex, UserList(UserIndex).Char.CharIndex, 0, 2) 'Fallas
    End If

 
    ' Reveló su condición de usuario al atacar, los npcs lo van a atacar
    UserList(UserIndex).flags.Ignorado = False
    
    UsuarioAtacaNpc = True
    
    Exit Function
   
ErrHandler:
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.description)
    
End Function
Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim index     As Integer
    Dim AttackPos As WorldPos
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Magic interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex) Then

        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub

        End If

    End If
    
    With UserList(UserIndex)

        'Quitamos stamina
        
        Dim EnergiaGolpe As Byte
        
        If .Invent.NudiEqpObjIndex > 0 Then
            EnergiaGolpe = ObjData(UserList(UserIndex).Invent.NudiEqpObjIndex).StaRequerido
        ElseIf .Invent.WeaponEqpObjIndex > 0 Then
            EnergiaGolpe = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaRequerido
        Else
            EnergiaGolpe = RandomNumber(1, 10)
        End If
        
        If .Stats.MinSta >= EnergiaGolpe Then
            Call QuitarSta(UserIndex, EnergiaGolpe)
        Else
            Call WriteLocaleMsg(UserIndex, 93)
        Exit Sub
        End If
        
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > _
                YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub

        End If
        
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
        
        'Look for user
        If index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, index)
            Exit Sub
        End If
        
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).npcindex
        
        'Look for NPC
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.Map).Pk = False Then
                    'Call WriteMensajes(UserIndex, eMensajes.Mensaje165)
                    Exit Sub

                End If
                
                Call UsuarioAtacaNpc(UserIndex, index)
            Else
                Call WriteLocaleMsg(UserIndex, 298)

            End If
            
            Exit Sub

        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
            
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

    End With

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, _
                               ByVal VictimaIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim ProbRechazo            As Long
    Dim Rechazo                As Boolean
    Dim ProbExito              As Long
    Dim PoderAtaque            As Long
    Dim UserPoderEvasion       As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma                   As Integer
    Dim SkillTacticas          As Long
    Dim SkillDefensa           As Long
    Dim ProbEvadir             As Long
    Dim Skill                  As eSkill
    
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
    
    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
        UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0

    End If
    
    'Esta usando un arma ???
    
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
            Skill = eSkill.Proyectiles
                ElseIf ObjData(Arma).proyectil = 2 Then
            PoderAtaque = PoderAtaqueArpon(AtacanteIndex)
            Skill = eSkill.ArmasArrojadizas
            
        Else
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)
            Skill = eSkill.armas

        End If

    Else
        PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
        Skill = eSkill.Wrestling

    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    ' Se reduce la evasion un 25%
    If UserList(VictimaIndex).flags.Meditando = True Then
        ProbEvadir = (100 - ProbExito) * 0.75
        ProbExito = MinimoInt(90, 100 - ProbEvadir)

    End If
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then

        'Fallo ???
        If Not UsuarioImpacto Then
            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

            If Rechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList( _
                        VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
                   
                Call WriteConsoleMsg(AtacanteIndex, "¡El usuario rechazó el ataque con su escudo!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡Has rechazado el ataque con el escudo!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(VictimaIndex, eSkill.Defensa)

            End If

        End If

    End If
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Function
    
ErrHandler:
    Dim AtacanteNick As String
    Dim VictimaNick  As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & _
            AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)

End Function
Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, _
                                    ByVal VictimaIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets
    '                    inválidos, y evitar un doble chequeo innecesario
    '***************************************************

    On Error GoTo ErrHandler
    
    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    
    With UserList(AtacanteIndex)

        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
            Call WriteLocaleMsg(AtacanteIndex, 8)
            Exit Function
        End If
        
 
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
        
            Select Case UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
            Case Is > 0
                    If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil = 1 And UserList(AtacanteIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).Snd1, .Pos.X, .Pos.Y))
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).Snd2, 0))
                    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO3, .Pos.X, .Pos.Y))
                    ElseIf ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil = 2 Then
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(68, .Pos.X, .Pos.Y))
                    Else
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
                    End If
            Case Else
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            End Select
         
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                    Call UserDañoUser(AtacanteIndex, VictimaIndex)

                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList( _
                        VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If


            Call UserIncinera(AtacanteIndex, VictimaIndex)
            
            If .Clase = eClass.ladron Or .Clase = eClass.Gladiador Or .Clase = eClass.Bardo Then
                Call Desarmar(AtacanteIndex, VictimaIndex)
                Call DoDesequipar(AtacanteIndex, VictimaIndex)
            End If
            
            Call GolpeParalizaUsuario(AtacanteIndex, VictimaIndex)
            Call UserDañoUser(AtacanteIndex, VictimaIndex)
            
        Else

            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                
                Select Case UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
                    Case Is > 0
                            If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil = 1 And UserList(AtacanteIndex).Invent.MunicionEqpObjIndex > 0 Then
                            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(145, .Pos.X, .Pos.Y))
                            ElseIf ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil = 2 Then
                            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(67, .Pos.X, .Pos.Y))
                            Else
                            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
                            End If
                    Case Else
                    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
                End Select
                
            Call WriteLocaleMsg(VictimaIndex, 460, "~" & UserList(AtacanteIndex).Char.CharIndex)
            Call WriteChatOverHeadLocale(AtacanteIndex, UserList(AtacanteIndex).Char.CharIndex, 0, 2) 'Fallas
            
            End If
        End If

    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.description)

End Function
Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar
    '11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal
    '***************************************************
    
    On Error GoTo ErrHandler

    Dim daño       As Long
    Dim lugar      As Byte
    Dim absorbido  As Long
    Dim defbarco   As Integer
    Dim defmontura As Integer
    Dim Obj        As ObjData
    Dim Resist     As Byte
    
    daño = CalcularDaño(AtacanteIndex)
           

    Call UserEnvenena(AtacanteIndex, VictimaIndex)
            Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

    
    With UserList(AtacanteIndex)

        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)

        End If
        
        If .flags.Montando = 1 Then
             Obj = ObjData(.Invent.MonturaObjIndex)
             daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
        End If
        
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)

        End If
        
        If UserList(VictimaIndex).flags.Montando = 1 And UserList(VictimaIndex).Invent.MonturaObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.MonturaObjIndex)
             defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        
        lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case lugar

            Case PartesCuerpo.bCabeza

                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco + defmontura - Resist
                    daño = daño - absorbido

                    If daño < 0 Then daño = 1

                End If
            
            Case Else

                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData

                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)

                    End If

                    absorbido = absorbido + defbarco + defmontura - Resist
                    daño = daño - absorbido

                    If daño < 0 Then daño = 1

                End If

        End Select
        
        'Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, lugar, _
                daño)
      '  Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, lugar, daño)
        
       Call WriteChatOverHeadLocale(AtacanteIndex, UserList(AtacanteIndex).Char.CharIndex, daño, 2)
       Call WriteChatOverHeadLocale(VictimaIndex, UserList(VictimaIndex).Char.CharIndex, daño, 2)
        
        UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
             If .Invent.NudiEqpObjIndex > 0 Then
                    Call SubirSkill(AtacanteIndex, eSkill.Wrestling)
            ElseIf .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles)
                ElseIf ObjData(.Invent.WeaponEqpObjIndex).proyectil = 2 Then
                    'es un Arpon. Sube ArmasArrojadizas
                   Call SubirSkill(AtacanteIndex, eSkill.ArmasArrojadizas)


                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, eSkill.armas)

                End If

            Else
                'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, eSkill.Wrestling)

            End If
                    
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)

            End If

        End If
        
        If UserList(VictimaIndex).Stats.MinHP <= 0 Then
        
                Call ContarMuerte(VictimaIndex, AtacanteIndex)
 
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer

            For j = 1 To MAXMASCOTAS

                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))

                    End If

                End If

            Next j
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex)
        Else
        
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)

        End If

    End With
    
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Sub
    
ErrHandler:
    Dim AtacanteNick As String
    Dim VictimaNick  As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDañoUser. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & _
            AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 05/05/2010
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
    '***************************************************

    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal       As Boolean
    Dim VictimaEsAtacable As Boolean
    
    With UserList(VictimIndex)

        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(VictimIndex)
            Call WriteLocaleMsg(VictimIndex, 123)
            .Char.FX = 0
            .Char.Loops = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageEfectoCharParticula(UserList(VictimIndex).Char.CharIndex, ParticleToLevel(VictimIndex), 0, True, True))
        End If

    End With
    
    UserList(VictimIndex).Counters.TiempoDeMapeo = 7
    UserList(attackerIndex).Counters.TiempoDeMapeo = 7
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)

End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS

        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1

        End If

    Next iCount

End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, _
                            ByVal VictimIndex As Integer) As Boolean

    '***************************************************
    'Autor: Unknown
    'Last Modification: 02/04/2010
    'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
    '24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
    '24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
    '02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
    '***************************************************
    On Error GoTo ErrHandler

    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(attackerIndex, 77)
        PuedeAtacar = False
        Exit Function

    End If
    
        If MapInfo(UserList(attackerIndex).Pos.Map).Pk = False Then  ' Or UserList(AttackerIndex).Pos.map = 1 Or UserList(AttackerIndex).Pos.map = 34 _
        Or UserList(AttackerIndex).Pos.map = 184 Or UserList(AttackerIndex).Pos.map = 183 Or UserList(AttackerIndex).Pos.map = 185 _
        Or UserList(AttackerIndex).Pos.map = 49 Or UserList(AttackerIndex).Pos.map = 194 Or UserList(AttackerIndex).Pos.map = 179 _
        Or UserList(AttackerIndex).Pos.map = 62 Or UserList(AttackerIndex).Pos.map = 64 Or UserList(AttackerIndex).Pos.map = 63 _
        Or UserList(AttackerIndex).Pos.map = 181 Or UserList(AttackerIndex).Pos.map = 180 Or UserList(AttackerIndex).Pos.map = 112 _
        Or UserList(AttackerIndex).Pos.map = 61 Or UserList(AttackerIndex).Pos.map = 183 Or UserList(AttackerIndex).Pos.map = 111 _
        Or UserList(AttackerIndex).Pos.map = 59 Or UserList(AttackerIndex).Pos.map = 183 Or UserList(AttackerIndex).Pos.map = 60 _
        Or UserList(AttackerIndex).Pos.map = 58 Or UserList(AttackerIndex).Pos.map = 183 Or UserList(AttackerIndex).Pos.map = 364 _
        Or UserList(AttackerIndex).Pos.map = 217 Or UserList(AttackerIndex).Pos.map = 183 Or UserList(AttackerIndex).Pos.map = 218 _
        Or UserList(AttackerIndex).Pos.map = 21 Or UserList(AttackerIndex).Pos.map = 208 Or UserList(AttackerIndex).Pos.map = 37 _
        Then

      Call WriteLocaleMsg(attackerIndex, 100)
      PuedeAtacar = False
      Exit Function
      
    End If
    
    If UserList(VictimIndex).Name = UserList(attackerIndex).Casamiento.Pareja Then
       Call WriteLocaleMsg(attackerIndex, 298)
       PuedeAtacar = False
       Exit Function
       
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(attackerIndex, 7)
        PuedeAtacar = False
        Exit Function

    End If
 
    
    
'Shermie80 no puede pueden atacar entre miembros de clan!
    If UserList(VictimIndex).GuildIndex > 0 Then
     If UserList(VictimIndex).GuildIndex = UserList(attackerIndex).GuildIndex Then
        'Call WriteMensajes(attackerIndex, eMensajes.Mensaje524)
        PuedeAtacar = False
        Exit Function
     End If
    End If
    'fin
    
 ' $ Nuevo Sistema de Facciones - Shermie80 $

    ' Republica + Republica
    If esRepu(VictimIndex) And esRepu(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
        
    End If
     
    ' Milicia + Milicia
    If esMili(VictimIndex) And esMili(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
        
    End If
    
    ' Republica + Milicia
    If esRepu(VictimIndex) And esMili(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function

    End If
    
    ' Milicia + Republica
    If esMili(VictimIndex) And esRepu(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function

    End If
    
    ' Ciudadano + Ciudadano
    If esCiuda(VictimIndex) And esCiuda(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
        
    End If
    
    ' Ciudadano + Armada Real
    If esCiuda(VictimIndex) And esArmada(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
        
    End If
    
    ' Armada Real + Ciudadano
    If esArmada(VictimIndex) And esCiuda(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
        
    End If

    ' Armada Real + Armada Real
    If esArmada(VictimIndex) And esArmada(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
        
    End If
    

  ' $ Fin $
  
    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(attackerIndex, VictimIndex)

        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE

            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then
                Call WriteLocaleMsg(attackerIndex, 101)
                PuedeAtacar = False
                Exit Function
               End If
            End If

    End Select

    'Sos un Armada atacando un ciudadano?
    If (esCiuda(VictimIndex)) And (esArmada(attackerIndex)) Or (esArmada(attackerIndex) And esArmada(VictimIndex)) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Sos un Mili atacando otro caos?
    If esRepu(VictimIndex) And esMili(attackerIndex) Then
        Call WriteLocaleMsg(attackerIndex, 97)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
  '  If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).Trigger = _
  '          eTrigger.ZONASEGURA Or MapData(UserList(AttackerIndex).Pos.Map, UserList(AttackerIndex).Pos.X, UserList( _
  '          AttackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
  '      Call WriteConsoleMsg( AttackerIndex, "No puedes pelear aquí.", FontTypeNames.FONTTYPE_INFO)
  '      PuedeAtacar = False
  '      Exit Function

 '   End If

    PuedeAtacar = True
    Exit Function

ErrHandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.description)

End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, _
                               ByVal npcindex As Integer, _
                               Optional ByVal Paraliza As Boolean = False, _
                               Optional ByVal IsPet As Boolean = False) As Boolean
    '***************************************************
    'Autor: Unknown Author (Original version)
    'Returns True if AttackerIndex can attack the NpcIndex
    'Last Modification: 16/11/2009
    '24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
    '14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
    'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
    '16/11/2009: ZaMa - Agrego validacion de pertenencia de npc.
    '02/04/2010: ZaMa - Los armadas ya no peuden atacar npcs no hotiles.
    '***************************************************
    
    Dim OwnerUserIndex As Integer
    
    'Estas muerto?
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(attackerIndex, 77)
        Exit Function

    End If
    
    'Sos consejero?
    If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
        'No pueden atacar NPC los Consejeros.
        Exit Function

    End If
    
    'Estas en modo Combate?
    If Not UserList(attackerIndex).flags.ModoCombate Then
        Call WriteLocaleMsg(attackerIndex, 102)
        PuedeAtacarNPC = False
        Exit Function
    End If
   
    
    'Es una criatura atacable?
    If Npclist(npcindex).Attackable = 0 Then
        Call WriteLocaleMsg(attackerIndex, 298)
        Exit Function

    End If
    
    'Es valida la distancia a la cual estamos atacando?
    If Distancia(UserList(attackerIndex).Pos, Npclist(npcindex).Pos) >= MAXDISTANCIAARCO Then
        Call WriteLocaleMsg(attackerIndex, 8)
        Exit Function

    End If
    
            'No era un Guardia, asi que es una criatura No-Hostil común.
            'Para asegurarnos que no sea una Mascota:
       If Npclist(npcindex).MaestroUser = 0 Then

            'Si sos ciudadano tenes que quitar el seguro para atacarla.
            If Not esRene(attackerIndex) Then
                
                ' Si sos armada no podes atacarlo directamente
                If esArmada(attackerIndex) Then
                    Call WriteLocaleMsg(attackerIndex, 272)
                    Exit Function

                End If
 
            End If

        End If

        
    'Es el NPC mascota de alguien?
    If Npclist(npcindex).MaestroUser > 0 Then
        If Not esRene(Npclist(npcindex).MaestroUser) Then
        
            'Es mascota de un Ciudadano.
            If esArmada(attackerIndex) Then
                'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                'Call WriteMensajes(attackerIndex, eMensajes.Mensaje370)
                Exit Function

            End If
            
            If Not esRene(attackerIndex) Then
                
                'El atacante no tiene el seguro puesto. Recibe penalización.
                'Call WriteMensajes(attackerIndex, eMensajes.Mensaje372)
                PuedeAtacarNPC = True
                Exit Function
                
            End If

        Else

            'Es mascota de un Criminal.
            If esCaos(Npclist(npcindex).MaestroUser) Then

                'Es Caos el Dueño.
                If esCaos(attackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteLocaleMsg(attackerIndex, 272)
                    Exit Function

                End If

            End If

        End If

    End If
    
    With Npclist(npcindex)
        ' El npc le pertenece a alguien?
        OwnerUserIndex = .Owner
        
        If OwnerUserIndex > 0 Then
            
            ' Puede atacar a su propia criatura!
            If OwnerUserIndex = attackerIndex Then
                PuedeAtacarNPC = True
                Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
                Exit Function

            End If

            
            ' Si son del mismo clan o party, pueden atacar (No renueva el timer)
            If Not SameClan(OwnerUserIndex, attackerIndex) Then
            
                ' Si se le agoto el tiempo
                If IntervaloPerdioNpc(OwnerUserIndex) Then ' Se lo roba :P
                    Call PerdioNpc(OwnerUserIndex)
                    Call ApropioNpc(attackerIndex, npcindex)
                    PuedeAtacarNPC = True
                    Exit Function
                    
                    ' Si lanzo un hechizo de para o inmo
                ElseIf Paraliza Then
                
                    ' Si ya esta paralizado o inmovilizado, no puedo inmovilizarlo de nuevo
                    If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
                        
                        'TODO_ZAMA: Si dejo esto asi, los pks con seguro peusto van a poder inmobilizar criaturas con dueño
                        ' Si es pk neutral, puede hacer lo que quiera :P.
                        If Not esRene(attackerIndex) And Not esRene(OwnerUserIndex) Then
                        
                            'El atacante es Armada
                            If esArmada(attackerIndex) Then
                                
                                'Intententa paralizar un npc de un armada?
                                If esArmada(OwnerUserIndex) Then
                                    'El atacante es Armada y esta intentando paralizar un npc de un armada: No puede
                                    Call WriteLocaleMsg(attackerIndex, 272)
                                    Exit Function
                                
                                    'El atacante es Armada y esta intentando paralizar un npc de un ciuda
 
                                End If
 
                            End If
                            
                            ' Al menos uno de los dos es criminal
                        Else

                            ' Si ambos son caos
                            If esCaos(attackerIndex) And esCaos(OwnerUserIndex) Then
                                'El atacante es Caos y esta intentando paralizar un npc de un Caos
                            'Call WriteMensajes(attackerIndex, eMensajes.Mensaje373)
                                Exit Function

                            End If

                        End If
                    
                        ' El npc no esta inmobilizado ni paralizado
                    Else

                        ' Si no tiene dueño, puede apropiarselo
                        If OwnerUserIndex = 0 Then

                            ' Siempre que no posea uno ya (el inmo/para no cambia pertenencia de npcs).
                            If UserList(attackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(attackerIndex, npcindex)

                            End If

                        End If
                        
                        ' Siempre se pueden paralizar/inmobilizar npcs con o sin dueño
                        ' que no tengan ese estado
                        PuedeAtacarNPC = True
                        Exit Function

                    End If
                    
                    ' No lanzó hechizos inmobilizantes
                Else
                    
                    ' El npc le pertenece a un ciudadano
                    If Not esRene(OwnerUserIndex) Then
                        
                        'El atacante es Armada y esta intentando atacar un npc de un Ciudadano
                        If esArmada(attackerIndex) Then
                        
                            'Intententa atacar un npc de un armada?
                            If esArmada(OwnerUserIndex) Then
                                'El atacante es Armada y esta intentando atacar el npc de un armada: No puede
                                'Call WriteMensajes(, eMensajes.Mensaje374)
                                Call WriteLocaleMsg(attackerIndex, 272)
                                Exit Function
   
                            End If
                            
                            ' No es aramda, puede ser criminal o ciuda
                        Else
                            
                            'El atacante es Ciudadano y esta intentando atacar un npc de un Ciudadano.
                            If Not esRene(attackerIndex) Then
                                
                                'El atacante es criminal y esta intentando atacar un npc de un Ciudadano.
                            Else
                                
                                PuedeAtacarNPC = True

                            End If

                        End If
                        
                        ' Es npc de un criminal
                    Else

                        If esCaos(OwnerUserIndex) Then

                            'Es Caos el Dueño.
                            If esCaos(attackerIndex) Then
                                'Un Caos intenta atacar una npc de un Caos. No puede atacar.
                                'Call WriteMensajes(attackerIndex, eMensajes.Mensaje373)
                                Exit Function

                            End If

                        End If

                    End If

                End If

            End If
            
            ' Si no tiene dueño el npc, se lo apropia
        Else

            ' Solo pueden apropiarse de npcs los caos, armadas o ciudas.
            If Not esRene(attackerIndex) Or esCaos(attackerIndex) Then

                    ' Si es una mascota atacando, no se apropia del npc
                    If Not IsPet Then

                        ' No es dueño de ningun npc => Se lo apropia.
                        If UserList(attackerIndex).flags.OwnedNpc = 0 Then
                            Call ApropioNpc(attackerIndex, npcindex)
                            ' Es dueño de un npc, pero no puede ser de este porque no tiene propietario.
                        Else

                            ' Se va a adueñar del npc (y perder el otro) solo si no inmobiliza/paraliza
                            If Not Paraliza Then Call ApropioNpc(attackerIndex, npcindex)

                        End If

                    End If

            End If

        End If

    End With
 
    
    PuedeAtacarNPC = True

End Function

Private Function SameClan(ByVal UserIndex As Integer, _
                          ByVal OtherUserIndex As Integer) As Boolean
    '***************************************************
    'Autor: ZaMa
    'Returns True if both players belong to the same clan.
    'Last Modification: 16/11/2009
    '***************************************************
    SameClan = (UserList(UserIndex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And UserList( _
            UserIndex).GuildIndex <> 0

End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal npcindex As Integer, ByVal ElDaño As Long)
    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/09/06 Nacho
    'Reescribi gran parte del Sub
    'Ahora, da toda la experiencia del npc mientras este vivo.
    '***************************************************
    Dim ExpaDar As Long
    
    Dim ExpDonador As Long
    
    If ElDaño <= 0 Then ElDaño = 0
    
    If Npclist(npcindex).Stats.MaxHP <= 0 Then Exit Sub
    
    If ElDaño > Npclist(npcindex).Stats.MinHP Then ElDaño = Npclist(npcindex).Stats.MinHP
    
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDaño * (Npclist(npcindex).GiveEXP / Npclist(npcindex).Stats.MaxHP))

    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
    'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
    'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(npcindex).flags.ExpCount Then
        ExpaDar = Npclist(npcindex).flags.ExpCount
        Npclist(npcindex).flags.ExpCount = 0
    Else
        Npclist(npcindex).flags.ExpCount = Npclist(npcindex).flags.ExpCount - ExpaDar

    End If
    
    If ExpModificada = True Then
        ExpaDar = ExpaDar * multiplicadorExp
    End If
    
    If UserList(UserIndex).Donador.activo > 0 Then
        ExpaDar = ExpaDar + (ExpaDar / 2)
    End If
    
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
        If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
            
            Call WriteUpdateExp(UserIndex)
            Call CheckUserLevel(UserIndex)
            Call WriteLocaleMsg(UserIndex, 140, ExpaDar)
        End If
    End If
    
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, _
                                 ByVal Destino As Integer) As eTrigger6
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'TODO: Pero que rebuscado!!
    'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
    On Error GoTo ErrHandler

    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).Trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).Trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE

        End If

    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE

    End If

    Exit Function
ErrHandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)

End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex

        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "¡¡" & UserList(AtacanteIndex).Name & " te ha envenenado!!", _
                            FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "¡¡Has envenenado a " & UserList(VictimaIndex).Name & "!!", _
                            FontTypeNames.FONTTYPE_FIGHT)

                End If

            End If

        End If

    End If
    
    Call FlushBuffer(VictimaIndex)

End Sub
Private Function PoderAtaqueArpon(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        If .Stats.UserSkills(eSkill.ArmasArrojadizas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.ArmasArrojadizas) * ModClase(.Clase).AtaqueArpon
        ElseIf .Stats.UserSkills(eSkill.ArmasArrojadizas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.ArmasArrojadizas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArpon
        ElseIf .Stats.UserSkills(eSkill.ArmasArrojadizas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.ArmasArrojadizas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArpon
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.ArmasArrojadizas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArpon
        End If
        
        PoderAtaqueArpon = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueNudi(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)

        If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.Clase).AtaqueWrestling
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueWrestling
        ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueWrestling
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
                    ModClase(.Clase).AtaqueWrestling

        End If
        
        PoderAtaqueNudi = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Sub UserIncinera(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Orbe As Boolean
    
    If UserList(VictimaIndex).flags.Incinerado = 1 Then Exit Sub
    
    'If UserList(AtacanteIndex).Invent.ItemsMagicosEqpObjIndex = 868 Then Orbe = True
    
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    If (ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil > 0 And ObjData(UserList(AtacanteIndex).Invent.MunicionEqpObjIndex).Snd3 > 0) Then Orbe = True
    End If
    
    If Orbe = False Then Exit Sub
    
    If Orbe = True Then
     If RandomNumber(1, 35) <= 5 Then
      UserList(VictimaIndex).flags.Incinerado = 1
      UserList(VictimaIndex).Counters.Incinerado = 150
      Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 8, 0))
      Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 123, 0))
      Call WriteLocaleMsg(VictimaIndex, 48)
     End If
    End If

   Call FlushBuffer(VictimaIndex)

End Sub
Public Sub GolpeParalizaNPC(ByVal UserIndex As Integer, ByVal npcindex As Integer)
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************
    Dim Orbe As Boolean
    Dim Probabilidad   As Integer
    Dim Resultado      As Integer
    Dim WrestlingSkill As Byte
    
    If Npclist(npcindex).flags.Paralizado = 1 Or Npclist(npcindex).flags.Inmovilizado = 1 Then Exit Sub

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    If UserList(UserIndex).Invent.MunicionEqpObjIndex = 1085 And ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 0 Then Orbe = True
    End If
 
    With UserList(UserIndex)
    If Orbe = False And UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 500)
        
       If UserList(UserIndex).Clase = eClass.Gladiador Or UserList(UserIndex).Clase = eClass.Bardo Then
        Resultado = RandomNumber(1, 300)
       End If
        
           If Resultado <= Probabilidad Then
 
           Npclist(npcindex).flags.Paralizado = 1
           Npclist(npcindex).Contadores.Paralisis = 30
           Call WriteConsoleMsg(UserIndex, "Tu golpe ha paralizado a la criatura.", FontTypeNames.FONTTYPE_INFO)
           Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(17, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
           Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessageCreateFX(Npclist(npcindex).Char.CharIndex, 8, 0))
           Else
           End If
    ElseIf Orbe = True Then
        If RandomNumber(1, 35) <= 5 Then
           Npclist(npcindex).flags.Paralizado = 1
           Npclist(npcindex).Contadores.Paralisis = 30
           Call WriteConsoleMsg(UserIndex, "Tu golpe ha paralizado a la criatura.", FontTypeNames.FONTTYPE_INFO)
           Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessagePlayWave(17, Npclist(npcindex).Pos.X, Npclist(npcindex).Pos.Y))
           Call SendData(SendTarget.ToNPCArea, npcindex, PrepareMessageCreateFX(Npclist(npcindex).Char.CharIndex, 8, 0))
           Else
        End If
    End If
    End With
    
End Sub
Public Sub GolpeParalizaUsuario(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)


    Dim Orbe As Boolean

    If UserList(VictimaIndex).flags.Inmovilizado = 1 Or UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    
    'If UserList(UserIndex).Invent.ItemsMagicosEqpObjIndex = 869 Then Orbe = True
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    If (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil > 0 And ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).Paraliza > 0) Then Orbe = True
    End If
 
    Dim Probabilidad   As Integer
    Dim Resultado      As Integer
    Dim WrestlingSkill As Byte
    
    
    
    With UserList(UserIndex)
    
    If Orbe = False And UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 500)
        
       If UserList(UserIndex).Clase = eClass.Gladiador Or UserList(UserIndex).Clase = eClass.Bardo Then
        Resultado = RandomNumber(1, 300)
       End If
       
        
      If Resultado <= Probabilidad Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = 150
        
        UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).Name

        Call WriteLocaleMsg(UserIndex, 135, UserList(VictimaIndex).Name)
        Call WriteLocaleMsg(VictimaIndex, 134, UserList(UserIndex).Name)

        
        Call WriteParalizeOK(VictimaIndex)
        
        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(17, .Pos.X, .Pos.Y))
        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 8, 0))
        Call FlushBuffer(VictimaIndex)
      End If
      
    ElseIf Orbe = True Then
        If RandomNumber(1, 35) <= 5 Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = 150
        
        
        UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).Name
        
        
        Call WriteLocaleMsg(UserIndex, 135, UserList(VictimaIndex).Name)
        Call WriteLocaleMsg(VictimaIndex, 134, UserList(UserIndex).Name)

        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(17, .Pos.X, .Pos.Y))
        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 8, 0))
        Call FlushBuffer(VictimaIndex)
        End If
    End If

     
    End With
    
End Sub

 


