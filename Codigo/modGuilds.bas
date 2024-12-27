Attribute VB_Name = "modGuilds"
'**************************************************************
' LinkAO v1.0
' modGuilds.bas
'
'**************************************************************
Option Explicit

Private GUILDINFOFILE             As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS          As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES           As Integer
'cantidad actual de clanes en el servidor

Private guilds(1 To MAX_GUILDS)   As clsClan
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES        As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION      As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

'Gemas para fundar clan
Private Const GEMA_LUNAR As Integer = 406
Private Const GEMA_DORADA As Integer = 410
Private Const GEMA_NARANJA As Integer = 408
Private Const GEMA_GRIS As Integer = 409

Public Enum ALINEACION_GUILD
    ALINEACION_REPUBLICANO = 1
    ALINEACION_IMPERIAL = 2
    ALINEACION_CAOTICO = 3
    ALINEACION_RENEGADO = 4
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD

    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45

End Enum

'numero de .wav del cliente

Public Enum RELACIONES_GUILD

    GUERRA = -1
    PAZ = 0
    ALIADOS = 1

End Enum

'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo LoadGuildsDB_Err
     If frmMain.Visible Then frmMain.AgregarConsola "Cargando guildsinfo.inf."
    Dim CantClanes As String
    Dim i          As Integer
    Dim TempStr    As String
    Dim Alin       As ALINEACION_GUILD
    
100    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

102    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
104    If IsNumeric(CantClanes) Then
106        CANTIDADDECLANES = CInt(CantClanes)
    Else
108        CANTIDADDECLANES = 0

    End If
    
110    For i = 1 To CANTIDADDECLANES
112        Set guilds(i) = New clsClan
114        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
116        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
118        Call guilds(i).Inicializar(TempStr, i, Alin)
120    Next i
      If frmMain.Visible Then frmMain.AgregarConsola "Se cargo el archivo guildsinfo.inf."
    Exit Sub

LoadGuildsDB_Err:
122     Call RegistrarError(Err.Number, Err.description, "modGuilds.LoadGuildsDB", Erl)
124     Resume Next
        
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, _
                                       ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NuevaA As Boolean
    Dim News   As String

    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(UserIndex, True, NuevaA)

        If NuevaA Then News = News & "El clan tiene nueva alineación."

        'If NuevoL Or NuevaA Then Call guilds(GuildIndex).SetGuildNews(News)
    End If

End Function

Public Function m_ValidarPermanencia(ByVal UserIndex As Integer, _
                                     ByVal SumaAntifaccion As Boolean, _
                                     ByRef CambioAlineacion As Boolean) As Boolean
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 14/12/2009
    '25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
    '14/12/2009: ZaMa - La alineacion del clan depende del lider
    '14/02/2010: ZaMa - Ya no es necesario saber si el lider cambia, ya que no puede cambiar.
    '***************************************************

    Dim GuildIndex As Integer

    m_ValidarPermanencia = True
    
    GuildIndex = UserList(UserIndex).GuildIndex

    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        
        ' Es el lider, bajamos 1 rango de alineacion
        If GuildLeader(GuildIndex) = UserList(UserIndex).Name Then
            Call LogClanes(UserList(UserIndex).Name & ", líder de " & guilds(GuildIndex).GuildName & _
                    " hizo bajar la alienación de su clan.")
        
            CambioAlineacion = True
            
            ' Por si paso de ser armada/legion a pk/ciuda, chequeo de nuevo
            Do
                Call UpdateGuildMembers(GuildIndex)
            Loop Until m_EstadoPermiteEntrar(UserIndex, GuildIndex)

        Else
            Call LogClanes(UserList(UserIndex).Name & " de " & guilds(GuildIndex).GuildName & _
                    " es expulsado en validar permanencia.")
        
            m_ValidarPermanencia = False

            If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
            
            CambioAlineacion = guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            
            Call LogClanes(UserList(UserIndex).Name & " de " & guilds(GuildIndex).GuildName & IIf(CambioAlineacion, _
                    " SI ", " NO ") & "provoca cambio de alineación. MAXANT:" & CambioAlineacion)
            
            Call m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
            
            ' Llegamos a la maxima cantidad de antifacciones permitidas, bajamos un grado de alineación
            If CambioAlineacion Then
                Call UpdateGuildMembers(GuildIndex)

            End If

        End If

    End If

End Function

Private Sub UpdateGuildMembers(ByVal GuildIndex As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Pulo detalles en el funcionamiento general.
    '***************************************************
    Dim GuildMembers() As String
    Dim TotalMembers   As Integer
    Dim MemberIndex    As Long
    Dim Sale           As Boolean
    Dim MemberName     As String
    Dim UserIndex      As Integer
    Dim Reenlistadas   As Integer
    
    ' Si devuelve true, cambio a neutro y echamos a todos los que estén de mas, sino no echamos a nadie
    If guilds(GuildIndex).CambiarAlineacion(BajarGrado(GuildIndex)) Then 'ALINEACION_NEUTRO)
        
        'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
        'internamente al iterador en el proceso
        GuildMembers = guilds(GuildIndex).GetMemberList()
        TotalMembers = UBound(GuildMembers)
        
        For MemberIndex = 0 To TotalMembers
            MemberName = GuildMembers(MemberIndex)
            
            'vamos a violar un poco de capas..
            UserIndex = NameIndex(MemberName)

            If UserIndex > 0 Then
                Sale = Not m_EstadoPermiteEntrar(UserIndex, GuildIndex)
            Else
                Sale = Not m_EstadoPermiteEntrarChar(MemberName, GuildIndex)

            End If

            If Sale Then
                If m_EsGuildLeader(MemberName, GuildIndex) Then  'hay que sacarlo de las facciones
                 
                    If UserIndex > 0 Then
                        If esArmada(UserIndex) <> 0 Then
                            'Call ExpulsarFaccionReal(UserIndex)
                            ' No cuenta como reenlistada :p.
                            UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango - 1
                        ElseIf esCaos(UserIndex) <> 0 Then
                            'Call ExpulsarFaccionCaos(UserIndex)
                            ' No cuenta como reenlistada :p.
                            UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango - 1

                        End If
                    End If

                Else    'sale si no es guildLeader
                    Call m_EcharMiembroDeClan(-1, MemberName)

                End If

            End If

        Next MemberIndex

    Else
        ' Resetea los puntos de antifacción
        guilds(GuildIndex).PuntosAntifaccion = 0

    End If

End Sub

Private Function BajarGrado(ByVal GuildIndex As Integer) As ALINEACION_GUILD
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 27/11/2009
    'Reduce el grado de la alineacion a partir de la alineacion dada
    '***************************************************

    Select Case guilds(GuildIndex).Alineacion

    ' '  Case ALINEACION_ARMADA
      '      BajarGrado = ALINEACION_CIUDA'
'
 '       Case ALINEACION_LEGION
  '          BajarGrado = ALINEACION_CRIMINAL
   '
    '    Case Else
     '       BajarGrado = ALINEACION_NEUTRO
'
    End Select

End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, _
                                       ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(UserIndex)

End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, _
                                 ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))

End Function

Private Function m_EsGuildFounder(ByRef PJ As String, _
                                  ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))

End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, _
                                     ByVal Expulsado As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'UI echa a Expulsado del clan de Expulsado
    Dim UserIndex As Integer
    Dim Gi        As Integer
    
    m_EcharMiembroDeClan = 0

    UserIndex = NameIndex(Expulsado)

    If UserIndex > 0 Then
        'pj online
        Gi = UserList(UserIndex).GuildIndex

        If Gi > 0 Then
            If m_PuedeSalirDeClan(Expulsado, Gi, Expulsador) Then
                Call guilds(Gi).DesConectarMiembro(UserIndex)
                Call guilds(Gi).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(Gi).GuildName & " Expulsador = " & _
                        Expulsador)
                UserList(UserIndex).GuildIndex = 0
                Call RefreshCharStatus(UserIndex)
                m_EcharMiembroDeClan = Gi
            Else
                m_EcharMiembroDeClan = 0

            End If

        Else
            m_EcharMiembroDeClan = 0

        End If

    Else
        'pj offline
        Gi = GetGuildIndexFromChar(Expulsado)

        If Gi > 0 Then
            If m_PuedeSalirDeClan(Expulsado, Gi, Expulsador) Then
                Call guilds(Gi).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(Gi).GuildName & " Expulsador = " & _
                        Expulsador)
                m_EcharMiembroDeClan = Gi
            Else
                m_EcharMiembroDeClan = 0

            End If

        Else
            m_EcharMiembroDeClan = 0

        End If

    End If

End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi As Integer

    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then Exit Sub
    
    Call guilds(Gi).SetURL(Web)
    
End Sub

Public Sub ChangeCodexAndDesc(ByRef desc As String, _
                              ByRef codex() As String, _
                              ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Long
    
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    
    With guilds(GuildIndex)
        Call .SetDesc(desc)
        
        For i = 0 To UBound(codex())
            Call .SetCodex(i, codex(i))
        Next i
        
        For i = i To CANTIDADMAXIMACODEX
            Call .SetCodex(i, vbNullString)
        Next i

    End With

End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 21/02/2010
    '21/02/2010: ZaMa - Ahora le avisa a los miembros que cambio el guildnews.
    '***************************************************

    Dim Gi As Integer

    With UserList(UserIndex)
        Gi = .GuildIndex
        
        If Gi <= 0 Or Gi > CANTIDADDECLANES Then Exit Sub
        
        If Not m_EsGuildLeader(.Name, Gi) Then Exit Sub
        
        Call guilds(Gi).SetGuildNews(Datos)
        
        Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & _
                " ha actualizado las noticias del clan!"))

    End With

End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, _
                               ByRef desc As String, _
                               ByRef GuildName As String, _
                               ByRef URL As String, _
                               ByRef codex() As String, _
                               ByVal Alineacion As ALINEACION_GUILD, _
                               ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim CantCodex   As Integer
    Dim i           As Integer
    Dim DummyString As String

    CrearNuevoClan = False

    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function

    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
        Exit Function

    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function

    End If

    CantCodex = UBound(codex()) + 1

    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

        'constructor custom de la clase clan
        Set guilds(CANTIDADDECLANES) = New clsClan
        
        With guilds(CANTIDADDECLANES)
            Call .Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
            
            'Damos de alta al clan como nuevo inicializando sus archivos
            Call .InicializarNuevoClan(UserList(FundadorIndex).Name)
            
            'seteamos codex y descripcion
            For i = 1 To CantCodex
                Call .SetCodex(i, codex(i - 1))
            Next i

            Call .SetDesc(desc)
            Call .SetGuildNews("Clan creado con alineación: " & Alineacion2String(Alineacion))
            Call .SetLeader(UserList(FundadorIndex).Name)
            Call .SetURL(URL)
            
            '"conectamos" al nuevo miembro a la lista de la clase
            Call .AceptarNuevoMiembro(UserList(FundadorIndex).Name)
            Call .ConectarMiembro(FundadorIndex)
           
        'Shermie80
        For i = 1 To MAX_INVENTORY_SLOTS
            
            If UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_LUNAR Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_DORADA Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_NARANJA Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_GRIS Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            End If
            
        Next i

        End With
        
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        
        For i = 1 To CANTIDADDECLANES - 1
            Call guilds(i).ProcesarFundacionDeOtroClan
        Next i

    Else
        refError = "No hay más slots para fundar clanes. Consulte a un administrador."
        Exit Function

    End If
    
    CrearNuevoClan = True

End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildIndex As Integer
    Dim i          As Integer
    Dim go         As Integer

    GuildIndex = UserList(UserIndex).GuildIndex

    If GuildIndex = 0 Then Exit Sub

    Dim enemies() As String
    
    With guilds(GuildIndex)

        If .CantidadEnemys Then
            ReDim enemies(0 To .CantidadEnemys - 1) As String
        Else
            ReDim enemies(0)

        End If
        
        Dim allies() As String
        
        If .CantidadAllies Then
            ReDim allies(0 To .CantidadAllies - 1) As String
        Else
            ReDim allies(0)

        End If
        
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = 0
        
        While i > 0

            enemies(go) = guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
            go = go + 1
        Wend
        
        i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
        go = 0
        
        While i > 0

            allies(go) = guilds(i).GuildName
            i = .Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
        Wend
    
        Call WriteGuildNews(UserIndex, .GetGuildNews, enemies, allies)
    
        If .EleccionesAbiertas Then
            Call WriteConsoleMsg(UserIndex, "Hoy es la votación para elegir un nuevo líder para el clan.", _
                    FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(UserIndex, _
                    "La elección durará 24 horas, se puede votar a cualquier miembro del clan.", _
                    FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(UserIndex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(UserIndex, "Sólo se computará un voto por miembro. Tu voto no puede ser cambiado.", _
                    FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, _
                                   ByVal GuildIndex As Integer, _
                                   ByVal QuienLoEchaUI As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False

    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function

    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function

            End If

        End If

    End If

    ' Ahora el lider es el unico que no puede salir del clan
    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).GetLeader) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, _
                                  ByVal Alineacion As ALINEACION_GUILD, _
                                  ByRef refError As String) As Boolean
    '***************************************************
    'Autor: Unknown
    'Last Modification: 27/11/2009
    'Returns true if can Found a guild
    '27/11/2009: ZaMa - Ahora valida si ya fundo clan o no.
    '***************************************************
       ' PuedeFundarUnClan = True
               Call WriteConsoleMsg(UserIndex, "Clanes Deshabilitados temporalmente.", FontTypeNames.FONTTYPE_INFO)
               Exit Function
            

       If Not TieneObjetos(GEMA_LUNAR, 1, UserIndex) Then
       refError = "Necesitas una Gema Lunar para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    ElseIf TieneObjetos(GEMA_NARANJA, 1, UserIndex) = False Then
        refError = "Necesitas una Gema Naranja para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    ElseIf TieneObjetos(GEMA_DORADA, 1, UserIndex) = False Then
       refError = "Necesitas una Gema Dorada para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    ElseIf TieneObjetos(GEMA_GRIS, 1, UserIndex) = False Then
        refError = "Necesitas una Gema Gris para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    End If
        Exit Function
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function

    End If
    
    If UserList(UserIndex).Stats.ELV < 40 Or UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        refError = "Para fundar un clan debes ser nivel 40 y tener 90 skills en liderazgo."
        Exit Function

    End If
    
         
    If Not TieneObjetos(GEMA_LUNAR, 1, UserIndex) Then
       refError = "Necesitas una Gema Lunar para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    ElseIf TieneObjetos(GEMA_NARANJA, 1, UserIndex) = False Then
        refError = "Necesitas una Gema Naranja para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    ElseIf TieneObjetos(GEMA_DORADA, 1, UserIndex) = False Then
       refError = "Necesitas una Gema Dorada para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    ElseIf TieneObjetos(GEMA_GRIS, 1, UserIndex) = False Then
        refError = "Necesitas una Gema Gris para poder fundar un clan."
        PuedeFundarUnClan = False
        Exit Function
    End If
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, _
                                           ByVal GuildIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Promedio As Long
    Dim ELV      As Integer
    Dim f        As Byte
    Dim ff       As Byte
    
    m_EstadoPermiteEntrarChar = False
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)

    End If

    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)

    End If

    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)

    End If
    
    If FileExist(CharPath & Personaje & ".chr") Then

        Select Case guilds(GuildIndex).Alineacion

        Case ALINEACION_GUILD.ALINEACION_IMPERIAL
               f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "Ciudadano"))
               ff = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoReal"))
               
               If f = 1 Or ff = 1 Then
                  m_EstadoPermiteEntrarChar = True
               End If
            
        Case ALINEACION_GUILD.ALINEACION_CAOTICO
               f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoCaos"))
               
               If f = 1 Then
                  m_EstadoPermiteEntrarChar = True
               End If
             
            Case ALINEACION_GUILD.ALINEACION_REPUBLICANO
               f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "Republicano"))
               ff = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "EjercitoMili"))
               
               If f = 1 Or ff = 1 Then
                  m_EstadoPermiteEntrarChar = True
               End If
            
            Case ALINEACION_GUILD.ALINEACION_RENEGADO
               f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "Renegado"))
               
               If f = 1 Then
                  m_EstadoPermiteEntrarChar = True
               End If
            
            Case Else
                 m_EstadoPermiteEntrarChar = False

        End Select

    End If

End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, _
                                       ByVal GuildIndex As Integer) As Boolean

    If UCase$(UserList(UserIndex).Name) = UCase$(modGuilds.GuildLeader(GuildIndex)) Then
        m_EstadoPermiteEntrar = True
        Exit Function
    End If
    
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_IMPERIAL
            m_EstadoPermiteEntrar = esArmada(UserIndex) Or esCiuda(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_CAOTICO
            m_EstadoPermiteEntrar = esCaos(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_REPUBLICANO
            m_EstadoPermiteEntrar = esMili(UserIndex) Or esRepu(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_RENEGADO
            m_EstadoPermiteEntrar = esRene(UserIndex)
        Case Else
            m_EstadoPermiteEntrar = True
    End Select

End Function

Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    Select Case S
        Case "Coático"
            String2Alineacion = ALINEACION_CAOTICO
        Case "Armada Real"
            String2Alineacion = ALINEACION_IMPERIAL
        Case "Milicia"
            String2Alineacion = ALINEACION_REPUBLICANO
        Case "Renegado"
            String2Alineacion = ALINEACION_RENEGADO
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_CAOTICO
            Alineacion2String = "Caótico"
        Case ALINEACION_GUILD.ALINEACION_IMPERIAL
            Alineacion2String = "Imperial"
        Case ALINEACION_GUILD.ALINEACION_REPUBLICANO
            Alineacion2String = "Republicano"
        Case ALINEACION_GUILD.ALINEACION_RENEGADO
            Alineacion2String = "Renegado"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case Relacion

        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"

        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"

        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"

        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"

    End Select

End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Select Case UCase$(Trim$(S))

        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ

        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA

        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS

        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ

    End Select

End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte
    Dim i   As Integer

    'old function by morgo

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))

        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            GuildNameValido = False
            Exit Function

        End If
    
    Next i

    GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    YaExiste = False
    GuildName = UCase$(GuildName)

    For i = 1 To CANTIDADDECLANES
        YaExiste = (UCase$(guilds(i).GuildName) = GuildName)

        If YaExiste Then Exit Function
    Next i

End Function

Public Function HasFound(ByRef UserName As String) As Boolean
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 27/11/2009
    'Returns true if it's already the founder of other guild
    '***************************************************
    Dim i    As Long
    Dim Name As String

    Name = UCase$(UserName)

    For i = 1 To CANTIDADDECLANES
        HasFound = (UCase$(guilds(i).Fundador) = Name)

        If HasFound Then Exit Function
    Next i

End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, _
                                  ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildIndex As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tú no perteneces a ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
        refError = "No eres el líder de tu clan"
        Exit Function

    End If
    
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas."
        Exit Function

    End If
    
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, _
                              ByRef Votado As String, _
                              ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildIndex As Integer
    Dim list()     As String
    Dim i          As Long

    v_UsuarioVota = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tú no perteneces a ningún clan."
        Exit Function

    End If

    With guilds(GuildIndex)

        If Not .EleccionesAbiertas Then
            refError = "No hay elecciones abiertas en tu clan."
            Exit Function

        End If
        
        list = .GetMemberList()

        For i = 0 To UBound(list())

            If UCase$(Votado) = list(i) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            refError = Votado & " no pertenece al clan."
            Exit Function

        End If
        
        If .YaVoto(UserList(UserIndex).Name) Then
            refError = "Ya has votado, no puedes cambiar tu voto."
            Exit Function

        End If
        
        Call .ContabilizarVoto(UserList(UserIndex).Name, Votado)
        v_UsuarioVota = True

    End With

End Function

Public Sub v_RutinaElecciones()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    On Error GoTo errh

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", _
            FontTypeNames.FONTTYPE_SERVER))

    For i = 1 To CANTIDADDECLANES

        If Not guilds(i) Is Nothing Then
            If guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & _
                        " es el nuevo líder de " & guilds(i).GuildName & ".", FontTypeNames.FONTTYPE_SERVER))

            End If

        End If

proximo:
    Next i

  '  Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> ¡No se dejen engañar!, ningun miembro del staff pedirá jamas datos privados de los usuarios, todo jugador tiene la obligación de negar de este tipo de datos a cualquier persona.", _
            FontTypeNames.FONTTYPE_INFO))
    Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.description)

    Resume proximo

End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'aca si que vamos a violar las capas deliveradamente ya que
    'visual basic no permite declarar metodos de clase
    Dim Temps As String

    If InStrB(PlayerName, "\") <> 0 Then
        PlayerName = Replace(PlayerName, "\", vbNullString)

    End If

    If InStrB(PlayerName, "/") <> 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)

    End If

    If InStrB(PlayerName, ".") <> 0 Then
        PlayerName = Replace(PlayerName, ".", vbNullString)

    End If

    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")

    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0

    End If

End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'me da el indice del guildname
    Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)

    For i = 1 To CANTIDADDECLANES

        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function

        End If

    Next i

End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, _
                                        ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex

        While i > 0

            'No mostramos dioses y admins
            If i <> UserIndex And ((UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or _
                    PlayerType.SemiDios)) <> 0 Or (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or _
                    PlayerType.Admin) <> 0)) Then m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList( _
                    i).Name & ","
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend

    End If

    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)

    End If

End Function

Public Function PrepareGuildsList() As String()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tStr() As String
    Dim i      As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 1 To CANTIDADDECLANES
            tStr(i - 1) = guilds(i).GuildName
        Next i

    End If
    
    PrepareGuildsList = tStr

End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim codex(CANTIDADMAXIMACODEX - 1) As String
    Dim Gi                             As Integer
    Dim i                              As Long

    Gi = GuildIndex(GuildName)

    If Gi = 0 Then Exit Sub
    
    With guilds(Gi)

        For i = 1 To CANTIDADMAXIMACODEX
            codex(i - 1) = .GetCodex(i)
        Next i
        
        Call Protocol.WriteGuildDetails(UserIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, .GetURL, _
                .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), .CantidadEnemys, _
                .CantidadAllies, .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), codex, .GetDesc)

    End With

End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Mariano Barrou (El Oso)
    'Last Modification: 12/10/06
    'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    '***************************************************
    Dim Gi              As Integer
    Dim guildList()     As String
    Dim MemberList()    As String
    Dim aspirantsList() As String

    With UserList(UserIndex)
        Gi = .GuildIndex
        
        guildList = PrepareGuildsList()
        
        If Gi <= 0 Or Gi > CANTIDADDECLANES Then
            'Send the guild list instead
            Call WriteGuildList(UserIndex, guildList)
            Exit Sub

        End If
        
        MemberList = guilds(Gi).GetMemberList()
        
        If Not m_EsGuildLeader(.Name, Gi) Then
            'Send the guild list instead
            Call WriteGuildMemberInfo(UserIndex, guildList, MemberList)
            Exit Sub

        End If
        
        aspirantsList = guilds(Gi).GetAspirantes()
        
        Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(Gi).GetGuildNews(), aspirantsList)

    End With

End Sub

Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0

    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()

    End If

End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0

    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()

    End If

End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, _
                                            ByVal Tipo As RELACIONES_GUILD) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0

    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)

    End If

End Function

Public Function GMEscuchaClan(ByVal UserIndex As Integer, _
                              ByVal GuildName As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, _
                FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
        Exit Function

    End If
    
    'devuelve el guildindex
    Gi = GuildIndex(GuildName)

    If Gi > 0 Then
        If UserList(UserIndex).EscucheClan <> 0 Then
            If UserList(UserIndex).EscucheClan = Gi Then
                'Already listening to them...
                Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = Gi
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList( _
                        UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)

            End If

        End If
        
        Call guilds(Gi).ConectarGM(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = Gi
        UserList(UserIndex).EscucheClan = Gi
    Else
        Call WriteConsoleMsg(UserIndex, "Error, el clan no existe.", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0

    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'el index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(UserIndex)

End Sub

Public Function r_DeclararGuerra(ByVal UserIndex As Integer, _
                                 ByRef GuildGuerra As String, _
                                 ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi  As Integer
    Dim GIG As Integer

    r_DeclararGuerra = 0
    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function

    End If
    
    GIG = GuildIndex(GuildGuerra)

    If guilds(Gi).GetRelacion(GIG) = GUERRA Then
        refError = "Tu clan ya está en guerra con " & GuildGuerra & "."
        Exit Function

    End If
        
    If Gi = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan."
        Exit Function

    End If

    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & Gi & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function

    End If

    Call guilds(Gi).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(Gi)
    Call guilds(Gi).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call guilds(GIG).SetRelacion(Gi, RELACIONES_GUILD.GUERRA)
    
    r_DeclararGuerra = GIG

End Function

Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, _
                                        ByRef GuildPaz As String, _
                                        ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
    Dim Gi  As Integer
    Dim GIG As Integer

    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & Gi & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If

    If guilds(Gi).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estás en guerra con ese clan."
        Exit Function

    End If
    
    If Not guilds(Gi).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar."
        Exit Function

    End If

    Call guilds(Gi).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(Gi)
    Call guilds(Gi).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call guilds(GIG).SetRelacion(Gi, RELACIONES_GUILD.PAZ)
    
    r_AceptarPropuestaDePaz = GIG

End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, _
                                             ByRef GuildPro As String, _
                                             ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'devuelve el index al clan guildPro
    Dim Gi  As Integer
    Dim GIG As Integer

    r_RechazarPropuestaDeAlianza = 0
    Gi = UserList(UserIndex).GuildIndex
    
    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & Gi & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If
    
    If Not guilds(Gi).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function

    End If
    
    Call guilds(Gi).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(Gi).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds( _
            GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function

Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, _
                                         ByRef GuildPro As String, _
                                         ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'devuelve el index al clan guildPro
    Dim Gi  As Integer
    Dim GIG As Integer

    r_RechazarPropuestaDePaz = 0
    Gi = UserList(UserIndex).GuildIndex
    
    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & Gi & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If
    
    If Not guilds(Gi).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function

    End If
    
    Call guilds(Gi).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(Gi).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds( _
            GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function

Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, _
                                            ByRef GuildAllie As String, _
                                            ByRef refError As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
    Dim Gi  As Integer
    Dim GIG As Integer

    r_AceptarPropuestaDeAlianza = 0
    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningún clan."
        Exit Function

    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & Gi & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)."
        Exit Function

    End If

    If guilds(Gi).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = _
                "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function

    End If
    
    If Not guilds(Gi).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function

    End If

    Call guilds(Gi).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(Gi)
    Call guilds(Gi).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call guilds(GIG).SetRelacion(Gi, RELACIONES_GUILD.ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function

Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, _
                                      ByRef OtroClan As String, _
                                      ByVal Tipo As RELACIONES_GUILD, _
                                      ByRef Detalle As String, _
                                      ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim OtroClanGI As Integer
    Dim Gi         As Integer

    r_ClanGeneraPropuesta = False
    
    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = Gi Then
        refError = "No puedes declarar relaciones con tu propio clan."
        Exit Function

    End If
    
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe."
        Exit Function

    End If
    
    If guilds(OtroClanGI).HayPropuesta(Gi, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = RELACIONES_GUILD.PAZ Then
        If guilds(Gi).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estás en guerra con " & OtroClan
            Exit Function

        End If

    ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then

        If guilds(Gi).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function

        End If

    End If
    
    Call guilds(OtroClanGI).SetPropuesta(Tipo, Gi, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal UserIndex As Integer, _
                               ByRef OtroGuild As String, _
                               ByVal Tipo As RELACIONES_GUILD, _
                               ByRef refError As String) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim OtroClanGI As Integer
    Dim Gi         As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan."
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan."
        Exit Function

    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not guilds(Gi).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada."
        Exit Function

    End If
    
    r_VerPropuesta = guilds(Gi).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, _
                                    ByVal Tipo As RELACIONES_GUILD) As String()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi            As Integer
    Dim i             As Integer
    Dim proposalCount As Integer
    Dim proposals()   As String
    
    Gi = UserList(UserIndex).GuildIndex
    
    If Gi > 0 And Gi <= CANTIDADDECLANES Then

        With guilds(Gi)
            proposalCount = .CantidadPropuestas(Tipo)
            
            'Resize array to contain all proposals
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String

            End If
            
            'Store each guild name
            For i = 0 To proposalCount - 1
                proposals(i) = guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
            Next i

        End With

    End If
    
    r_ListaDePropuestas = proposals

End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, _
                                   ByVal guild As Integer, _
                                   ByRef Detalles As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")

    End If

    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")

    End If

    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")

    End If

    Call guilds(guild).InformarRechazoEnChar(Aspirante, Detalles)

End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")

    End If

    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")

    End If

    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")

    End If

    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)

End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, _
                                    ByRef Nombre As String, _
                                    ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi           As Integer
    Dim NroAspirante As Integer

    a_RechazarAspirante = False
    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function

    End If

    NroAspirante = guilds(Gi).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan."
        Exit Function

    End If

    Call guilds(Gi).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(Gi).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, _
                                    ByRef Nombre As String) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi           As Integer
    Dim NroAspirante As Integer

    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        Exit Function

    End If
    
    NroAspirante = guilds(Gi).NumeroDeAspirante(Nombre)

    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(Gi).DetallesSolicitudAspirante(NroAspirante)

    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim UserFile    As clsIniReader
    Dim Miembro     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    
    On Error GoTo Error

    Gi = UserList(UserIndex).GuildIndex
    
    Personaje = UCase$(Personaje)
    
    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No eres el líder de tu clan.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)

    End If

    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)

    End If

    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)

    End If
    
    NroAsp = guilds(Gi).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = guilds(Gi).GetMemberList()
        
        For i = 0 To UBound(list())

            If Personaje = list(i) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro del clan.", _
                    FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

    End If
    
    'ahora traemos la info
    
    Set UserFile = New clsIniReader
    
    With UserFile
        .Initialize (CharPath & Personaje & ".chr")
        
        ' Get the character's current guild
        GuildActual = val(.GetValue("GUILD", "GuildIndex"))

        If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
            GuildName = "<" & guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"

        End If
        
        'Get previous guilds
        Miembro = .GetValue("GUILD", "Miembro")

        If Len(Miembro) > 400 Then
            Miembro = ".." & Right$(Miembro, 400)

        End If
        
        'Shermie80 clanes modificar esto xd
        Call Protocol.WriteCharacterInfo(UserIndex, Personaje, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), _
                .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), .GetValue("STATS", _
                "Banco"), .GetValue("GUILD", "Pedidos"), GuildName, Miembro, .GetValue( _
                "FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), .GetValue("FACCIONES", _
                "CiudMatados"), .GetValue("FACCIONES", "ReneMatados"))

    End With
    
    Set UserFile = Nothing
    
    Exit Sub
Error:
    Set UserFile = Nothing

    If Not (FileExist(CharPath & Personaje & ".chr", vbArchive)) Then
        Call LogError("El usuario " & UserList(UserIndex).Name & " (" & UserIndex & _
                " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
    Else
        Call LogError("[" & Err.Number & "] " & Err.description & _
                " En la rutina SendDetallesPersonaje, por el usuario " & UserList(UserIndex).Name & " (" & UserIndex _
                & " ), pidiendo información sobre el personaje " & Personaje)

    End If

End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, _
                                 ByRef clan As String, _
                                 ByRef Solicitud As String, _
                                 ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ViejoSolicitado   As String
    Dim ViejoGuildINdex   As Integer
    Dim ViejoNroAspirante As Integer
    Dim NuevoGuildIndex   As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro."
        Exit Function

    End If
    
    If EsNewbie(UserIndex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function

    End If

    NuevoGuildIndex = GuildIndex(clan)

    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe, avise a un administrador."
        Exit Function

    End If
    
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = "Tú no puedes entrar a un clan de alineación " & Alineacion2String(guilds( _
                NuevoGuildIndex).Alineacion)
        Exit Function

    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function

    End If

    ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).Name & ".chr", "GUILD", "ASPIRANTEA")

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)

        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).Name)

            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).Name, ViejoNroAspirante)

            End If

        Else

            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If

    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).Name, Solicitud)
    a_NuevoAspirante = True

End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, _
                                   ByRef Aspirante As String, _
                                   ByRef refError As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Gi           As Integer
    Dim NroAspirante As Integer
    Dim AspiranteUI  As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    Gi = UserList(UserIndex).GuildIndex

    If Gi <= 0 Or Gi > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function

    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, Gi) Then
        refError = "No eres el líder de tu clan"
        Exit Function

    End If
    
    NroAspirante = guilds(Gi).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan."
        Exit Function

    End If
    
    AspiranteUI = NameIndex(Aspirante)

    If AspiranteUI > 0 Then

        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, Gi) Then
            refError = Aspirante & " no puede entrar a un clan de alineación " & Alineacion2String(guilds( _
                    Gi).Alineacion)
            Call guilds(Gi).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(Gi).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function

        End If

    Else

        If Not m_EstadoPermiteEntrarChar(Aspirante, Gi) Then
            refError = Aspirante & " no puede entrar a un clan de alineación " & Alineacion2String(guilds( _
                    Gi).Alineacion)
            Call guilds(Gi).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetGuildIndexFromChar(Aspirante) Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(Gi).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function

        End If

    End If

    'el pj es aspirante al clan y puede entrar
    
    Call guilds(Gi).RetirarAspirante(Aspirante, NroAspirante)
    Call guilds(Gi).AceptarNuevoMiembro(Aspirante)
    
    ' If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)

    End If
    
    a_AceptarAspirante = True

End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildName = guilds(GuildIndex).GuildName

End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader

End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)

End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String

    '***************************************************
    'Autor: ZaMa
    'Returns the guild founder's name
    'Last Modification: 25/03/2009
    '***************************************************
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
    GuildFounder = guilds(GuildIndex).Fundador

End Function


