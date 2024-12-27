Attribute VB_Name = "mod_Particulas"
'*************************************************************
'ImperiumAO 1.4.6
'*************************************************************
'Este modulo contiene TODOS los procedimientos que conforma
'el Sistema de Particulas ORE.
'*************************************************************

Option Explicit

Public Type RGB
    r As Long
    g As Long
    b As Long
End Type


Private Type Particle
    PartCountLive As Integer
    destruir  As Boolean
    Friction As Single
    X As Single
    Y As Single
    Vector_X As Single
    Vector_Y As Single
    angle As Single
    Grh As Grh
    alive_counter As Single
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Public Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    Friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    Grh_List() As Long
    ColortInt(0 To 3) As RGB
    speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
 
Private Type particle_group
    PartCountLive As Integer
    active As Boolean
    destruir As Boolean
    Creando As Integer
    Creada As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Single
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Dim Particle_Group_List() As particle_group
Dim Particle_Group_Count  As Long
Dim Particle_Group_Last   As Long

Public TotalStreams           As Long
Public StreamData()          As Stream

 Public Sub CargarParticulas()
 
  On Error GoTo CargarParticulas_Err
  
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    Dim Leer As New clsIniManager
    
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "particles.ini", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo CargarParticulas_Err
    End If
     
    Call Leer.Initialize(Resource_Path & "particles.ini")
    
    TotalStreams = val(Leer.GetValue("INIT", "Total"))
    
 
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams

        With StreamData(loopc)
            .Name = Leer.GetValue(val(loopc), "Name")
            .NumOfParticles = Leer.GetValue(val(loopc), "NumOfParticles")
            .x1 = Leer.GetValue(val(loopc), "X1")
            .y1 = Leer.GetValue(val(loopc), "Y1")
            .x2 = Leer.GetValue(val(loopc), "X2")
            .y2 = Leer.GetValue(val(loopc), "Y2")
            
            If loopc = 25 Or loopc = 49 Or loopc = 66 Or loopc = 84 Or loopc = 86 Or loopc = 117 Or loopc = 126 Or loopc = 10 Then
                
                .x1 = .x1 - 16
                .y1 = .y1 - 16
                .x2 = .x2 - 16
                .y2 = .y2 - 16
            
            End If
            
            .angle = Leer.GetValue(val(loopc), "Angle")
            .vecx1 = Leer.GetValue(val(loopc), "VecX1")
            .vecx2 = Leer.GetValue(val(loopc), "VecX2")
            .vecy1 = Leer.GetValue(val(loopc), "VecY1")
            .vecy2 = Leer.GetValue(val(loopc), "VecY2")
            .life1 = Leer.GetValue(val(loopc), "Life1")
            .life2 = Leer.GetValue(val(loopc), "Life2")
            .Friction = Leer.GetValue(val(loopc), "Friction")
            .spin = Leer.GetValue(val(loopc), "Spin")
            .spin_speedL = Leer.GetValue(val(loopc), "Spin_SpeedL")
            .spin_speedH = Leer.GetValue(val(loopc), "Spin_SpeedH")
            .AlphaBlend = Leer.GetValue(val(loopc), "AlphaBlend")
            .gravity = Leer.GetValue(val(loopc), "Gravity")
            .grav_strength = Leer.GetValue(val(loopc), "Grav_Strength")
            .bounce_strength = Leer.GetValue(val(loopc), "Bounce_Strength")
            .XMove = Leer.GetValue(val(loopc), "XMove")
            .YMove = Leer.GetValue(val(loopc), "YMove")
            .move_x1 = Leer.GetValue(val(loopc), "move_x1")
            .move_x2 = Leer.GetValue(val(loopc), "move_x2")
            .move_y1 = Leer.GetValue(val(loopc), "move_y1")
            .move_y2 = Leer.GetValue(val(loopc), "move_y2")
            .life_counter = Leer.GetValue(val(loopc), "life_counter")
            .speed = val(Leer.GetValue(val(loopc), "Speed"))
            
            Dim temp As Integer: temp = val(Leer.GetValue(val(loopc), "resize"))
             
            .grh_resize = IIf((temp = -1), True, False)
            .grh_resizex = val(Leer.GetValue(val(loopc), "rx"))
            .grh_resizey = val(Leer.GetValue(val(loopc), "ry"))
            
        
            .NumGrhs = Leer.GetValue(val(loopc), "NumGrhs")
  
            ReDim .Grh_List(1 To .NumGrhs)
            GrhListing = Leer.GetValue(val(loopc), "Grh_List")
              
            For i = 1 To StreamData(loopc).NumGrhs
                .Grh_List(i) = ReadField(i, GrhListing, 44)
            Next i

            .Grh_List(i - 1) = .Grh_List(i - 1)
            
            For ColorSet = 1 To 4
 
                TempSet = Leer.GetValue(val(loopc), "ColorSet" & ColorSet)
                .ColortInt(ColorSet - 1).r = ReadField(1, TempSet, 44)
                .ColortInt(ColorSet - 1).g = ReadField(2, TempSet, 44)
                .ColortInt(ColorSet - 1).b = ReadField(3, TempSet, 44)
                
            Next ColorSet
        End With
    
    Next loopc
    Set Leer = Nothing
    Delete_File Resource_Path & "particles.ini"
    
    
    Exit Sub

CargarParticulas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Recursos.CargarParticulas", Erl)
    If General_File_Exists(Resource_Path & "particles.ini", vbNormal) Then Delete_File Resource_Path & "particles.ini"
    
End Sub

Public Function SetMapParticle(ByVal ParticulaInd As Long, _
                               ByVal X As Integer, _
                               ByVal Y As Integer, _
                               Optional ByVal particle_life As Long = 0) As Long
   
     On Error GoTo General_Particle_Create_Err
      If ParticulaInd <= 0 Or ParticulaInd > UBound(StreamData) Then Exit Function


    'Mermas ok
    Dim rgb_list(0 To 3) As Long
    Dim PartIndex As Integer

    With StreamData(ParticulaInd)
        rgb_list(0) = RGB(.ColortInt(0).r, .ColortInt(0).g, .ColortInt(0).b)
        rgb_list(1) = RGB(.ColortInt(1).r, .ColortInt(1).g, .ColortInt(1).b)
        rgb_list(2) = RGB(.ColortInt(2).r, .ColortInt(2).g, .ColortInt(2).b)
        rgb_list(3) = RGB(.ColortInt(3).r, .ColortInt(3).g, .ColortInt(3).b)
    
 
        PartIndex = Particle_Group_Create(X, Y, .Grh_List, rgb_list(), .NumOfParticles, ParticulaInd, .AlphaBlend, _
                IIf(particle_life = 0, .life_counter, particle_life), .speed, , .x1, .y1, .angle, .vecx1, .vecx2, _
                .vecy1, .vecy2, .life1, .life2, .Friction, .spin_speedL, .gravity, .grav_strength, .bounce_strength, _
                .x2, .y2, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, .spin_speedH, .spin, .grh_resize, .grh_resizex, .grh_resizey)
 
    End With
 
    SetMapParticle = PartIndex
    
    Exit Function

General_Particle_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.General_Particle_Create", Erl)
    Resume Next
    
End Function

Public Function SetCharacterParticle(ByVal ParticulaInd As Long, ByVal charindex As Integer, Optional ByVal particle_life As Single = 0) As Long
   
    On Error GoTo General_Char_Particle_Create_Err
        If ParticulaInd = 0 Then Exit Function
    'Mermas ok
    Dim rgb_list(0 To 3) As Long
    Dim PartIndex As Integer
    
    
    With StreamData(ParticulaInd)
        rgb_list(0) = RGB(.ColortInt(0).r, .ColortInt(0).g, .ColortInt(0).b)
        rgb_list(1) = RGB(.ColortInt(1).r, .ColortInt(1).g, .ColortInt(1).b)
        rgb_list(2) = RGB(.ColortInt(2).r, .ColortInt(2).g, .ColortInt(2).b)
        rgb_list(3) = RGB(.ColortInt(3).r, .ColortInt(3).g, .ColortInt(3).b)
 
        PartIndex = Char_Particle_Group_Create(charindex, .Grh_List, rgb_list, .NumOfParticles, ParticulaInd, _
                .AlphaBlend, IIf(particle_life = 0, .life_counter, particle_life), .speed, , .x1, .y1 - 15, .angle, _
                .vecx1, .vecx2, .vecy1, .vecy2, .life1, .life2, .Friction, .spin_speedL, .gravity, .grav_strength, _
                .bounce_strength, .x2, .y2 - 15, .XMove, .move_x1, .move_x2, .move_y1, .move_y2, .YMove, .spin_speedH, _
                .spin, .grh_resize, .grh_resizex, .grh_resizey)

    End With

    SetCharacterParticle = PartIndex

    
    Exit Function

General_Char_Particle_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.General_Char_Particle_Create", Erl)
    Resume Next
    
End Function
Private Function Particle_Group_Next_Open() As Long


    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If Particle_Group_Last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until Particle_Group_List(loopc).active = False
        If loopc = Particle_Group_Last Then
            Particle_Group_Next_Open = Particle_Group_Last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
    
End Function
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
       'Mermas ok
    On Error GoTo Particle_Group_Check_Err
    
    If particle_group_index > 0 And particle_group_index <= Particle_Group_Last Then
        If Particle_Group_List(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
        
    Exit Function

Particle_Group_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Check", Erl)
    Resume Next
End Function
Private Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Single = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    
    On Error GoTo Particle_Group_Create_Err
    'Mermas Ok
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
        Exit Function

Particle_Group_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Create", Erl)
    Resume Next
End Function
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
    
    On Error GoTo Particle_Group_Remove_Err
    'Mermas Ok
    
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_List(particle_group_index).never_die = False
        Particle_Group_List(particle_group_index).alive_counter = 0
        
        Particle_Group_List(particle_group_index).destruir = True
    
        Rem Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True

    End If
        
    Exit Function

Particle_Group_Remove_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Remove", Erl)
    Resume Next
End Function
Public Function Particle_Group_Remove_All() As Boolean
      'Mermas Ok
    On Error GoTo Particle_Group_Remove_All_Err
    
    Dim Index As Long
    
    For Index = 1 To Particle_Group_Last
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
        Exit Function

Particle_Group_Remove_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Remove_All", Erl)
    Resume Next
    
End Function
Private Function Particle_Group_Find(ByVal id As Long) As Long
   '*****************************************************************
    On Error GoTo ErrorHandler:
    'Mermas Ok
    Dim loopc As Long
    
    loopc = 1
    Do Until Particle_Group_List(loopc).id = id
        If loopc = Particle_Group_Last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
    
End Function
Public Function Particle_Get_Type(ByVal particle_group_index As Long) As Byte

    'Mermas ok
    On Error GoTo Particle_Get_Type_Err
    If Particle_Group_Check(particle_group_index) Then
        Particle_Get_Type = Particle_Group_List(particle_group_index).stream_type

    End If


Exit Function
 

Particle_Get_Type_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Get_Type", Erl)
    Resume Next
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
    On Error GoTo Particle_Group_Destroy_Err
     'Mermas Ok
    Dim temp  As particle_group

    Dim i     As Integer

    Dim ii    As Integer

    Dim b     As Integer

    Dim antes As Integer
    
    If Particle_Group_List(particle_group_index).map_x > 0 And Particle_Group_List(particle_group_index).map_y > 0 Then
        MapData(Particle_Group_List(particle_group_index).map_x, Particle_Group_List(particle_group_index).map_y).particle_group = 0
    ElseIf Particle_Group_List(particle_group_index).char_index Then

        If Char_Check(Particle_Group_List(particle_group_index).char_index) Then

            For i = 1 To charlist(Particle_Group_List(particle_group_index).char_index).particle_count

                If charlist(Particle_Group_List(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    antes = charlist(Particle_Group_List(particle_group_index).char_index).particle_count
                    charlist(Particle_Group_List(particle_group_index).char_index).particle_count = charlist(Particle_Group_List(particle_group_index).char_index).particle_count - 1
                    charlist(Particle_Group_List(particle_group_index).char_index).particle_group(i) = 0
                    
                    ii = i
                    
                    For b = ii To antes - 1
                        charlist(Particle_Group_List(particle_group_index).char_index).particle_group(b) = charlist(Particle_Group_List(particle_group_index).char_index).particle_group(b + 1)
                        ' charlist(particle_group_list(particle_group_index).char_index).particle_group(b + 1) = 0
                    Next b

                    Rem       ReDim Preserve charlist(particle_group_list(particle_group_index).char_index).particle_group(1 To charlist(particle_group_list(particle_group_index).char_index).particle_count)
                    Rem Else
                    Rem ReDim charlist(particle_group_list(particle_group_index).char_index).particle_group(0)
                    '  End If
                    
                    Exit For
                    
                End If
                
            Next i
            
        End If
        
    ElseIf particle_group_index = meteo_particle Then
        meteo_particle = 0
    End If
    
    
    Particle_Group_List(particle_group_index) = temp
    
    If particle_group_index = Particle_Group_Last Then

        Do Until Particle_Group_List(Particle_Group_Last).active
            Particle_Group_Last = Particle_Group_Last - 1

            If Particle_Group_Last = 0 Then
                Particle_Group_Count = 0
                Exit Sub

            End If

        Loop
        Debug.Print Particle_Group_Last & "," & UBound(Particle_Group_List)
        ReDim Preserve Particle_Group_List(1 To Particle_Group_Last)

    End If

    Particle_Group_Count = Particle_Group_Count - 1
        
    Exit Sub

Particle_Group_Destroy_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Destroy", Erl)
    Resume Next
    
End Sub

Public Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
    
    'Mermas ok
    On Error GoTo Particle_Group_Render_Err
    
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    
    Dim no_move As Boolean
 
    With Particle_Group_List(particle_group_index)
    
    temp_rgb(0) = .rgb_list(1)
    temp_rgb(1) = .rgb_list(1)
    temp_rgb(2) = .rgb_list(2)
    temp_rgb(3) = .rgb_list(0)
    
    If Particle_Group_List(particle_group_index).alive_counter Then
        'See if it is time to move a particle
        Particle_Group_List(particle_group_index).frame_counter = Particle_Group_List(particle_group_index).frame_counter + (timerTicksPerFrame * 2)
        If Particle_Group_List(particle_group_index).frame_counter > Particle_Group_List(particle_group_index).frame_speed Then
            Particle_Group_List(particle_group_index).frame_counter = 0
            no_move = False
                    
        Else
            no_move = True
                    
        End If
  
        Dim Cantidad As Long

        Cantidad = Particle_Group_List(particle_group_index).particle_count
      
        For loopc = 1 To Cantidad
                
            Particle_Render .particle_stream(loopc), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(RandomNumber(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alpha_blend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin, .grh_resize, .grh_resizex, .grh_resizey, _
                        particle_group_index, Particle_Group_List(particle_group_index).destruir
        Next loopc
                
    If no_move = False Then
        If .never_die = False Then
            .alive_counter = .alive_counter - 1
        End If
    End If
    
    
    Else
    
        Particle_Group_List(particle_group_index).destruir = True
            
        If Particle_Group_List(particle_group_index).PartCountLive <= 2 Then
                  
            Particle_Group_Destroy particle_group_index
            Exit Sub

        End If
        
        
        Particle_Group_List(particle_group_index).frame_counter = Particle_Group_List(particle_group_index).frame_counter + (timerTicksPerFrame * 2)
        
        If Particle_Group_List(particle_group_index).frame_counter > Particle_Group_List(particle_group_index).frame_speed Then
            Particle_Group_List(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True

        End If
        
        
        'If it's still alive render all the particles inside
        For loopc = 1 To Particle_Group_List(particle_group_index).particle_count
        
            'Render particle
            Particle_Render .particle_stream(loopc), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(RandomNumber(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alpha_blend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin, .grh_resize, .grh_resizex, .grh_resizey, _
                        particle_group_index, Particle_Group_List(particle_group_index).destruir
        Next loopc

        Particle_Group_List(particle_group_index).destruir = True

    End If
    
    End With
        Exit Sub

Particle_Group_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Render", Erl)
    Resume Next
    
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, _
                            ByVal screen_x As Integer, _
                            ByVal screen_y As Integer, _
                            ByVal grh_index As Long, _
                            ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, _
                            Optional ByVal y1 As Integer, _
                            Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, _
                            Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, _
                            Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, _
                            Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, _
                            Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, _
                            Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, _
                            Optional ByVal x2 As Integer, _
                            Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, _
                            Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional particle_group_index As Long, Optional destruir As Boolean)
    '**************************************************************
    
    On Error GoTo Particle_Render_Err
    
    If no_move = False Then
        If temp_particle.alive_counter = 0 And Not destruir Then
            InitGrh temp_particle.Grh, grh_index, alpha_blend
            temp_particle.X = RandomNumber(x1, x2)
            temp_particle.Y = RandomNumber(y1, y2)
            temp_particle.Vector_X = RandomNumber(vecx1, vecx2)
            temp_particle.Vector_Y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.Friction = fric
            Particle_Group_List(particle_group_index).PartCountLive = Particle_Group_List(particle_group_index).PartCountLive + 1
        Else


            If temp_particle.alive_counter = 0 And destruir Then
                temp_particle.Grh.GrhIndex = 0
                
            End If
            
            If gravity = True Then
                temp_particle.Vector_Y = temp_particle.Vector_Y + grav_strength

                If temp_particle.Y > 0 Then
                    temp_particle.Vector_Y = bounce_strength

                End If

            End If
            
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + RandomNumber(spin_speedL, spin_speedH) / 5
            Do While temp_particle.Grh.angle >= 360
                temp_particle.Grh.angle = temp_particle.Grh.angle - 360
            Loop
 
            
            If XMove = True Then temp_particle.Vector_X = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.Vector_Y = RandomNumber(move_y1, move_y2)

        End If
        
        temp_particle.X = temp_particle.X + (temp_particle.Vector_X \ temp_particle.Friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.Vector_Y \ temp_particle.Friction)
    
        temp_particle.alive_counter = temp_particle.alive_counter - 1

    End If
    
    temp_particle.grh_resize = grh_resize
    temp_particle.grh_resizex = grh_resizex
    temp_particle.grh_resizey = grh_resizey
    
    
    If grh_resize = True Then
    
        If temp_particle.Grh.GrhIndex Then
        
            Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            
            Exit Sub

        End If

    End If

    If temp_particle.Grh.GrhIndex Then
        Grh_Render temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, rgb_list(), True, True, alpha_blend

    End If
    Exit Sub

Particle_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Render", Erl)
    Resume Next
    
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Single = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
        On Error GoTo Particle_Group_Make_Err
    
    'Mermas ok
    If particle_group_index > Particle_Group_Last Then
        Particle_Group_Last = particle_group_index
        ReDim Preserve Particle_Group_List(1 To Particle_Group_Last)
    End If
    
    Particle_Group_Count = Particle_Group_Count + 1
    
    Particle_Group_List(particle_group_index).active = True
    
    If (map_x <> -1) And (map_y <> -1) Then
        Particle_Group_List(particle_group_index).map_x = map_x
        Particle_Group_List(particle_group_index).map_y = map_y
    End If
    
    ReDim Particle_Group_List(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    Particle_Group_List(particle_group_index).grh_index_list() = grh_index_list()
    Particle_Group_List(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        Particle_Group_List(particle_group_index).alive_counter = -1
        Particle_Group_List(particle_group_index).never_die = True
    Else
        Particle_Group_List(particle_group_index).alive_counter = alive_counter
        Particle_Group_List(particle_group_index).never_die = False

    End If
    
    Particle_Group_List(particle_group_index).alpha_blend = alpha_blend
    
    Particle_Group_List(particle_group_index).stream_type = stream_type
    
    Particle_Group_List(particle_group_index).frame_speed = frame_speed
    
    Particle_Group_List(particle_group_index).x1 = x1
    Particle_Group_List(particle_group_index).y1 = y1
    Particle_Group_List(particle_group_index).x2 = x2
    Particle_Group_List(particle_group_index).y2 = y2
    Particle_Group_List(particle_group_index).angle = angle
    Particle_Group_List(particle_group_index).vecx1 = vecx1
    Particle_Group_List(particle_group_index).vecx2 = vecx2
    Particle_Group_List(particle_group_index).vecy1 = vecy1
    Particle_Group_List(particle_group_index).vecy2 = vecy2
    Particle_Group_List(particle_group_index).life1 = life1
    Particle_Group_List(particle_group_index).life2 = life2
    Particle_Group_List(particle_group_index).fric = fric
    Particle_Group_List(particle_group_index).spin = spin
    Particle_Group_List(particle_group_index).spin_speedL = spin_speedL
    Particle_Group_List(particle_group_index).spin_speedH = spin_speedH
    Particle_Group_List(particle_group_index).gravity = gravity
    Particle_Group_List(particle_group_index).grav_strength = grav_strength
    Particle_Group_List(particle_group_index).bounce_strength = bounce_strength
    Particle_Group_List(particle_group_index).XMove = XMove
    Particle_Group_List(particle_group_index).YMove = YMove
    Particle_Group_List(particle_group_index).move_x1 = move_x1
    Particle_Group_List(particle_group_index).move_x2 = move_x2
    Particle_Group_List(particle_group_index).move_y1 = move_y1
    Particle_Group_List(particle_group_index).move_y2 = move_y2
    
    Particle_Group_List(particle_group_index).rgb_list(0) = rgb_list(0)
    Particle_Group_List(particle_group_index).rgb_list(1) = rgb_list(1)
    Particle_Group_List(particle_group_index).rgb_list(2) = rgb_list(2)
    Particle_Group_List(particle_group_index).rgb_list(3) = rgb_list(3)
    
    Particle_Group_List(particle_group_index).grh_resize = grh_resize
    Particle_Group_List(particle_group_index).grh_resizex = grh_resizex
    Particle_Group_List(particle_group_index).grh_resizey = grh_resizey
    
    Particle_Group_List(particle_group_index).id = id
 
    Particle_Group_List(particle_group_index).particle_count = particle_count
    ReDim Particle_Group_List(particle_group_index).particle_stream(1 To particle_count)
    
    If (map_x <> -1) And (map_y <> -1) Then
    
      MapData(map_x, map_y).particle_group = particle_group_index
    
    End If
    
    Exit Sub

Particle_Group_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Particle_Group_Make", Erl)
    Resume Next
    
    
End Sub

Private Function Map_Particle_Group_Get(ByVal map_x As Integer, _
                                       ByVal map_y As Integer) As Long
    On Error GoTo Map_Particle_Group_Get_Err
        'Mermas ok
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0

    End If
        
    Exit Function

Map_Particle_Group_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Map_Particle_Group_Get", Erl)
    Resume Next
End Function

Private Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Single = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    
    On Error GoTo Char_Particle_Group_Create_Err
 
'mermas ok
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
    Exit Function

Char_Particle_Group_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Char_Particle_Group_Create", Erl)
    Resume Next
End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
          
    On Error GoTo Char_Particle_Group_Remove_Err
     'mermas ok
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        If char_part_index = 0 Then Exit Function
        
        Particle_Group_List(char_part_index).alive_counter = 0
        Particle_Group_List(char_part_index).never_die = False
        Particle_Group_List(char_part_index).destruir = True
    End If
    
    Exit Function

Char_Particle_Group_Remove_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Char_Particle_Group_Remove", Erl)
    Resume Next
    
End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
     
    On Error GoTo Char_Particle_Group_Remove_All_Err
    
   'mermas ok
    Dim i As Integer
    
    If Char_Check(char_index) Then
         For i = 1 To charlist(char_index).particle_count
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
    End If
        
    
    Exit Function

Char_Particle_Group_Remove_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Char_Particle_Group_Remove_All", Erl)
    Resume Next
    
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************
    '*****************************************************************
    On Error GoTo ErrorHandler:

   'mermas ok
Dim i As Integer

    For i = 1 To charlist(char_index).particle_count

        If Particle_Group_List(charlist(char_index).particle_group(i)).stream_type = stream_type Then
            If Particle_Group_List(charlist(char_index).particle_group(i)).destruir = False Then
                Char_Particle_Group_Find = charlist(char_index).particle_group(i)
                Exit Function

            End If

        End If

    Next i

Char_Particle_Group_Find = -1
ErrorHandler:
End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
       'mermas ok
    If charlist(char_index).particle_count = 0 Then
        charlist(char_index).particle_count = 1
        ReDim charlist(char_index).particle_group(1 To 1)
        Char_Particle_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0

        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function

 

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Single = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                

    'Mermas ok
    On Error GoTo Char_Particle_Group_Make_Err
    
    'Update array size
    If particle_group_index > Particle_Group_Last Then
        Particle_Group_Last = particle_group_index
        ReDim Preserve Particle_Group_List(1 To Particle_Group_Last)
    End If
    Particle_Group_Count = Particle_Group_Count + 1
    
    'Make active
    Particle_Group_List(particle_group_index).active = True
    
    'Char index
    Particle_Group_List(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim Particle_Group_List(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    Particle_Group_List(particle_group_index).grh_index_list() = grh_index_list()
    Particle_Group_List(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then

        Particle_Group_List(particle_group_index).alive_counter = -1
        Particle_Group_List(particle_group_index).never_die = True
    Else
        '  Debug.Print alive_counter
        Particle_Group_List(particle_group_index).alive_counter = alive_counter
        Particle_Group_List(particle_group_index).never_die = False

    End If
    

    'alpha blending
    Particle_Group_List(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    Particle_Group_List(particle_group_index).stream_type = stream_type
    
    'speed
    Particle_Group_List(particle_group_index).frame_speed = frame_speed
    
    Particle_Group_List(particle_group_index).x1 = x1
    Particle_Group_List(particle_group_index).y1 = y1
    Particle_Group_List(particle_group_index).x2 = x2
    Particle_Group_List(particle_group_index).y2 = y2
    Particle_Group_List(particle_group_index).angle = angle
    Particle_Group_List(particle_group_index).vecx1 = vecx1
    Particle_Group_List(particle_group_index).vecx2 = vecx2
    Particle_Group_List(particle_group_index).vecy1 = vecy1
    Particle_Group_List(particle_group_index).vecy2 = vecy2
    Particle_Group_List(particle_group_index).life1 = life1
    Particle_Group_List(particle_group_index).life2 = life2
    Particle_Group_List(particle_group_index).fric = fric
    Particle_Group_List(particle_group_index).spin = spin
    Particle_Group_List(particle_group_index).spin_speedL = spin_speedL
    Particle_Group_List(particle_group_index).spin_speedH = spin_speedH
    Particle_Group_List(particle_group_index).gravity = gravity
    Particle_Group_List(particle_group_index).grav_strength = grav_strength
    Particle_Group_List(particle_group_index).bounce_strength = bounce_strength
    Particle_Group_List(particle_group_index).XMove = XMove
    Particle_Group_List(particle_group_index).YMove = YMove
    Particle_Group_List(particle_group_index).move_x1 = move_x1
    Particle_Group_List(particle_group_index).move_x2 = move_x2
    Particle_Group_List(particle_group_index).move_y1 = move_y1
    Particle_Group_List(particle_group_index).move_y2 = move_y2
    
    Particle_Group_List(particle_group_index).rgb_list(0) = rgb_list(0)
    Particle_Group_List(particle_group_index).rgb_list(1) = rgb_list(1)
    Particle_Group_List(particle_group_index).rgb_list(2) = rgb_list(2)
    Particle_Group_List(particle_group_index).rgb_list(3) = rgb_list(3)
    
    Particle_Group_List(particle_group_index).grh_resize = grh_resize
    Particle_Group_List(particle_group_index).grh_resizex = grh_resizex
    Particle_Group_List(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    Particle_Group_List(particle_group_index).id = id
    
    'create particle stream
    Particle_Group_List(particle_group_index).particle_count = particle_count
    ReDim Particle_Group_List(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
    
    Exit Sub
    
Char_Particle_Group_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "mod_Particulas.Char_Particle_Group_Make", Erl)
    Resume Next
End Sub






