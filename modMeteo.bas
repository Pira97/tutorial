Attribute VB_Name = "modMeteo"
Option Explicit

Public ambientLight As D3DCOLORVALUE
Public AmbientColor As Long

Public meteo_state As Integer

Public meteo_hour As Integer

Public meteo_color As D3DCOLORVALUE

Public m_Color_Dia As D3DCOLORVALUE
Public m_Color_Noche As D3DCOLORVALUE
Public m_Color_Tarde As D3DCOLORVALUE
Public m_Color_Manana As D3DCOLORVALUE

Public m_Afecta As Boolean

'Meteorologia
Public tHora As Byte
Public tMinuto As Byte
Public tSeg As Byte
Public tCartel As Integer


Public Function Meteo_Change_Time()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
'Change the meteo time for start the animation with alphacolor
'**************************************************************
Dim tmpCartel As Byte

If meteo_hour = tHora Then
    If Not AmbientColor = -1 Then Exit Function
End If

meteo_hour = tHora
If tHora >= 5 And tHora <= 7 Then
    meteo_color = m_Color_Manana
ElseIf tHora >= 8 And tHora <= 17 Then
    meteo_color = m_Color_Dia
ElseIf tHora = 18 Or tHora = 19 Then
    meteo_color = m_Color_Tarde
Else
    meteo_color = m_Color_Noche
End If
 
frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex("c" & CStr(tHora))

If ambientLight.r > meteo_color.r Then
    meteo_state = 1 'Animacion Desendiente
Else
    meteo_state = 2 'Animacion Asendente
End If
End Function
Public Function AsignarHora()

    Dim Horario() As String
    
    Horario() = Split(Time, ":")
    
    tHora = CByte(Horario(0))
    tMinuto = CByte(Horario(1))
    
End Function

Public Function Get_Time_String() As String

Get_Time_String = mid(Time, 1, 5) & "... "
    
Select Case tHora
    Case 5, 6, 7
        Get_Time_String = Get_Time_String & "el sol se asoma lentamente en el horizonte"
    Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
        Get_Time_String = Get_Time_String & "¡no pierdas el tiempo!"
    Case 18, 19
        Get_Time_String = Get_Time_String & "lentamente el dia termina"
    Case Else
        Get_Time_String = Get_Time_String & "¿despierto a estas horas? ¡no olvides visitar El Mesón Hostigado!"
End Select

If Queclima = 1 Then
 Get_Time_String = mid(Time, 1, 5) & "... " & "¡Hay lluvia en las tierras de LinkAO!, ¡ten cuidado con tu energía!"
ElseIf Queclima = 2 Then
 Get_Time_String = mid(Time, 1, 5) & "... " & "¡Parece que hay una fuerte tormenta electrica, cuida tu energía!"
ElseIf Queclima = 3 Then

 Get_Time_String = mid(Time, 1, 5) & "... " & "¡Parece que está nevando!"
End If

End Function

Public Function Meteo_Render()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
'Rendering the animation with desvan
'**************************************************************
Dim change As Boolean

If meteo_state = 0 Then Exit Function
If m_Afecta = False Then Exit Function

Select Case meteo_state
    Case 1
        If ambientLight.r > meteo_color.r Then
            ambientLight.r = ambientLight.r - 1
            change = True
        End If
        
        If ambientLight.g > meteo_color.g Then
            ambientLight.g = ambientLight.g - 1
            change = True
        End If
        
        If ambientLight.b > meteo_color.b Then
            ambientLight.b = ambientLight.b - 1
            change = True
        End If
        
        If change = False Then meteo_state = 0
        AmbientColor = D3DColorXRGB(ambientLight.r, ambientLight.g, ambientLight.b)
        
    Case 2
        If ambientLight.r < meteo_color.r Then
            ambientLight.r = ambientLight.r + 1
            change = True
        End If
        
        If ambientLight.g < meteo_color.g Then
            ambientLight.g = ambientLight.g + 1
            change = True
        End If
        
        If ambientLight.b < meteo_color.b Then
            ambientLight.b = ambientLight.b + 1
            change = True
        End If
        
        If change = False Then meteo_state = 0
        AmbientColor = D3DColorXRGB(ambientLight.r, ambientLight.g, ambientLight.b)
        
        
End Select
End Function
Public Function Meteo_Init_Time()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
'Start all shades of the day
'**************************************************************
With m_Color_Dia
    .a = 255
    .b = 255
    .r = 255
    .g = 255
End With

With m_Color_Noche
    .a = 255
    .b = 170
    .r = 170
    .g = 170
End With

With m_Color_Tarde
    .a = 255
    .b = 200
    .r = 230
    .g = 200
End With

With m_Color_Manana
    .a = 255
    .b = 230
    .r = 200
    .g = 200
End With

meteo_hour = -1
meteo_state = -1
tCartel = -1
End Function
Public Function Meteo_Clean()
    meteo_hour = -1
    meteo_state = -1
    tCartel = -1
End Function

