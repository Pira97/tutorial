Attribute VB_Name = "History"
Option Explicit

'Aca los archivos de texto, MOTD y UPDATE

Public Type tMOTD

MOTD() As String
MotdMaxLines As Integer
fuente() As Byte

End Type

Public Type tUpdate

Update() As String 'Lineas
UpdateMaxLines As Integer

End Type

Public Update As tUpdate

Public MOTD As tMOTD


Sub LoadMotd()

    If frmmain.Visible Then frmmain.AgregarConsola "Cargando MOTD.ini"
    
    Dim i As Integer
    
    MOTD.MotdMaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))

    ReDim MOTD.MOTD(1 To MOTD.MotdMaxLines)
    ReDim MOTD.fuente(1 To MOTD.MotdMaxLines)
    
    For i = 1 To MOTD.MotdMaxLines
        MOTD.MOTD(i) = GetVar(DatPath & "Motd.ini", "Motd", "Line" & i)
        MOTD.fuente(i) = GetVar(DatPath & "Motd.ini", "Motd", "Fuente" & i)
    Next i
    
    If frmmain.Visible Then frmmain.AgregarConsola "MOTD.ini se cargó correctamente. " & Time
    
End Sub

Sub SendMOTD(ByVal UserIndex As Integer)

 
    Dim j As Long
    
    For j = 1 To MOTD.MotdMaxLines
        Call WriteConsoleMsg(UserIndex, MOTD.MOTD(j), MOTD.fuente(j))
    Next j
 
End Sub

Sub LoadUpdate()

    Dim i As Integer
    
    If frmmain.Visible Then frmmain.AgregarConsola "Cargando Actualizaciones.ini"
    
    Dim Inverso As Integer
    
    Update.UpdateMaxLines = val(GetVar(App.Path & "\Dat\Actualizaciones.ini", "INIT", "NumLines"))
    
    Inverso = Update.UpdateMaxLines
    ReDim Update.Update(1 To Update.UpdateMaxLines)
    
    For i = 1 To Update.UpdateMaxLines
        Update.Update(i) = GetVar(DatPath & "Actualizaciones.ini", "UPDATE", "Line" & Inverso)
        Inverso = (Inverso - 1)
    Next i
    
    If frmmain.Visible Then frmmain.AgregarConsola "Actualizaciones.ini se cargó correctamente. " & Time
    
End Sub

Sub SendUpdate(ByVal UserIndex As Integer)


    Dim j As Long
    
    For j = 1 To Update.UpdateMaxLines
        Call WriteEjecutarAccion(UserIndex, 4, Update.Update(j))
    Next j

End Sub
