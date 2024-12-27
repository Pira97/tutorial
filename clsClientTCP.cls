VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
Public Sub ParseUserCommand(ByVal strCmd As String)

    
    
    Dim strArg As String, strCompa As String
    Dim bFound As Boolean
    Dim Comando As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    Dim argumentos
    
    Dim tmpArr() As String
    Dim tmpInt As Integer
    
    Dim i As Integer ' Agrega esta declaración para usar la variable en el bucle For


    ' Asegurar longitud máxima de 255 caracteres
    If Len(strCmd) > 255 Then strCmd = mid$(strCmd, 1, 255)

    ' Convertir a mayúsculas para comandos no sensibles a mayúsculas

    strCompa = UCase$(Trim(strCmd))

    ' Dividir comando y argumentos
    argumentos = Split(strCompa, " ")
    Comando = argumentos(0) ' El primer elemento es el comando
    CantidadArgumentos = UBound(argumentos)
    notNullArguments = CantidadArgumentos > 0

    ' Manejo de comandos
    Select Case Comando
        Case "/PING"
            Call AddtoRichTextBox("Ping: " & MSRender & "ms", 0, 0, 0, 0, 0, True)

        Case "/SALIR"
            Call WriteQuit

        Case "/MEDITAR"
            If DeadCheck Then Exit Sub
            Call WriteMeditate

        Case "/ONLINE"
            Call WriteOnline

        Case "/ONLINECLAN"
            Call WriteGuildOnline

        Case "/GM", "/DM", "/AYUDA"
            frmGMAyuda.Show vbModeless, frmMain
            frmGMAyuda.TxtSoporte.SetFocus

        Case "/BALANCE"
            If DeadCheck Then Exit Sub
            MsgBox "Balance"

        Case "/EST"
            Call WriteRequestStats

        Case "/BOVEDA"
            If DeadCheck Then Exit Sub
            Call WriteBankStart

        Case "/COMERCIAR"
            If DeadCheck Then Exit Sub
            If Not Comerciando Then Call WriteCommerceStart

        Case "/INFORMACION"
            Call WriteInformation

        Case "/ENLISTAR"
            If DeadCheck Then Exit Sub
            Call WriteEnlist

        Case "/FUNDARCLAN"
            If DeadCheck Then Exit Sub
            Call WriteGuildFundate

        Case "/SALIRCLAN"
            If DeadCheck Then Exit Sub
            Call WriteGuildLeave

        Case "/CERRARCLAN"
            If DeadCheck Then Exit Sub
            Call writeCloseGuild

        Case "/DIVORCIO"
            If DeadCheck Then Exit Sub
            Call writeDivorciar

        Case "/PENA"
            MsgBox "Pena"

        Case "/HOGAR"
            Call WriteSeleccionarHogar(0)

        Case "/RETIRAR"
            If DeadCheck Then Exit Sub
            frmPregunta.SetAccion 2
            If frmMain.Visible Then frmPregunta.Show , frmMain

        Case "/FLIST"
            Call WriteOnAmigoandMapa

        Case "/PERDON"
            If DeadCheck Then Exit Sub
            MsgBox "Perdon"

        Case "/EVENTOS"
            Call WriteHayEventos

        Case "/SKIN"
            If DeadCheck Then Exit Sub
            Call WriteAbrirForms(1)

        Case "/SUM"
            ' Comando para invocar un personaje
            If CantidadArgumentos = 1 Then
                Call WriteSummonChar(argumentos(1))
            Else
                Call ShowConsoleMsg("Faltan parámetros. Utilice /sum NICKNAME.")
            End If
            Case "/CT"
    ' Verifica que haya argumentos y la cantidad sea correcta
    If notNullArguments And CantidadArgumentos = 3 Then
        
        ' Validamos que todos los argumentos sean del tipo esperado
        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And _
           ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And _
           ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
           
            ' Llamamos a la función para crear el teletransporte
        
        
            ' Mostrar mensaje si algún argumento no es válido
            Call ShowConsoleMsg("Valor incorrecto. Uso correcto: /CT MAPA X Y.")
        End If
        
    Else
        ' Mostrar mensaje si faltan argumentos
        Call ShowConsoleMsg("Faltan parámetros. Uso correcto: /CT MAPA X Y.")
    End If

End Select
    
End Sub


Public Function DeadCheck() As Boolean

If CurrentUser.Muerto Then
    Call AddtoRichTextBox(Locale_GUI_Frase(245), 0, 0, 0, 0, 0, 0, 12)
    DeadCheck = True
End If

End Function
Public Function ActualizarEst(Optional ByVal MaxHP As Integer = -1, Optional ByVal MinHP As Integer = -1, Optional ByVal MaxMAN As Integer = -1, _
    Optional ByVal MinMAN As Integer = -1, Optional ByVal MaxSTA As Integer = -1, Optional ByVal MinSTA As Integer = -1, _
    Optional ByVal GLD As Long = -1, Optional ByVal Nivel As Integer = -1, Optional ByVal PasarNivel As Long = -1, Optional ByVal EXP As Long = -1, _
    Optional ByVal Fuerza As Integer = -1, Optional ByVal Agilidad As Integer = -1, _
    Optional ByVal MinHAM As Integer = -1, Optional ByVal MaxHAM As Integer = -1, _
    Optional ByVal MinSED As Integer = -1, Optional ByVal MaxSED As Integer = -1, _
    Optional ByVal ActualizarTodos As Boolean = False)

Dim ActualizarCual As Byte

If MaxHP <> -1 Then
    CurrentUser.UserMaxHP = MaxHP
    ActualizarCual = 1
End If

If MinHP <> -1 Then

    If MinHP < 0 Then MinHP = 0
    CurrentUser.UserMinHP = MinHP
    ActualizarCual = 1
    
End If

If MaxMAN <> -1 Then
    CurrentUser.UserMaxMAN = MaxMAN
    ActualizarCual = 2
End If

If MinMAN <> -1 Then
    CurrentUser.UserMinMAN = MinMAN
    
    If CurrentUser.UserMinMAN > CurrentUser.UserMaxMAN Then _
        CurrentUser.UserMinMAN = CurrentUser.UserMaxMAN
        
    ActualizarCual = 2
End If

If MaxSTA <> -1 Then
    CurrentUser.UserMaxSTA = MaxSTA
    ActualizarCual = 3
End If

If MinSTA <> -1 Then

    If CurrentUser.UserMinSTA > CurrentUser.UserMaxSTA Then _
        CurrentUser.UserMinSTA = CurrentUser.UserMaxSTA

    CurrentUser.UserMinSTA = MinSTA
    ActualizarCual = 3
End If

If GLD <> -1 Then
    CurrentUser.UserGLD = GLD
    ActualizarCual = 4
End If

If Nivel <> -1 Then
    CurrentUser.UserLvl = Nivel
    ActualizarCual = 5
End If

If PasarNivel <> -1 Then
    CurrentUser.UserPasarNivel = PasarNivel
    ActualizarCual = 5
End If
    
If EXP <> -1 Then
    CurrentUser.UserExp = EXP
    ActualizarCual = 5
End If

If Fuerza <> -1 Then
    frmMain.lblStrg = Fuerza
    frmMain.lblStrg.ForeColor = IIf(Fuerza < 35, &HFFFFFF, &H40C0&)
End If

If Agilidad <> -1 Then
    frmMain.lblDext = Agilidad
    frmMain.lblDext.ForeColor = IIf(Agilidad < 35, &HFFFFFF, &H40C0&)
End If

If MinHAM <> -1 Then
    CurrentUser.UserMinHAM = MinHAM
    ActualizarCual = 6
End If

If MinSED <> -1 Then
    CurrentUser.UserMinAGU = MinSED
    ActualizarCual = 6
End If

If MaxHAM <> -1 Then
    CurrentUser.UserMaxHAM = MaxHAM
    ActualizarCual = 6
End If

If MaxSED <> -1 Then
    CurrentUser.UserMaxAGU = MaxSED
    ActualizarCual = 6
End If

If Not ActualizarTodos Then
    Select Case ActualizarCual
        Case 1
            Call ActualizarHP
        Case 2
            Call ActualizarMAN
        Case 3
            Call ActualizarSTA
        Case 4
            Call ActualizarGLD
        Case 5
            Call ActualizarExp
        Case 6
            Call ActualizarHAM
    End Select
Else
    Call ActualizarHP
    Call ActualizarMAN
    Call ActualizarSTA
    Call ActualizarGLD
    Call ActualizarExp
    Call ActualizarHAM
End If

End Function
Private Sub ActualizarHP()

If CurrentUser.UserMinHP <= 0 Then
    CurrentUser.UserMinHP = 0
    CurrentUser.Muerto = True
    Engine_Scroll_Pixels scroll_pixels_per_frameBackUp * VelocidadMuerto
    frmMain.lblVida.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.shpvida.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.shpvida.FillColor = &H808080
Else
    If CurrentUser.Muerto Then
        CurrentUser.Muerto = False
    End If
    
    If CurrentUser.Logged Then
        If (CurrentUser.Montando = False) Then
            Engine_Scroll_Pixels scroll_pixels_per_frameBackUp
        End If
    End If
    
    frmMain.lblVida.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.shpvida.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.shpvida.FillColor = &HC0&
End If

End Sub
Private Sub ActualizarGLD()
frmMain.GldLbl.Caption = CurrentUser.UserGLD
End Sub
Private Sub ActualizarMAN()

If CurrentUser.UserMaxMAN > 0 Then
    frmMain.shpmana.Width = (((CurrentUser.UserMinMAN + 1 / 100) / (CurrentUser.UserMaxMAN + 1 / 100)) * 91)
    frmMain.lblMana.Visible = True
    frmMain.lblMana.Caption = CurrentUser.UserMinMAN & "/" & CurrentUser.UserMaxMAN
Else
    frmMain.shpmana.Width = 0
    frmMain.lblMana.Visible = False
End If

End Sub
Private Sub ActualizarSTA()
frmMain.shpEnergia.Width = (((CurrentUser.UserMinSTA / 100) / (CurrentUser.UserMaxSTA / 100)) * 89)
frmMain.lblEnergia.Caption = CurrentUser.UserMinSTA & "/" & CurrentUser.UserMaxSTA
End Sub
Private Sub ActualizarExp()

frmMain.lblLvl.Caption = CurrentUser.UserLvl

Call UserExpPerc

If CurrentUser.UserPercExp <> 0 Then
    frmMain.ExpShp.Width = (((CurrentUser.UserExp / 100) / (CurrentUser.UserPasarNivel / 100)) * 467)
Else
    frmMain.ExpShp.Width = 0
End If
        
frmMain.Label2(1).Caption = IIf(frmMain.UltPos = 1, CurrentUser.UserExp & "/" & CurrentUser.UserPasarNivel, CurrentUser.UserPercExp & "%")

If CurrentUser.UserPasarNivel = 0 Then
    frmMain.Label2(1).Caption = Locale_GUI_Frase(173)
End If

End Sub

Private Sub ActualizarHAM()

frmMain.shpSed.Width = (((CurrentUser.UserMinAGU / 100) / (CurrentUser.UserMaxAGU / 100)) * 89)
frmMain.shpHambre.Width = (((CurrentUser.UserMinHAM / 100) / (CurrentUser.UserMaxHAM / 100)) * 88)
frmMain.lblHambre.Caption = CurrentUser.UserMinHAM & "/" & CurrentUser.UserMaxHAM
frmMain.lblSed.Caption = CurrentUser.UserMinAGU & "/" & CurrentUser.UserMaxAGU

End Sub

