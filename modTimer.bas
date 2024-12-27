Attribute VB_Name = "modTimer"
Option Explicit

'Intervalos
Private Const CONST_INTERVALO_Conectar    As Long = 1000
Private Const CONST_INTERVALO_AutoUsar As Long = 100
Private Const CONST_INTERVALO_HEADING     As Long = 120

Private HFormularios As Long

Private CloseForm As Integer

Private hFXTimer As Long
Private hMinutoTimer As Long
Public Intervalos As tIntervalos

Public Type tIntervalos
    Conectar As Long
    AutoUsar As Long
    Heading As Long
End Type
Public Sub FormularioTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 340)
    
    On Error GoTo ErrorHandler_err
    
    If Enabled Then
        If HFormularios <> 0 Then KillTimer 0, HFormularios
        HFormularios = SetTimer(0, 0, Intervalo, AddressOf FormularioTimerProc)
    Else
        If HFormularios = 0 Then Exit Sub
        KillTimer 0, HFormularios
        HFormularios = 0
    End If

    Exit Sub

ErrorHandler_err:
    Call RegistrarError(Err.number, Err.Description, "modTimer.FormularioTimer", Erl)
    Resume Next
    
End Sub
Private Sub FormularioTimerProc()

    On Error GoTo error_Err
    
    If Connected Then
    
        
        'Unload the connect form
        Unload frmCharList
        Unload frmCrearPersonaje
        Unload frmConnect
        Unload frmCrearCuenta
        Unload frmRecuperarCuenta
        Unload frmMensaje
        Unload frmPregunta
        
        Call FormularioTimer(False)
    
    End If
    
    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "modTimer.FormularioTimerProc", Erl)
    Resume Next
    
End Sub
Public Sub FXTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 10000)
    
    On Error GoTo ErrorHandler_err
    
    If Enabled Then
        If hFXTimer <> 0 Then KillTimer 0, hFXTimer
        hFXTimer = SetTimer(0, 0, Intervalo, AddressOf FXTimerProc)
    Else
        If hFXTimer = 0 Then Exit Sub
        KillTimer 0, hFXTimer
        hFXTimer = 0
    End If

    Exit Sub

ErrorHandler_err:
    Call RegistrarError(Err.number, Err.Description, "modTimer.FXTimer", Erl)
    Resume Next
    
End Sub
Private Sub FXTimerProc()

    On Error GoTo error_Err
    If Connected Then
    
        Dim Horario() As String
        Horario() = Split(Time, ":")
 
        If tHora <> CByte(Horario(0)) Then
            Meteo_Change_Time
        End If
        
    
        tHora = CByte(Horario(0))
        tMinuto = CByte(Horario(1))
        
    End If
    
    Exit Sub

error_Err:
    Call RegistrarError(Err.number, Err.Description, "modTimer.FXTimerProc", Erl)
    Resume Next
    
End Sub


Public Function IntervaloPermiteHeading(Optional ByVal Actualizar As Boolean = True) As Boolean
    
    On Error GoTo IntervaloPermiteHeading_Err
    
    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - Intervalos.Heading >= CONST_INTERVALO_HEADING Then
    
        If Actualizar Then
            Intervalos.Heading = TActual
        End If

        IntervaloPermiteHeading = True
    Else
        IntervaloPermiteHeading = False
    End If

    
    Exit Function

IntervaloPermiteHeading_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteHeading", Erl)
    Resume Next
    
End Function


Public Function IntervaloPermiteConectar() As Boolean
    
    On Error GoTo IntervaloPermiteConectar_Err
    
    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - Intervalos.Conectar >= CONST_INTERVALO_Conectar Then
        
        ' Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Intervalos.Conectar = TActual
        IntervaloPermiteConectar = True
        Call FormParser.Parse_Form(frmConnect, e_normal)
    Else
        If frmCharList.Visible = True Then
            frmMensaje.msg.Caption = Locale_Error(43)
            frmMensaje.Show vbModal, frmConnect
            
        ElseIf frmConnect.Visible = True Then
            frmMensaje.msg.Caption = Locale_Error(43)
            frmMensaje.Show vbModal, frmConnect
        ElseIf frmCrearPersonaje.Visible = True Then
            frmCrearPersonaje.lblInfo.Caption = Locale_Error(43)

            
        End If
        
        IntervaloPermiteConectar = False

    End If

    
    Exit Function

IntervaloPermiteConectar_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteConectar", Erl)
    Resume Next
    
End Function

Public Function IntervaloIntervaloAutoUsar() As Boolean
    
    On Error GoTo IntervaloAutoUsar_Err
    
    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - Intervalos.AutoUsar >= CONST_INTERVALO_AutoUsar Then
        
        ' Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Intervalos.AutoUsar = TActual
        IntervaloIntervaloAutoUsar = True
    Else
        IntervaloIntervaloAutoUsar = True
    End If

    
    Exit Function

IntervaloAutoUsar_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloAutoUsar", Erl)
    Resume Next
    
End Function
  
Public Sub MinutoTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 1000)

    On Error GoTo ErrorHandler_err
    
    If Enabled Then
        If hMinutoTimer <> 0 Then KillTimer 0, hMinutoTimer
        hMinutoTimer = SetTimer(0, 0, Intervalo, AddressOf MinutoTimerProc)
    Else
        If hMinutoTimer = 0 Then Exit Sub
        KillTimer 0, hMinutoTimer
        hMinutoTimer = 0
    End If
    
    Exit Sub

ErrorHandler_err:
    Call RegistrarError(Err.number, Err.Description, "modTimer.MinutoTimer", Erl)
    Resume Next
    
End Sub

Private Sub MinutoTimerProc()

    On Error GoTo ErrorHandler_err
    
    Dim N As Long
    
    If IsAppActive Then
        If frmMain.Visible Then
            If CurrentUser.Logged Then
                
                Static MostrarMS As Long, ExternosSeg As Long
                
                MostrarMS = MostrarMS + 1
                ExternosSeg = ExternosSeg + 1
                
                'If ExternosSeg >= 2 Then
                
                    'Call Externos
                    'ExternosSeg = 0
                    
                'End If
    
                If MostrarMS >= 10 Then
                    If FPSFLAG Then
                        If pausa Then Exit Sub
                        Call WritePing
                    End If
                    
                    MostrarMS = 0
                End If
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler_err:
    Call RegistrarError(Err.number, Err.Description, "modTimer.MiinutoTimerProc", Erl)
    Resume Next
    
End Sub

