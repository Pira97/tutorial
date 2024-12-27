Attribute VB_Name = "Cuentas"
'---------------------------------------------------------------------------------------
' Module    : Cuentas
' Author    : Shermie80
' Date      : 10/03/2015
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public AccountPath As String

Public Type AccountUser
    Name As String
    Head As Integer
    body As Integer
    casco As Integer
    weapon As Integer
    shield As Integer
    nivel As Byte
    Clase As Byte
    Mapa As Integer
    color As Byte
    gameMaster As Boolean
End Type

Public Function AddUserInAccount(ByVal Name As String, ByVal Account As String)
    
    Dim aFile As String
    
    aFile = AccountPath & Account & ".cnt"
    
    Dim NumPJS As Byte
    
    NumPJS = CByte(GetVar(aFile, "PJS", "NumPjs")) + 1
    
    Call WriteVar(aFile, "PJS", "PJ" & NumPJS, Name)
    Call WriteVar(aFile, "PJS", "NumPjs", NumPJS)


End Function
Public Function IsPjOfAccount(ByVal Name As String, ByVal Account As String) As Boolean

    On Error GoTo ErrorHandler
      
    Dim aFile As String
    
    aFile = AccountPath & Account & ".cnt"
    
    Dim NumPJS As Byte
    NumPJS = CByte(GetVar(aFile, "PJS", "NumPjs"))
    
    If Not NumPJS = 0 Then
        Dim i As Byte
        For i = 1 To NumPJS
            If UCase$(Name) = UCase$(GetVar(aFile, "PJS", "PJ" & i)) Then
                IsPjOfAccount = True
                Exit Function
            End If
        Next i
    End If
    
    IsPjOfAccount = False
    
103 Exit Function

ErrorHandler:
104     Call RegistrarError(Err.Number, Err.description, "Cuentas.IsPJOfAccount", Erl)
106     Resume Next
End Function

Public Sub SaveNewAccount(ByVal UserIndex As Integer, ByVal Cuenta As String, ByVal Password As String, ByVal UserCode As String)
        
        On Error GoTo SaveNewAccount_Err

        Dim Salt As String * 10: Salt = RandomString(10) ' Alfanumerico
    
        Dim oSHA256 As CSHA256

102     Set oSHA256 = New CSHA256

        Dim PasswordHash As String * 64: PasswordHash = oSHA256.SHA256(Password & Salt)
        Dim UserCodeHash As String * 64: UserCodeHash = oSHA256.SHA256(UserCode & Salt)
    
106     Set oSHA256 = Nothing

114     Call SaveNewAccountCharfile(Cuenta, PasswordHash, Salt, UserCodeHash)
            
        Call WriteAbrirFormularios(UserIndex, 1) 'Show Account
        Call WriteShowMessageBox(UserIndex, 32) 'Cuenta creada con exito
        
        Exit Sub

SaveNewAccount_Err:
116     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveNewAccount", Erl)
118     Resume Next
        
End Sub
     
Public Sub SaveNewAccountCharfile(Cuenta As String, PasswordHash As String, Salt As String, Codigo As String)

        On Error GoTo ErrorHandler

        Dim n As Integer
        n = FreeFile()
        
        Open AccountPath & Cuenta & ".cnt" For Append As n
        
            Print #n, "[" & Cuenta & "]"
            Print #n, "Cuenta=" & Cuenta
            Print #n, "Password=" & PasswordHash
            Print #n, "Salt=" & Salt
            Print #n, "Ban=0"
            Print #n, "UserCodigo=" & Codigo
            Print #n, "Conectada=0"
            Print #n, "CuentaGM=0"
            Print #n, "Donador=0"
            Print #n, "MacAdress=0"
            Print #n, "HDserial=0"
            Print #n, "[PJS]"
            Print #n, "NumPjs=0"
            Print #n, "PJ1="
            Print #n, "PJ2="
            Print #n, "PJ3="
            Print #n, "PJ4="
            Print #n, "PJ5="
            Print #n, "PJ6="
            Print #n, "PJ7="
            Print #n, "PJ8="
            Print #n, "PJ9="
            Print #n, "PJ10="
            
        Close #n
    Exit Sub

ErrorHandler:
104     Call RegistrarError(Err.Number, Err.description, "Cuentas.SaveNewAccountCharfile", Erl)
106     Resume Next
End Sub

Public Sub LoginAccountCharfile(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler

    
2        Dim i                  As Long
3        Dim NumberOfCharacters As Byte
4        Dim Characters()       As AccountUser
5        Dim CurrentCharacter   As String
6        Dim Cuenta As String
        
7        Cuenta = UserList(UserIndex).Account
    
9        NumberOfCharacters = val(GetVar(AccountPath & UCase$(Cuenta) & ".cnt", "PJS", "NumPjs"))

         If NumberOfCharacters > 0 Then
         
             ReDim Characters(1 To NumberOfCharacters) As AccountUser
             
10           For i = 1 To NumberOfCharacters
        
11                CurrentCharacter = GetVar(AccountPath & UCase$(Cuenta) & ".cnt", "PJS", "PJ" & i)
    
13                Characters(i).Name = CurrentCharacter
14                Characters(i).Head = ObtenerCabeza(CurrentCharacter)
15                Characters(i).body = ObtenerCuerpo(CurrentCharacter)
16                Characters(i).casco = ObtenerCasco(CurrentCharacter)
17                Characters(i).weapon = ObtenerArma(CurrentCharacter)
18                Characters(i).shield = ObtenerEscudo(CurrentCharacter)
19                Characters(i).Mapa = ReadField(1, ObtenerMapa(CurrentCharacter), Asc("-"))
20                Characters(i).Clase = ObtenerClase(CurrentCharacter)
21                Characters(i).nivel = ObtenerNivel(CurrentCharacter)
22                Characters(i).color = ObtenerColor(CurrentCharacter)
23                Characters(i).gameMaster = EsGmChar(CurrentCharacter)
                
24           Next i
    
25           Call WriteAddPj(UserIndex, Cuenta, NumberOfCharacters, Characters)

        
        Else
            Call WriteEjecutarAccion(UserIndex, 6)
        
        End If
        
        Call WriteAbrirFormularios(UserIndex, 1) 'ShowAccount
        
44      Exit Sub

ErrorHandler:
29    Call RegistrarError(Err.Number, Err.description, "ModCuentas.LoginAccountCharfile", Erl)
      Resume Next
End Sub
Public Function ObtenerColor(ByVal Name As String) As Byte

        On Error GoTo ErrorHandler

100     ObtenerColor = GetVar(CharPath & UCase$(Name & ".chr"), "FACCIONES", "Status")
    
        Exit Function
ErrorHandler:
102     ObtenerColor = "1"

End Function
Public Function ObtenerClase(ByVal Name As String) As Byte

        On Error GoTo ErrorHandler

100     ObtenerClase = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Clase")

        Exit Function
ErrorHandler:
102     ObtenerClase = "1"

End Function

Public Function ObtenerNivel(ByVal Name As String) As Byte

        On Error GoTo ErrorHandler

100     ObtenerNivel = GetVar(CharPath & UCase$(Name & ".chr"), "STATS", "ELV")

        Exit Function
ErrorHandler:
102     ObtenerNivel = 1

End Function
Public Function ObtenerMapa(ByVal Name As String) As String

        On Error GoTo ErrorHandler

        Dim Mapa As String

100     ObtenerMapa = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Position")
    
        Exit Function
ErrorHandler:
102     ObtenerMapa = "1-50-50"
    
End Function
Public Function ObtenerCabeza(ByVal Name As String) As Integer

        On Error GoTo ErrorHandler

        Dim Head       As String

        Dim EstaMuerto As Byte

100     EstaMuerto = GetVar(CharPath & UCase$(Name & ".chr"), "FLAGS", "Muerto")

102     If EstaMuerto = 0 Then
104         Head = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Head")
        Else
106         Head = iCabezaMuerto
        End If

108     ObtenerCabeza = Head

        Exit Function
ErrorHandler:
110     ObtenerCabeza = 1

End Function

Public Function ObtenerCuerpo(ByVal Name As String) As Integer

        On Error GoTo ErrorHandler

        Dim EstaMuerto As Byte

        Dim cuerpo     As Long

100     EstaMuerto = GetVar(CharPath & UCase$(Name & ".chr"), "flags", "Muerto")

102     If EstaMuerto = 0 Then
104         cuerpo = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Body")
106         ObtenerCuerpo = cuerpo
        Else
108         ObtenerCuerpo = iCuerpoMuerto

        End If

        Exit Function
ErrorHandler:
110     ObtenerCuerpo = 1

End Function

Public Function ObtenerCasco(ByVal Name As String) As Integer

        On Error GoTo ErrorHandler

100     ObtenerCasco = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Casco")
        Exit Function
ErrorHandler:
102     ObtenerCasco = 0

End Function

Public Function ObtenerArma(ByVal Name As String) As Integer

        On Error GoTo ErrorHandler

100     ObtenerArma = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Arma")
        Exit Function
ErrorHandler:
102     ObtenerArma = 0

End Function

Public Function ObtenerEscudo(ByVal Name As String) As Integer

        On Error GoTo ErrorHandler

100     ObtenerEscudo = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Escudo")
        Exit Function
ErrorHandler:
102     ObtenerEscudo = 0

End Function


Public Function CuentaExiste(ByVal Cuenta As String) As Boolean
        
        On Error GoTo CuentaExiste_Err

            CuentaExiste = FileExist(AccountPath & UCase$(Cuenta) & ".cnt", vbNormal)
        
6        Exit Function

CuentaExiste_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.CuentaExiste", Erl)
108     Resume Next
        
End Function
Public Function ChequeosServerIni(ByVal UserIndex As Integer, ByVal UserName As String, ByVal UserAccount As String, ByRef MensajeAdvertencia As Integer, ByRef CierraConexion As Boolean) As Boolean
    
    On Error GoTo ErrorHandler
    
    '//Mermas 27/08/2021
    'Acá hacemos los chequeos generales del servidor, para no tener que copiarlo 20 veces :P
    
    ChequeosServerIni = False
    MensajeAdvertencia = 36
    
    
    'Server solo GMs
    If ServerSoloGMs <> 0 Then
    
        If UserAccount = "" Then
            
            If Not EsGmChar(UserName) Then
                MensajeAdvertencia = 70  'El servidor se encuentra restringido para Game Master. Por favor intente más tarde.
                CierraConexion = True
                Exit Function
            End If
            
        ElseIf UserName = "" Then
            
            If Not EsGmAccount(UserAccount) Then
                MensajeAdvertencia = 70 'El servidor se encuentra restringido para Game Master. Por favor intente más tarde.
                CierraConexion = True
                Exit Function
            End If
            
        End If
    End If
    
    '¿Este IP ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameIP(UserIndex, UserList(UserIndex).ip) = True Then
            MensajeAdvertencia = 25 ' Su cuenta ha alcanzado el límite de conexiones simultáneas permitidas para este servidor.
            CierraConexion = True
            Exit Function
        End If
    End If
    
    'Si puede crear PJs
    If PuedeCrearPersonajes = 0 Then
        MensajeAdvertencia = 3 ' La creación de personajes está prohibida temporalmente. Intente más tarde.
        CierraConexion = True
        Exit Function
    End If
 
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        MensajeAdvertencia = 15 'El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intentarlo mas tarde.
        CierraConexion = True
        Exit Function
    End If
 
 
    If EnPausa Then
        MensajeAdvertencia = 2 'El servidor se encuentra restringido por mantenimiento. Por favor intente más tarde.
        CierraConexion = False
        Exit Function
    End If
    
    ChequeosServerIni = True

    Exit Function

ErrorHandler:
        MensajeAdvertencia = 36 'La acción no puede ser realizada debido a un error, consulte con algún Admnistrador para regularizar la situación.
        CierraConexion = True
        ChequeosServerIni = False
116     Call RegistrarError(Err.Number, Err.description, "General.ChequeosServerIni", Erl)
118     Resume Next
        
End Function


Function EntrarCuenta(ByVal UserIndex As Integer, Cuenta As String, CuentaPassword As String, MacAddress As String, ByVal HDserial As Long) As Boolean
        
        On Error GoTo EntrarCuenta_Err
        
        EntrarCuenta = False
        
'100     If CheckMAC(MacAddress) Then
'102         Call WriteShowMessageBox(UserIndex, 71) 'Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0001
'            Exit Function
'        End If
'
'104     If CheckHD(HDserial) Then
'106         Call WriteShowMessageBox(UserIndex, 42) 'Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0002
'            Exit Function
'        End If

        'Existe ya la cuenta?
        If CuentaExiste(Cuenta) Then
        
             If Not ObtenerBaneo(Cuenta) Then

                Dim PasswordHash As String, Salt As String
                
120             PasswordHash = GetVar(AccountPath & UCase$(Cuenta) & ".cnt", Cuenta, "Password")
122             Salt = GetVar(AccountPath & UCase$(Cuenta) & ".cnt", Cuenta, "Salt")

124              If PasswordValida(SDesencriptar(CuentaPassword), PasswordHash, Salt) Then

128                  Call WriteVar(AccountPath & LCase$(Cuenta) & ".cnt", Cuenta, "MacAdress", MacAddress)
130                 Call WriteVar(AccountPath & LCase$(Cuenta) & ".cnt", Cuenta, "HDserial", HDserial)
132                 Call WriteVar(AccountPath & LCase$(Cuenta) & ".cnt", Cuenta, "UltimoAcceso", Date & " " & Time)
134                 Call WriteVar(AccountPath & LCase$(Cuenta) & ".cnt", Cuenta, "UltimaIP", UserList(UserIndex).ip)

136                 UserList(UserIndex).Account = Cuenta

138                 EntrarCuenta = True

                Else
                    
142                  Call WriteShowMessageBox(UserIndex, 29) 'Contraseña incorrecta

                End If

             Else
             
                Call WriteShowMessageBox(UserIndex, 19) 'Su cuenta está baneada. Si crees que esto es un erorr, comunicate con algún Administrador para regularizr tu situación
            
             End If
            
        Else
        
            Call WriteShowMessageBox(UserIndex, 39) 'La cuenta no existe.
            
        End If

        
        Exit Function

EntrarCuenta_Err:
148     Call RegistrarError(Err.Number, Err.description, "TCP.EntrarCuenta", Erl)
150     Resume Next
        
End Function

Public Function ObtenerBaneo(ByVal Name As String) As Boolean
        
        On Error GoTo ObtenerBaneo_Err
        
104     ObtenerBaneo = val(GetVar(AccountPath & UCase$(Name) & ".cnt", Name, "Ban")) = 1
        
        Exit Function

ObtenerBaneo_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerBaneo", Erl)
108     Resume Next
        
End Function


Public Function PasswordValida(Password As String, PasswordHash As String, Salt As String) As Boolean
        
        On Error GoTo PasswordValida_Err
        

        Dim oSHA256 As CSHA256

100     Set oSHA256 = New CSHA256

102     PasswordValida = (PasswordHash = oSHA256.SHA256(Password & Salt))
    
104     Set oSHA256 = Nothing

        
        Exit Function

PasswordValida_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.PasswordValida", Erl)
108     Resume Next
        
End Function



Public Function PinValido(PinGM As String, PinHash As String, Salt As String) As Boolean
        
        On Error GoTo PinValido
        

        Dim oSHA256 As CSHA256

100     Set oSHA256 = New CSHA256

102     PinValido = (PinHash = oSHA256.SHA256(PinGM & Salt))
    
104     Set oSHA256 = Nothing

        
        Exit Function

PinValido:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.PinValido", Erl)
108     Resume Next
        
End Function



Public Function ObtenerCantidadDePersonajesByUserIndex(ByVal UserIndex As Integer) As Byte
        
    On Error GoTo ObtenerCantidadDePersonajesByUserIndex_Err

    ObtenerCantidadDePersonajesByUserIndex = val(GetVar(AccountPath & UCase$(UserList(UserIndex).Name) & ".cnt", "PJS", "NumPjs"))

    Exit Function

ObtenerCantidadDePersonajesByUserIndex_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerCantidadDePersonajesByUserIndex", Erl)
108     Resume Next
        
End Function

Public Function CheckDataNewAccount(ByVal Cuenta As String, ByVal Password As String, ByVal UserCode As String, Optional ByRef MensajeAdvertencia As Integer = 0) As Boolean
    
    On Error GoTo CheckDataNewAccount_Err
    
    Dim loopc As Integer
    Dim CharAscii As Integer
    
    MensajeAdvertencia = 36
    
    If Len(Cuenta) = 0 Then
        MensajeAdvertencia = 62  '  'Ingrese un nombre valido
        CheckDataNewAccount = False
        Exit Function
    End If
    
    If Len(Cuenta) > 20 Then
        MensajeAdvertencia = 63 ' Menos de 20 letras
        CheckDataNewAccount = False
        Exit Function
    End If
 
    For loopc = 1 To Len(Cuenta)
        CharAscii = Asc(mid$(Cuenta, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia = 64  'Caracteres invalidos
            CheckDataNewAccount = False
            Exit Function
        End If
    Next loopc
        
    If Len(Password) = 0 Then
        MensajeAdvertencia = 65 'Ingrese una contraseña valida
        CheckDataNewAccount = False
        Exit Function
    End If
    
    If Len(Password) > 30 Then
        MensajeAdvertencia = 66 ' ) ' contra no mayor de 30
        CheckDataNewAccount = False
        Exit Function
    End If
 
    For loopc = 1 To Len(Password)
        CharAscii = Asc(mid$(Password, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia = 69  'Caracteres invalidos
            CheckDataNewAccount = False
            Exit Function
        End If
    Next loopc
    
    If Len(UserCode) = 0 Then
        MensajeAdvertencia = 67 '¡Ingrese un código válido!
        CheckDataNewAccount = False
        Exit Function
    End If
 
    'No permitimos letras en el pin
    If Not IsNumeric(UserCode) Then
        MensajeAdvertencia = 53  ' Locale_Error(53)  '"Solo se permite usar numeros en el código de cuenta."
        CheckDataNewAccount = False
        Exit Function
    End If

    If Len(UserCode) > 4 Then
        MensajeAdvertencia = 68  ' code no mayor a 4
        CheckDataNewAccount = False
        Exit Function
    End If
    
    CheckDataNewAccount = True
    
    Exit Function
    
CheckDataNewAccount_Err:
        CheckDataNewAccount = False
        MensajeAdvertencia = 36
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.CheckDataNewAccount", Erl)
108     Resume Next
        
End Function


Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    'if backspace allow

    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters

    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function

    End If
    
    If KeyAscii > 126 Then
        Exit Function

    End If
    
    'Check for bad special characters in between

    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii _
            = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

End Function


Public Function CheckDataLoginAccount(ByVal Cuenta As String, ByVal Password As String, Optional ByRef MensajeAdvertencia As Integer = 0) As Boolean
    
    On Error GoTo CheckDataLoginAccount_err
    
    Dim loopc As Integer
    Dim CharAscii As Integer
    
    MensajeAdvertencia = 36
 
    If Len(Cuenta) = 0 Then
        MensajeAdvertencia = 62
        CheckDataLoginAccount = False
        Exit Function
    End If
    
    If Len(Cuenta) > 20 Then
        MensajeAdvertencia = 63
        CheckDataLoginAccount = False
        Exit Function
    End If
    
    For loopc = 1 To Len(Cuenta)
        CharAscii = Asc(mid$(Cuenta, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia = 64
            CheckDataLoginAccount = False
            Exit Function
        End If
    Next loopc
 
    
    If Len(Password) = 0 Then
        MensajeAdvertencia = 65 'Ingrese una contraseña valida
        CheckDataLoginAccount = False
        Exit Function
    End If

    If Len(Password) > 30 Then
        MensajeAdvertencia = 66
        CheckDataLoginAccount = False
        Exit Function
    End If
    
    For loopc = 1 To Len(Password)
        CharAscii = Asc(mid$(Password, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia = 69
            CheckDataLoginAccount = False
            Exit Function
        End If
    Next loopc

    
    CheckDataLoginAccount = True
    
    Exit Function
    
CheckDataLoginAccount_err:
        CheckDataLoginAccount = False
        MensajeAdvertencia = 36
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.CheckDataLoginAccount", Erl)
108     Resume Next
        
End Function

Public Function CuentaConectada(ByVal Name As String) As Byte
        
        On Error GoTo CuentaConectada_Err
        
        Name = UCase$(Name)
104     CuentaConectada = GetVar(AccountPath & Name & ".cnt", Name, "Conectada")

        Exit Function

CuentaConectada_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.CuentaConectada", Erl)
108     Resume Next
        
End Function

