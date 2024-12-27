Attribute VB_Name = "wskapiAO"
'**************************************************************
' wskapiAO.bas
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

''
' Modulo para manejar Winsock
'
 
    'Si la variable esta en TRUE , al iniciar el WsApi se crea
    'una ventana LABEL para recibir los mensajes. Al detenerlo,
    'se destruye.
    'Si es FALSE, los mensajes se envian al form frmMain (o el
    'que sea).
    #Const WSAPI_CREAR_LABEL = True

    Private Const SD_BOTH As Long = &H2

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Public Declare Function GetWindowLong _
                   Lib "user32" _
                   Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                           ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong _
                   Lib "user32" _
                   Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                           ByVal nIndex As Long, _
                                           ByVal dwNewLong As Long) As Long
    Public Declare Function CallWindowProc _
                   Lib "user32" _
                   Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                            ByVal hwnd As Long, _
                                            ByVal msg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long) As Long

    Private Declare Function CreateWindowEx _
                    Lib "user32" _
                    Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                             ByVal lpClassName As String, _
                                             ByVal lpWindowName As String, _
                                             ByVal dwStyle As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hwndParent As Long, _
                                             ByVal hMenu As Long, _
                                             ByVal hInstance As Long, _
                                             lpParam As Any) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

    Private Const WS_CHILD = &H40000000
    Public Const GWL_WNDPROC = (-4)

    Private Const SIZE_RCVBUF As Long = 8192
    Private Const SIZE_SNDBUF As Long = 8192

    ''
    'Esto es para agilizar la busqueda del slot a partir de un socket dado,
    'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
    '
    ' @param Sock sock
    ' @param slot slot
    '
    Public Type tSockCache

        Sock As Long
        slot As Long

    End Type

    Public WSAPISock2Usr As New Collection

    ' ====================================================================================
    ' ====================================================================================

    Public OldWProc      As Long
    Public ActualWProc   As Long
    Public hWndMsg       As Long

    ' ====================================================================================
    ' ====================================================================================

    Public SockListen     As Long
    Public LastSockListen As Long
 

' ====================================================================================
' ====================================================================================

Public Sub IniciaWsApi(ByVal hwndParent As Long)
 
        Call LogApiSock("IniciaWsApi")
        Debug.Print "IniciaWsApi"

        #If WSAPI_CREAR_LABEL Then
            hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
        #Else
            hWndMsg = hwndParent
        #End If 'WSAPI_CREAR_LABEL

        OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
        ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

        Dim desc As String
        Call StartWinsock(desc)

 
End Sub

Public Sub LimpiaWsApi()
 
        Call LogApiSock("LimpiaWsApi")

        If WSAStartedUp Then
            Call EndWinsock

        End If

        If OldWProc <> 0 Then
            SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
            OldWProc = 0

        End If

        #If WSAPI_CREAR_LABEL Then

            If hWndMsg <> 0 Then
                DestroyWindow hWndMsg

            End If

        #End If

 
End Sub

Public Function BuscaSlotSock(ByVal S As Long) As Long
 
        On Error GoTo hayerror
    
        BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
        Exit Function
    
hayerror:
        BuscaSlotSock = -1
 
End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal slot As Long)
    Debug.Print "AgregaSockSlot" & " " & Time
 
        If WSAPISock2Usr.count > MaxUsers Then
            Call CloseSocket(slot)
            Exit Sub

        End If

        WSAPISock2Usr.Add CStr(slot), CStr(Sock)

        'Dim Pri As Long, Ult As Long, Med As Long
        'Dim LoopC As Long
        '
        'If WSAPISockChacheCant > 0 Then
        '    Pri = 1
        '    Ult = WSAPISockChacheCant
        '    Med = Int((Pri + Ult) / 2)
        '
        '    Do While (Pri <= Ult) And (Ult > 1)
        '        If Sock < WSAPISockChache(Med).Sock Then
        '            Ult = Med - 1
        '        Else
        '            Pri = Med + 1
        '        End If
        '        Med = Int((Pri + Ult) / 2)
        '    Loop
        '
        '    Pri = IIf(Sock < WSAPISockChache(Med).Sock, Med, Med + 1)
        '    Ult = WSAPISockChacheCant
        '    For LoopC = Ult To Pri Step -1
        '        WSAPISockChache(LoopC + 1) = WSAPISockChache(LoopC)
        '    Next LoopC
        '    Med = Pri
        'Else
        '    Med = 1
        'End If
        'WSAPISockChache(Med).Slot = Slot
        'WSAPISockChache(Med).Sock = Sock
        'WSAPISockChacheCant = WSAPISockChacheCant + 1

 
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long)
         Dim cant As Long

        cant = WSAPISock2Usr.count

        On Error Resume Next

        WSAPISock2Usr.Remove CStr(Sock)

        Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.count

 
End Sub

Public Function WndProc(ByVal hwnd As Long, _
                        ByVal msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
 
        On Error Resume Next

        Dim Ret      As Long
        Dim Tmp()    As Byte
        Dim S        As Long
        Dim e        As Long
        Dim n        As Integer
        Dim UltError As Long
    
        Select Case msg

            Case 1025
                S = wParam
                e = WSAGetSelectEvent(lParam)
            
                Select Case e

                    Case FD_ACCEPT

                        If S = SockListen Then
                            Call EventoSockAccept(S)

                        End If
                
                        '    Case FD_WRITE
                        '        N = BuscaSlotSock(s)
                        '        If N < 0 And s <> SockListen Then
                        '            'Call apiclosesocket(s)
                        '            call WSApiCloseSocket(s)
                        '            Exit Function
                        '        End If
                        '
            
                        '        Call IntentarEnviarDatosEncolados(N)
                        '
                        '        Dale = UserList(N).ColaSalida.Count > 0
                        '        Do While Dale
                        '            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
                        '            If Ret <> 0 Then
                        '                If Ret = WSAEWOULDBLOCK Then
                        '                    Dale = False
                        '                Else
                        '                    'y aca que hacemo' ?? help! i need somebody, help!
                        '                    Dale = False
                        '                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
                        '                End If
                        '            Else
                        '            '    Debug.Print "Dato de la cola enviado"
                        '                UserList(N).ColaSalida.Remove 1
                        '                Dale = (UserList(N).ColaSalida.Count > 0)
                        '            End If
                        '        Loop
        
                    Case FD_READ
                        n = BuscaSlotSock(S)

                        If n < 0 And S <> SockListen Then
                            'Call apiclosesocket(s)
                            Call WSApiCloseSocket(S)
                            Exit Function

                        End If
                    
                        'create appropiate sized buffer
                        ReDim Preserve Tmp(SIZE_RCVBUF - 1) As Byte
                    
                        Ret = recv(S, Tmp(0), SIZE_RCVBUF, 0)

                        ' Comparo por = 0 ya que esto es cuando se cierra
                        ' "gracefully". (mas abajo)
                        If Ret < 0 Then
                            UltError = Err.LastDllError

                            If UltError = WSAEMSGSIZE Then
                                Debug.Print "WSAEMSGSIZE"
                                Ret = SIZE_RCVBUF
                            Else
                                Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                                Call LogApiSock("Error en Recv: N=" & n & " S=" & S & " Str=" & GetWSAErrorString( _
                                        UltError))
                            
                                'no hay q llamar a CloseSocket() directamente,
                                'ya q pueden abusar de algun error para
                                'desconectarse sin los 10segs. CREEME.
                                Call CloseSocketSL(n)
                                Call Cerrar_Usuario(n)
                                Exit Function

                            End If

                        ElseIf Ret = 0 Then
                            Call CloseSocketSL(n)
                            Call Cerrar_Usuario(n)

                        End If
                    
                        ReDim Preserve Tmp(Ret - 1) As Byte
                    
                        Call EventoSockRead(n, Tmp)
                
                    Case FD_CLOSE
                        n = BuscaSlotSock(S)

                        If S <> SockListen Then Call apiclosesocket(S)
                    
                        If n > 0 Then
                            Call BorraSlotSock(S)
                            UserList(n).ConnID = -1
                            UserList(n).ConnIDValida = False
                            Call EventoSockClose(n)

                        End If

                End Select
        
            Case Else
                WndProc = CallWindowProc(OldWProc, hwnd, msg, wParam, lParam)

        End Select

 
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal slot As Integer, ByRef str As String) As Long
         Dim Ret     As String
        Dim Retorno As Long
        Dim data()  As Byte
    
        ReDim Preserve data(Len(str) - 1) As Byte

        data = StrConv(str, vbFromUnicode)

        Retorno = 0
    
        If UserList(slot).ConnID <> -1 And UserList(slot).ConnIDValida Then
            Ret = send(ByVal UserList(slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)

            If Ret < 0 Then
                Ret = Err.LastDllError

                If Ret = WSAEWOULDBLOCK Then
 
 
                
                    ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
                    Call UserList(slot).outgoingData.WriteASCIIStringFixed(str)

                End If

            End If

        ElseIf UserList(slot).ConnID <> -1 And Not UserList(slot).ConnIDValida Then

            If Not UserList(slot).Counters.Saliendo Then
                Retorno = -1

            End If

        End If
    
        WsApiEnviar = Retorno
 
End Function

Public Sub LogApiSock(ByVal str As String)
 
        On Error GoTo ErrHandler

        Dim nfile As Integer
        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & str
        Close #nfile

        Exit Sub

ErrHandler:

 
End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
         '==========================================================
        'USO DE LA API DE WINSOCK
        '========================
    
        Dim NewIndex  As Integer
        Dim Ret       As Long
        Dim Tam       As Long, sa As sockaddr
        Dim NuevoSock As Long
        Dim i         As Long
        Dim tStr      As String
    
        Tam = sockaddr_size
    
        '=============================================
        'SockID es en este caso es el socket de escucha,
        'a diferencia de socketwrench que es el nuevo
        'socket de la nueva conn
    
        'Modificado por Maraxus
        'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
        Ret = accept(SockID, sa, Tam)

        If Ret = INVALID_SOCKET Then
            i = Err.LastDllError
            Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
            Exit Sub

        End If
    
        If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
            Call WSApiCloseSocket(NuevoSock)
            Exit Sub

        End If

        'If Ret = INVALID_SOCKET Then
        '    If Err.LastDllError = 11002 Then
        '        ' We couldn't decide if to accept or reject the connection
        '        'Force reject so we can get it out of the queue
        '        Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
        '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
        '    Else
        '        i = Err.LastDllError
        '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
        '        Exit Sub
        '    End If
        'End If

        NuevoSock = Ret
        
        
                 If setsockopt(NuevoSock, 6, TCP_NODELAY, True, 4) <> 0 Then
             i = Err.LastDllError
             Call LogCriticEvent("Error al setear el delay " & i & ": " & GetWSAErrorString(i))

            End If
            
        'Seteamos el tamaño del buffer de entrada
        If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
            i = Err.LastDllError
            Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))

        End If

        'Seteamos el tamaño del buffer de salida
        If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
            i = Err.LastDllError
            Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))

        End If

        'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
        'tStr = "Limite de conexiones para su IP alcanzado."
        'Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
        'Call WSApiCloseSocket(NuevoSock)
        'Exit Sub
        'End If
    
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   BIENVENIDO AL SERVIDOR!!!!!!!!
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
        NewIndex = NextOpenUser ' Nuevo indice
    
        If NewIndex <= MaxUsers Then
        
            'Make sure both outgoing and incoming data buffers are clean
            Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.length)
            Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.length)


        
            UserList(NewIndex).ip = GetAscIP(sa.sin_addr)

            'Busca si esta banneada la ip
            For i = 1 To BanIps.count

                If BanIps.Item(i) = UserList(NewIndex).ip Then
                    'Call apiclosesocket(NuevoSock)
                    Call WriteSendMsgBox(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                    Call FlushBuffer(NewIndex)
                    'Call SecurityIp.IpRestarConexion(sa.sin_addr)
                    Call WSApiCloseSocket(NuevoSock)
                    Exit Sub

                End If

            Next i
         If aDos.MaxConexiones(UserList(NewIndex).ip) Then
        UserList(NewIndex).ConnID = -1
        Call aDos.RestarConexion(UserList(NewIndex).ip)
        Call apiclosesocket(NuevoSock)
    End If
            If NewIndex > LastUser Then LastUser = NewIndex
        
            UserList(NewIndex).ConnID = NuevoSock
            UserList(NewIndex).ConnIDValida = True
        
            Call AgregaSlotSock(NuevoSock, NewIndex)
        Else
            Dim str    As String
            Dim data() As Byte
        
            str = Protocol.PrepareMessageSendMsgBox( _
                    "El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        
            ReDim Preserve data(Len(str) - 1) As Byte
        
            data = StrConv(str, vbFromUnicode)
        
     
            Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
            Call WSApiCloseSocket(NuevoSock)

        End If
    
 
End Sub

Public Sub EventoSockRead(ByVal slot As Integer, ByRef Datos() As Byte)
 
        With UserList(slot)
    

            
               'SEGURIDAD ////////////
    
    
           If .flags.UserLogged Then
            Security.NAC_D_Byte Datos, UserList(slot).Redundance
        Else
            Security.NAC_D_Byte Datos, 13 'DEFAULT
        End If
    
        
        'SEGURIDAD ////////////
        
            Call .incomingData.WriteBlock(Datos)
    
            If .ConnID <> -1 Then
                Do While HandleIncomingData(slot)
                Loop
            Else
                Exit Sub

            End If

        End With

 
End Sub

Public Sub EventoSockClose(ByVal slot As Integer)
 
        'Es el mismo user al que está revisando el centinela??
        'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
        If Centinela.RevisandoUserIndex = slot Then Call modCentinela.CentinelaUserLogout
    
  
        If UserList(slot).flags.UserLogged Then
            Call CloseSocketSL(slot)
            Call Cerrar_Usuario(slot)
        Else
            Call CloseSocket(slot)

        End If

 
End Sub

Public Sub WSApiReiniciarSockets()
         Dim i As Long

        'Cierra el socket de escucha
        If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
        'Cierra todas las conexiones
        For i = 1 To MaxUsers

            If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
                Call CloseSocket(i)
            End If
            
            'Call ResetUserSlot(i)
        Next i
    
        For i = 1 To MaxUsers
            Set UserList(i).incomingData = Nothing
            Set UserList(i).outgoingData = Nothing
        Next i
    
        ' No 'ta el PRESERVE :p
        ReDim UserList(1 To MaxUsers)

        For i = 1 To MaxUsers
            UserList(i).ConnID = -1
            UserList(i).ConnIDValida = False
        
            Set UserList(i).incomingData = New clsByteQueue
            Set UserList(i).outgoingData = New clsByteQueue
        Next i
    
        LastUser = 1
        NumUsers = 0
    
        Call LimpiaWsApi
        Call Sleep(100)
        Call IniciaWsApi(frmMain.hwnd)
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

 
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
         Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
        Call ShutDown(Socket, SD_BOTH)
 
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, _
                                ByRef lpCallerData As WSABUF, _
                                ByRef lpSQOS As FLOWSPEC, _
                                ByVal Reserved As Long, _
                                ByRef lpCalleeId As WSABUF, _
                                ByRef lpCalleeData As WSABUF, _
                                ByRef Group As Long, _
                                ByVal dwCallbackData As Long) As Long
         Dim sa As sockaddr
    
        'Check if we were requested to force reject

        If dwCallbackData = 1 Then
            CondicionSocket = CF_REJECT
            Exit Function

        End If
    
        'Get the address

        CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen
    
        If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
            CondicionSocket = CF_REJECT
            Exit Function

        End If

        CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
 
End Function
