VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CoverAO - Servidor  - Versión 1.0"
   ClientHeight    =   6180
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   7680
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6180
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command4 
      Caption         =   "IDs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   23
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mensajes todos los clientes (Solo testeo)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   50
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pop Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox BroadMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtChat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   4455
      End
      Begin VB.CheckBox chkServerHabilitado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo Game Masters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton CMDDUMP 
         Caption         =   "Crear Log critico de usuarios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   2760
         Width           =   2295
      End
   End
   Begin VB.Timer TIMER_AI 
      Interval        =   380
      Left            =   7080
      Top             =   0
   End
   Begin VB.Timer GameTimer 
      Interval        =   40
      Left            =   5640
      Top             =   0
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   0
   End
   Begin VB.Timer packetResend 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5160
      Top             =   0
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6600
      Top             =   0
   End
   Begin VB.CommandButton mnuSystray 
      Caption         =   "Systray"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton mnuCerrar 
      Caption         =   "Apagar SIN BackUp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Apagar CON BackUp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   4935
      Begin VB.CommandButton Command9 
         Caption         =   "Clean Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Save Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Height          =   1755
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label lblHora 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora del servidor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Logs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   390
      End
      Begin VB.Label CantUsuarios 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de usuarios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Record de usuarios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reload MD5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Timing Minuts:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5160
      TabIndex        =   0
      Top             =   3240
      Width           =   2415
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "MSG 2: Loading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "MSG 1: Loading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label lblCharSave 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Save Char: Loading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Evento Experiencia: Loading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2040
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Evento Oro: Loading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Evento Drop: Loading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   1560
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Public Lineas As Long

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA

    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64

End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hwnd As Long, _
                              lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA _
                Lib "SHELL32" (ByVal dwMessage As Long, _
                               lpData As NOTIFYICONDATA) As Integer
                               
                               


Private Function setNOTIFYICONDATA(hwnd As Long, _
                                   id As Long, _
                                   flags As Long, _
                                   CallbackMessage As Long, _
                                   Icon As Long, _
                                   Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp

End Function
 
Sub CheckIdleUser(ByVal iUserIndex As Integer)
 
    
    On Error GoTo hayerror

    With UserList(iUserIndex)
    
        'Conexion activa? y es un usuario loggeado?
        If .ConnID <> -1 And .flags.UserLogged Then
        
        .Counters.IdleCount = .Counters.IdleCount + 1
            'If Not EsGm(iUserIndex) Then
                If .Counters.IdleCount >= IdleLimit Then
                    
                    Call WriteShowMessageBox(iUserIndex, 33)
    
                    Call Cerrar_Usuario(iUserIndex)
                
                End If
                
            'End If
            
        End If

    End With
    
    Exit Sub
 
hayerror:
156     Call RegistrarError(Err.Number, Err.description, "frmMain.CheckIdleUser", Erl)
158     Resume Next
End Sub

Private Sub Auditoria_Timer()

Call mMainLoop.Auditoria

End Sub

Private Sub AutoSave_Timer()
 
    On Error GoTo ErrHandler

    'fired every minuto
    Static MinutosLatsClean As Long
    Static MinsPJesSave     As Long
    
    Static MensajeSpamUno   As Long
    Static MensajeSpamDos   As Long
    Static hora As String
    
    Dim i As Long
    
    MinsPJesSave = MinsPJesSave + 1
    MensajeSpamUno = MensajeSpamUno + 1
    MensajeSpamDos = MensajeSpamDos + 1
    MinutosLatsClean = MinutosLatsClean + 1

    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    Call ModAreas.AreasOptimizacion
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

    'Actualizamos el centinela
    'Call modCentinela.PasarMinutoCentinela

    'Actualizamos la lluvia
    'Call tLluviaEvent

    If MensajeSpamUno >= IntervaloMensajeAutomatico1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(437))
        MensajeSpamUno = 0
    End If

    If MensajeSpamDos >= IntervaloMensajeAutomatico2 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(438))
        MensajeSpamDos = 0
    End If
    
    'If MinutosLatsClean >= 15 Then
        'Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        'MinutosLatsClean = 0
    'End If
    
    If MinsPJesSave >= MinutosGuardarUsuarios Then
    
        For i = 1 To LastUser
            Call GuardarUsuarios
        Next i
        
        MinsPJesSave = 0
        
    End If
 
    'Mermas, mejor juntar todos los last User porque pensa que cada procedimiento revisa los user uno por uno, cuando podes hacer el loop una vez
    For i = 1 To LastUser
        
        Call PurgarPenas(i)
 
        Call CheckIdleUser(i)
        
    Next i
 
    Call CargarLabelsMain(MinsPJesSave, MensajeSpamUno, MensajeSpamDos)

    'Call seguridad_clones_limpiar
 

    Exit Sub
ErrHandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)

    Resume Next

End Sub

 

Private Sub CMDDUMP_Click()

    On Error Resume Next

    Dim i As Integer

    For i = 1 To MaxUsers
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & _
                " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
    Next i

    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text & vbCrLf & "<LinkAO Staff>"))

    txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
    
End Sub

Public Sub InitMain(ByVal f As Byte)

    If f = 1 Then
        Call mnuSystray_Click
    Else
        frmMain.Show

    End If

End Sub

Private Sub Command2_Click()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, _
            FontTypeNames.FONTTYPE_SERVER))
    ''''''''''''''''SOLO PARA EL TESTEO'''''''
    ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
    txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text

End Sub

Private Sub Command3_Click()
   Call MD5sCarga

End Sub

Private Sub Command4_Click()
frmRecibeDatos.Show
End Sub

Private Sub Command5_Click()
    frmServidor.Visible = True
End Sub


Private Sub Command8_Click()

    Dim n As Integer
    Dim hora As String
    Dim texto As String
 
    hora = Format(Time, "hh:mm")
        
    texto = "---------" & hora & "---------" & vbCrLf & frmMain.txtStatus.Text
    
    n = FreeFile()
    Open App.Path & "\Documentacion\ & Log (" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & ")" & ".log" For Output Shared As n
    Print #n, texto
    Close #n
    
End Sub

Private Sub Command9_Click()
frmMain.txtStatus.Text = ""
Lineas = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
   
    If Not Visible Then

        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hwnd, hProcess
                AppActivate hProcess

            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp

                If hHook Then UnhookWindowsHookEx hHook: hHook = 0

        End Select
        
    Else
    
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                frmMain.Visible = True
                Me.Show
        End Select
        
    End If
   
End Sub

Private Sub QuitarIconoSystray()

    On Error Resume Next

    'Borramos el icono del systray
    Dim i   As Integer
    Dim nid As NOTIFYICONDATA

    nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

    i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next


    Call QuitarIconoSystray

        Call LimpiaWsApi

    Dim loopc As Integer

    For loopc = 1 To MaxUsers

        If UserList(loopc).ConnID <> -1 Then Call CloseSocket(loopc)
    Next
    
    'Log
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " server cerrado."
    Close #n
    Call seguridad_clones_destruir
    End

    Set SonidosMapas = Nothing

End Sub

Private Sub GameTimer_Timer()
    Call mMainLoop.GameTimer
End Sub

Private Sub mnuCerrar_Click()
    If MsgBox( _
            "¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", _
            vbYesNo) = vbYes Then
        Dim f

        For Each f In Forms

            Unload f
        Next

    End If
End Sub


Private Sub mnusalir_Click()
    Call mnuCerrar_Click

End Sub

Public Sub mnuMostrar_Click()

    On Error Resume Next

    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0

End Sub

Private Sub KillLog()

    On Error Resume Next

    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"

    End If

End Sub


 
Private Sub mnuSystray_Click()
    Dim i   As Integer
    Dim S   As String
    Dim nid As NOTIFYICONDATA

    S = "LinkAO - Server - Versión 1.0"
    nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)
    
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False
    
    If frmRecibeDatos.Visible = True Then frmRecibeDatos.Show
End Sub

Private Sub PacketResend_Timer()
    Call mMainLoop.packetResend
End Sub

Private Sub TIMER_AI_Timer()
Call mMainLoop.TIMER_AI
 End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc

End Sub

Private Sub tLluviaEvent()

On Error GoTo ErrorHandler
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 2 Then
                
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(Queclima))
            End If
    ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(Queclima))
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
            Lloviendo = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(Queclima))
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 2 Then
                
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle(Queclima))
            End If
    End If
End If

Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Public Sub CargarLabelsMain(Optional ByVal MinsPJesSave As Long = 0, Optional ByVal MensajeSpamUno As Long = 0, Optional ByVal MensajeSpamDos As Long = 0)
     
    frmMain.lblCharSave.Caption = "Próximo CharSave: " & MinutosGuardarUsuarios - MinsPJesSave

    frmMain.Label6.Caption = "Mensaje automatico 1: " & IntervaloMensajeAutomatico1 - MensajeSpamUno
    
    frmMain.Label7.Caption = "Mensaje automatico 2: " & IntervaloMensajeAutomatico2 - MensajeSpamDos
    
  
    
End Sub
Public Function AgregarConsola(ByVal texto As String)

Lineas = Lineas + 1

texto = "(" & Lineas & ") " & texto & " (" & Time & ")"

If Len(frmMain.txtStatus.Text) = 0 Then
    frmMain.txtStatus.Text = texto
Else
    frmMain.txtStatus.Text = frmMain.txtStatus.Text & vbCrLf & texto
End If

End Function
Private Sub txtStatus_Change()

If Lineas >= 500 Then
    Call Command8_Click
    
    Lineas = 0
    frmMain.txtStatus.Text = ""
    
End If

End Sub
 
    
    


