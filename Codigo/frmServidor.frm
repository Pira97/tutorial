VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servidor"
   ClientHeight    =   5880
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6660
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recargar"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdRecargarAdministradores 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reload GMs"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdRecargarClanes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clanes"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1290
         Width           =   1575
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Server.ini"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.ListBox listDats 
         Height          =   1230
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Administracion"
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   6375
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardias en pos original"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1860
         Width           =   1935
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset Listen"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset sockets"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Debug UserList"
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unban All IPs (PELIGRO!)"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unban All (PELIGRO!)"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Debug Npcs"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Administracion"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pausar el servidor"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Stats de Slots"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Trafico"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Config. Intervalos"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Backup"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   6375
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cargar Mapas"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar Chars"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar Mapas"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir (Esc)"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdForzarCierre 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cierre forzado"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1935
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdForzarCierre_Click()
    If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
        Dim f
        For Each f In Forms
            Unload f
        Next
    End If
End Sub
Private Sub cmdRecargarAdministradores_Click()
Call loadAdministrativeUsers
End Sub
Private Sub cmdRecargarClanes_Click()
Call LoadGuildsDB
End Sub
Private Sub Command10_Click()
    frmTrafic.Show
End Sub

Private Sub Command11_Click()
    frmConID.Show
End Sub

Private Sub Command12_Click()
    frmDebugNpc.Show
End Sub
Private Sub Command15_Click()

    On Error Resume Next

    Dim Fn       As String
    Dim cad$
    Dim n        As Integer, K As Integer

    Dim sENtrada As String

    sENtrada = InputBox( _
            "Escribe ""estoy DE acuerdo"" entre comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes.", _
            "UnBan", "hola")

    If sENtrada = "estoy DE acuerdo" Then

        Fn = App.Path & "\logs\GenteBanned.log"
    
        If FileExist(Fn, vbNormal) Then
            n = FreeFile
            Open Fn For Input Shared As #n

            Do While Not EOF(n)
                K = K + 1
                Input #n, cad$
                Call UnBan(cad$)
            
            Loop
            Close #n
            MsgBox "Se han habilitado " & K & " personajes."
            Kill Fn

        End If

    End If

End Sub
Private Sub Command16_Click()
    Call LoadSini
End Sub
Private Sub Command18_Click()

    Me.MousePointer = 11

    Call GuardarUsuarios
    
    Me.MousePointer = 0
    
End Sub
Private Sub Command19_Click()
    Dim i        As Long, n As Long

    Dim sENtrada As String

    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distinción de mayúsculas minúsculas para desbanear a todos los personajes", "UnBan", "hola")

    If sENtrada = "estoy DE acuerdo" Then
    
        n = BanIps.count

        For i = 1 To BanIps.count
            BanIps.Remove 1
        Next i
    
        MsgBox "Se han habilitado " & n & " ipes"

    End If

End Sub
Private Sub Command2_Click()
    frmServidor.Visible = False
End Sub
Private Sub Command20_Click()
        If MsgBox("¿Está seguro que desea reiniciar los sockets? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
            Call WSApiReiniciarSockets
            Call DesconectarCuenta
        End If
End Sub

'Barrin 29/9/03
Private Sub Command21_Click()

    If EnPausa = False Then
        EnPausa = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(3)) 'Iniciando mantenimiento
        Command21.Caption = "Reanudar el servidor"
    Else
        EnPausa = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(436)) 'Finalizado mantenimiento
        Command21.Caption = "Pausar el servidor"
    End If
End Sub

Private Sub Command22_Click()
    Me.Visible = False
    frmAdmin.Show
End Sub
Private Sub Command26_Click()

        'Cierra el socket de escucha
        If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
        'Inicia el socket de escucha
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

End Sub
Private Sub Command27_Click()
    frmUserList.Show
End Sub
Private Sub Command5_Click()

    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.AgregarConsola "Reiniciando."

    FrmStat.Show

    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
    If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"

         Call apiclosesocket(SockListen)
 
    Dim loopc As Integer

    For loopc = 1 To MaxUsers
        Call CloseSocket(loopc)
    Next

    LastUser = 0
    NumUsers = 0

    Call FreeNPCs
    Call FreeCharIndexes

    Call LoadSini
    Call CargarBackUp
    Call LoadOBJData

         SockListen = ListenForConnect(Puerto, hWndMsg, "")

     If frmMain.Visible Then frmMain.AgregarConsola "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
    Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()
    FrmInterv.Show
End Sub
Private Sub Form_Deactivate()
    frmServidor.Visible = False
End Sub

Private Sub Form_Load()
        Command20.Visible = True
        Command26.Visible = True
        
    'Listamos el contenido de la carpeta Dats
    Dim sFilename As String
        sFilename = dir$(DatPath)
    
    Do While sFilename > vbNullString
    
      Call listDats.AddItem(sFilename)
      sFilename = dir$()
    
    Loop
    
End Sub

Private Sub listDats_Click()
    
    'Chequeamos si hay algun item seleccionado.
    'Lo pongo para prevenir errores.
    If listDats.ListIndex < 0 Then Exit Sub
    
    Select Case UCase$(listDats.Text)
        
        Case "APUESTAS.DAT"
            Call CargaApuestas
            
        Case "ARMASHERRERO.DAT"
            Call LoadArmasHerreria
        
        Case "ARMADURASHERRERO.DAT"
            Call LoadArmadurasHerreria
        
        Case "BALANCE.DAT"
            Call LoadBalance
        
        Case "BANIPS.DAT"
            Call BanIpCargar
            
        Case "CASCOSHERRERO.DAT"
             Call LoadCascosHerreria
            
            
        Case "CUIDADES.DAT"
            Call CargarCiudades
            
            
        Case "DROPGLOBAL.DAT"
            MsgBox "Drop Global"
            
        Case "ESCUDOSHERRERO.DAT"
             Call LoadEscudosHerreria
            
        Case "HECHIZOS.DAT"
            Call CargarHechizos
            
        Case "MOTD.INI"
            Call LoadMotd
            
        Case "ACTUALIZACIONES.INI"
            Call LoadUpdate
            
        Case "NIVELES.DAT"
            Call CargarELU
 
        Case "NPCS.DAT"
            Call CargaNpcsDat(True)

        Case "OBJ.DAT"
            Call LoadOBJData
        
        Case "OBJCARPINTERO.DAT"
             Call LoadObjCarpintero
        
        Case "OBJDRUIDA.DAT"
             Call LoadObjDruida
           
        Case "OBJSASTRE.DAT"
             Call LoadObjSastre
        
        Case "NOMBRESINVALIDOS.txt"
            Call CargarForbidenWords

        Case Else
            Exit Sub
            
    End Select
    
End Sub



