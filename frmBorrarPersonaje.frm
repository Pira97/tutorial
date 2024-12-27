VERSION 5.00
Begin VB.Form frmBorrarPersonaje 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMail 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Cuenta"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Borrar personaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmBorrarPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msgvalue As Integer
Private Sub Command1_Click()
#If UsarWrench = 1 Then
                If frmMain.Socket1.Connected Then
                    frmMain.Socket1.Disconnect
                    frmMain.Socket1.Cleanup
                    DoEvents
                End If
            #Else
                If frmMain.Winsock1.state <> sckClosed Then
                    frmMain.Winsock1.Close
                    DoEvents
                End If
            #End If
        
     
msgvalue = MsgBox("Está seguro de borrar el personaje?", vbInformation + vbYesNo, "Mensaje de Alerta")
 
Select Case msgvalue
 
Case 6 'Yes
         
               Delete_UserName = nombrepj
            Delete_UserPassword = UserPassword
            Delete_cuenta = UserAccount
               EstadoLogin = E_MODO.BorrarPersonaje

            #If UsarWrench = 1 Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
            #Else
                frmMain.Winsock1.Connect CurServerIp, CurServerPort
            #End If
     ' AQUI VA EL Grupo de instrucciones que se va a ejecutar al presionar la opción SI
 
Case 7 'No
 Unload Me
 ' AQUI VA EL Grupo de instrucciones que se va a ejecutar al presionar la opción NO
 
End Select
 

          
End Sub

 
