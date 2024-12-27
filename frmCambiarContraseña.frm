VERSION 5.00
Begin VB.Form frmCambiarContraseña 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$640"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChPass 
      Caption         =   "$589"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtNewPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "$640"
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
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$642"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$641"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmCambiarContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChPass_Click()

    If ChPass.Value = 1 Then
        txtPassword.PasswordChar = vbNullString
        txtNewPassword.PasswordChar = vbNullString
    Else
        txtPassword.PasswordChar = "*"
        txtNewPassword.PasswordChar = "*"
    End If

End Sub

Private Sub cmdAceptar_click()

    If IntervaloPermiteConectar Then
                    
        If Not CheckDat Then Exit Sub
        
        Cuenta.UserPassword = txtNewPassword.Text
        Cuenta.UserCode = txtPassword.Text
        Cuenta.EsChange = 1
        
        cmdAceptar.Caption = Locale_GUI_Frase(644) 'Conectando al servidor...
        cmdAceptar.Enabled = False
        
        Call FormParser.Parse_Form(Me, E_WAIT)
        
        If frmMain.Socket1.Connected Then
            EstadoLogin = E_MODO.CambiarContraseña
            Call Login
            Exit Sub
            
        Else
            EstadoLogin = E_MODO.CambiarContraseña
            frmMain.Socket1.HostName = CurServerIP
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
            
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(frmCharList)
Call FormParser.Parse_Form(Me)

Me.txtPassword.Text = vbNullString
Me.txtNewPassword.Text = vbNullString

End Sub

Private Function CheckDat() As Boolean

If txtPassword.Text = vbNullString Then
    MsgBox Locale_Error(49) ' Debes ingresar una contraseña válida
    CheckDat = False
    Exit Function
End If

If txtNewPassword.Text = vbNullString Then
    MsgBox Locale_Error(49) 'Debes ingresar una contraseña válida
    CheckDat = False
    Exit Function
End If

CheckDat = True

Exit Function

End Function


