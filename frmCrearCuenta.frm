VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Creación de cuenta"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   8985
   ControlBox      =   0   'False
   FillColor       =   &H00877365&
   Icon            =   "frmCrearCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearCuenta.frx":000C
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox passTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1125
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Introduzca la contraseña que llevará su cuenta."
      Top             =   2640
      Width           =   3180
   End
   Begin VB.TextBox txtPin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtPin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   6960
      MaxLength       =   4
      TabIndex        =   6
      ToolTipText     =   "Asegúrese de ser el mismo pin mostrado. "
      Top             =   3600
      Width           =   1530
   End
   Begin VB.TextBox checkeopin 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "Vuelva a introducir el código de seguridad."
      Top             =   2625
      Width           =   3255
   End
   Begin VB.TextBox nameTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Introduzca el nombre que llevará su cuenta."
      Top             =   1200
      Width           =   3030
   End
   Begin VB.TextBox pass1Txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Vuelva a introducir la contraseña para confirmar."
      Top             =   3450
      Width           =   3180
   End
   Begin VB.TextBox usercodeTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   3
      ToolTipText     =   "Introduzca un código de seguridad."
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Image Buttons 
      Height          =   315
      Index           =   1
      Left            =   5760
      Tag             =   "1"
      Top             =   4080
      Width           =   1770
   End
   Begin VB.Image Buttons 
      Height          =   315
      Index           =   0
      Left            =   7680
      Tag             =   "1"
      Top             =   4200
      Width           =   1050
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
' soporte@Link-AO.com.ar
'   - Relase Number 1
'*****************************************************************

Option Explicit
Private Sub Buttons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Audio.PlayWave(SND_CLICK)
    Call Form_MouseMove(Button, Shift, X, Y)

    Select Case Index
  
        Case 0
            If frmMain.Socket1.State <> sckClosed Then frmMain.Socket1.Disconnect
            Unload Me
            
        Case 1
 
 
            'Creamos la acc
            Cuenta.UserAccount = nameTxt.Text
            Cuenta.UserPassword = passTxt.Text
            Cuenta.UserCode = usercodeTxt.Text
            
            If Not CheckData Then Exit Sub

            If IntervaloPermiteConectar Then
                Call FormParser.Parse_Form(Me, E_WAIT)

                If frmMain.Socket1.Connected Then
                    EstadoLogin = E_MODO.CrearNuevaCuenta
                    Call Login
                    Exit Sub
                    
                Else
                    EstadoLogin = E_MODO.CrearNuevaCuenta
                    frmMain.Socket1.HostName = CurServerIP
                    frmMain.Socket1.RemotePort = CurServerPort
                    frmMain.Socket1.Connect
                    
                End If
                
            End If
                

        End Select

End Sub

 
Public Sub InitCuenta()

On Error Resume Next

'Me.Picture = General_Load_Picture_From_Resource_Ex("_14")

Call FormParser.Parse_Form(frmConnect, e_normal)

Call FormParser.Parse_Form(Me, e_normal)


txtPin(1).Text = RandomNumber(1000, 9000)
Me.Show vbModeless, frmConnect

End Sub

Private Sub Form_Load()

'Me.Picture = General_Load_Picture_From_Resource_Ex("_14")
 
End Sub

Private Sub nameTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub pass1Txt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub passTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub usercodeTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub checkeopin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub txtPin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub usercodeTxt_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If
End Sub

Private Sub checkeopin_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If
    
End Sub

Private Sub txtPin_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
   Unload Me
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
     
    If Buttons(0).Tag = "1" Then
        'Buttons(0).Picture = Nothing
        Buttons(0).Tag = "0"
    End If
    
    If Buttons(1).Tag = "1" Then
        'Buttons(1).Picture = Nothing
        Buttons(1).Tag = "0"
    End If
  
End Sub
Private Sub Buttons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    
     Case 0
      ' Buttons(0).Picture = General_Load_Picture_From_Resource_Ex("_5")
       Buttons(0).Tag = "1"
      
     Case 1
      ' Buttons(1).Picture = General_Load_Picture_From_Resource_Ex("_6")
       Buttons(1).Tag = "1"
  End Select
End Sub
Private Sub Buttons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 
    If Index = 0 Then
        If Buttons(Index).Tag = "0" Then
            Call Form_MouseMove(Button, Shift, X, Y)
           ' Buttons(Index).Picture = General_Load_Picture_From_Resource_Ex("_67")
            Buttons(Index).Tag = "1"
        End If
    ElseIf Index = 1 Then
        If Buttons(Index).Tag = "0" Then
            Call Form_MouseMove(Button, Shift, X, Y)
           ' Buttons(Index).Picture = General_Load_Picture_From_Resource_Ex("_7")
            Buttons(Index).Tag = "1"
        End If
    End If
 
End Sub
Private Function CheckData() As Boolean

    Dim loopc As Integer
    Dim CharAscii As Integer
    
    
    If Len(Cuenta.UserAccount) = 0 Then
        MensajeAdvertencia Locale_GUI_Frase(258) 'Ingrese un nombre valido
        CheckData = False
        Exit Function
    End If
    
    If Len(Cuenta.UserAccount) > 20 Then
        MensajeAdvertencia Locale_GUI_Frase(259) 'Menos de 20 letras
        CheckData = False
        Exit Function
    End If
 
    For loopc = 1 To Len(Cuenta.UserAccount)
        CharAscii = Asc(mid$(Cuenta.UserAccount, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia Locale_GUI_Frase(251) 'Caracteres invalidos
            CheckData = False
            Exit Function
        End If
    Next loopc
 
    If Len(Cuenta.UserPassword) = 0 Then
        MensajeAdvertencia Locale_GUI_Frase(256) 'Ingrese una contraseña valida
        CheckData = False
        Exit Function
    End If
  
    
    If Len(Cuenta.UserPassword) > 30 Then
        MensajeAdvertencia Locale_GUI_Frase(260) ' contra no mayor de 30
        CheckData = False
        Exit Function
    End If
    
    For loopc = 1 To Len(Cuenta.UserPassword)
        CharAscii = Asc(mid$(Cuenta.UserPassword, loopc, 1))
        If LegalCharacter(CharAscii) = False Then
            MensajeAdvertencia Locale_GUI_Frase(257) 'Caracteres invalidos
            CheckData = False
            Exit Function
        End If
    Next loopc
    
     If Len(pass1Txt.Text) = 0 Then
        MensajeAdvertencia Locale_Error(57) 'Ingrese la confirmación de su contraseña.
        CheckData = False
        Exit Function
    End If
 
  
    If Not Cuenta.UserPassword = pass1Txt.Text Then
        MensajeAdvertencia Locale_Error(55) ' "Las contraseñas no coinciden."
        CheckData = False
        Exit Function
    End If
    
    If Len(Cuenta.UserCode) = 0 Then
        MensajeAdvertencia Locale_Error(58) 'Ingresa un código para tu cuenta.
        CheckData = False
        Exit Function
    End If
 
    'No permitimos letras en el pin
    If Not IsNumeric(Cuenta.UserCode) Then
        MensajeAdvertencia Locale_Error(53)  '"Solo se permite usar numeros en el código de cuenta."
        CheckData = False
        Exit Function
    End If
 
    If Len(checkeopin.Text) = 0 Then
        MensajeAdvertencia Locale_Error(59) 'Ingresa la confirmación del código de tu cuenta.
        CheckData = False
        Exit Function
    End If
 
    'No permitimos letras en el pin
    If Not IsNumeric(checkeopin.Text) Then
        MensajeAdvertencia Locale_Error(53)  '"Solo se permite usar numeros en el código de cuenta."
        CheckData = False
        Exit Function
    End If

    If Len(Cuenta.UserCode) > 4 Then
        MensajeAdvertencia Locale_Error(68)   ' code no mayor a 4
        CheckData = False
        Exit Function
    End If
    
    'Checkeo codigo de seguridad
    If Not Cuenta.UserCode = checkeopin.Text Then
        MensajeAdvertencia Locale_Error(61)  ' Los códigos no coinciden.
        CheckData = False
        Exit Function
    End If
 
    'Pin en blanco
    If Len(txtPin(0).Text) = 0 Then
        MensajeAdvertencia Locale_Error(60) ' Ingresa el pin que se muestra en pantalla para continuar.
        CheckData = False
        Exit Function
    End If
    
    'Checkeo de pin
    If Not txtPin(0).Text = txtPin(1).Text Then
        MensajeAdvertencia Locale_Error(56)  'El pin de verificación no coincide.
        CheckData = False
        Exit Function
    End If

    CheckData = True
    
End Function
