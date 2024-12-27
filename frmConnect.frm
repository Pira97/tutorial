VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "LinkAO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox servidores 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      IntegralHeight  =   0   'False
      ItemData        =   "frmConnect.frx":15F94E
      Left            =   8160
      List            =   "frmConnect.frx":15F950
      TabIndex        =   2
      Top             =   2760
      Width           =   2865
   End
   Begin MSWinsockLib.Winsock EstadoDelServidor 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtNombre 
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
      Height          =   315
      Left            =   5040
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3105
      Width           =   1935
   End
   Begin VB.TextBox txtPasswd 
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
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4455
      Width           =   1935
   End
   Begin VB.Image cmdButton 
      Height          =   360
      Index           =   5
      Left            =   11040
      Top             =   600
      Width           =   375
   End
   Begin VB.Image RecordarCuenta 
      Height          =   315
      Left            =   7080
      Top             =   5040
      Width           =   300
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   8775
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image WebLink 
      Height          =   345
      Left            =   9000
      Top             =   8640
      Width           =   2925
   End
   Begin VB.Image cmdButton 
      Height          =   300
      Index           =   4
      Left            =   11640
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image cmdButton 
      Height          =   300
      Index           =   3
      Left            =   11325
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image cmdButton 
      Height          =   555
      Index           =   2
      Left            =   5160
      Tag             =   "0"
      Top             =   6240
      Width           =   1755
   End
   Begin VB.Image cmdButton 
      Height          =   555
      Index           =   1
      Left            =   5040
      Tag             =   "0"
      Top             =   5640
      Width           =   2115
   End
   Begin VB.Image cmdButton 
      Height          =   615
      Index           =   0
      Left            =   5160
      Tag             =   "0"
      Top             =   4920
      Width           =   1755
   End
End
Attribute VB_Name = "frmConnect"
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
Private Sub Form_Activate()

On Error Resume Next

If RecordarCuentaIni = True Then
frmConnect.txtPasswd.SetFocus
Else
frmConnect.txtNombre.SetFocus
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call CerrarJuego
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub
Private Sub Form_Load()

    Debug.Print "Open Load frmConnect: " & Time
    
    'Me.Picture = General_Load_Picture_From_Resource_Ex("_58")
    
    Me.Caption = Form_Caption
    Dim J
    For Each J In cmdButton()
    J.Tag = "0"
    DoEvents
    Next
    
    Call cmdbutton_MouseUp(5, 0, 0, 0, 0)
    
    'CurrentUser.LogeoAlgunaVez = False
    If RecordarCuentaIni = True Then
        RecordarCuenta.Picture = General_Load_Picture_From_Resource_Ex("recordarsi")
        Me.txtNombre = UserAccountRecorded
        RecordarCuenta.Tag = 1
    Else
        RecordarCuenta.Picture = General_Load_Picture_From_Resource_Ex("recordarno")
        RecordarCuenta.Tag = 0
    End If
    
        
    Call FormParser.Parse_Form(frmConnect)
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

For i = 0 To cmdButton.UBound
    If cmdButton(i).Tag = "0" Then
        cmdButton(i).Picture = Nothing
        cmdButton(i).Tag = "1"
    End If
    DoEvents
Next i
 
End Sub
Private Sub cmdbutton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo errorhandler
    
    Call Audio.PlayWave(SND_CLICK)
    Call imgAccionRestaurar

        Select Case Index
    
            Case 0
                'frmCargando.Analizar

                If IntervaloPermiteConectar Then
                    Cuenta.UserAccount = frmConnect.txtNombre.Text
                    Cuenta.UserPassword = frmConnect.txtPasswd.Text
 
                    If RecordarCuentaIni = True Then
                        UserAccountRecorded = Cuenta.UserAccount
                    End If
            
                    If CheckUserData() = True Then
                        Call FormParser.Parse_Form(Me, E_WAIT)
                        Call Protocol.Connect(E_MODO.Normal)
                       End If
        
                End If
    
    
            Case 1
                If IntervaloPermiteConectar Then
                    Call FormParser.Parse_Form(Me, E_WAIT)
                    Call Protocol.Connect(E_MODO.CrearNuevaCuenta)
                End If
                
            Case 2
                If IntervaloPermiteConectar Then
                    Call FormParser.Parse_Form(Me, E_WAIT)
                    Call Protocol.Connect(E_MODO.RecuperarCuenta)
                End If
                
            Case 3
                Me.WindowState = 1
    
            Case 4
                Call CerrarJuego
            
            Case 5
            
                frmConnect.servidores.list(0) = Locale_GUI_Frase(637) & " " & "(Cargando estado)"
                
                If IntervaloPermiteConectar Then
                        
                    If frmConnect.EstadoDelServidor.State <> sckClosed Then
                        frmConnect.EstadoDelServidor.Close
                    End If
                
                    EstadoDelServidor.Connect CurServerIP, CurServerPort
            
                End If
    
        End Select
 
        Exit Sub
errorhandler:
    MsgBox "Se ha encontrado una aplicación prohibida. El programa ahora cerrará. Por favor, cierre las aplicaciones listadas a continuación y vuelva a iniciar el juego." & " (" & Err.Description & " - " & Err.number & ")", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar"
End Sub
Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 'Hacemos click en el botón
 
 Select Case Index
 
    Case 0
      'cmdButton(0).Picture = General_Load_Picture_From_Resource_Ex("_0")
      cmdButton(0).Tag = "0"
    Case 1
      'cmdButton(1).Picture = General_Load_Picture_From_Resource_Ex("_1")
      cmdButton(1).Tag = "0"
    Case 2
      'cmdButton(2).Picture = General_Load_Picture_From_Resource_Ex("_2")
      cmdButton(2).Tag = "0"
    Case 3
      'cmdButton(3).Picture = General_Load_Picture_From_Resource_Ex("_3")
      cmdButton(3).Tag = "0"
    Case 4
      'cmdButton(4).Picture = General_Load_Picture_From_Resource_Ex("_4")
      cmdButton(4).Tag = "0"
    Case 5
      'cmdButton(5).Picture = General_Load_Picture_From_Resource_Ex("_71")
      cmdButton(5).Tag = "0"
  End Select
  
End Sub

Private Sub cmdButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 'Pasamos el mouse el botón
 
 Select Case Index
 
    Case 0
    If cmdButton(0).Tag = "1" Then
     ' cmdButton(0).Picture = General_Load_Picture_From_Resource_Ex("_59")
      cmdButton(0).Tag = "0"
    End If
    
    Case 1
    If cmdButton(1).Tag = "1" Then
      'cmdButton(1).Picture = General_Load_Picture_From_Resource_Ex("_60")
      cmdButton(1).Tag = "0"
    End If
    
    Case 2
    If cmdButton(2).Tag = "1" Then
      'cmdButton(2).Picture = General_Load_Picture_From_Resource_Ex("_61")
      cmdButton(2).Tag = "0"
    End If
    
    Case 3
    If cmdButton(3).Tag = "1" Then
      'cmdButton(3).Picture = General_Load_Picture_From_Resource_Ex("_62")
      cmdButton(3).Tag = "0"
    End If
    
    Case 4
    If cmdButton(4).Tag = "1" Then
      'cmdButton(4).Picture = General_Load_Picture_From_Resource_Ex("_63")
      cmdButton(4).Tag = "0"
    End If
    
    Case 5
    If cmdButton(5).Tag = "1" Then
      'cmdButton(5).Picture = General_Load_Picture_From_Resource_Ex("_70")
      cmdButton(5).Tag = "0"
    End If
 End Select
 
 Call imgAccionRestaurar(Index)
 
End Sub
Private Sub imgAccionRestaurar(Optional ByVal NoIndex As Integer = 1000)

Dim i As Integer

For i = 0 To cmdButton.UBound
    If i <> NoIndex Then
        cmdButton(i).Picture = Nothing
        cmdButton(i).Tag = "1"
    End If
    DoEvents
Next i

End Sub

Private Sub Image1_Click()


End Sub

Private Sub EstadoDelServidor_Connect()
frmConnect.servidores.list(0) = Locale_GUI_Frase(637) & " " & Locale_GUI_Frase(639) & " - " & Time

If frmConnect.EstadoDelServidor.State <> sckClosed Then
    frmConnect.EstadoDelServidor.Close
End If
        
End Sub
 
Private Sub EstadoDelServidor_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
frmConnect.servidores.list(0) = Locale_GUI_Frase(637) & " " & Locale_GUI_Frase(638) & " - " & Time 'offline
End Sub
Private Sub RecordarCuenta_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
     If RecordarCuenta.Tag = 0 Then
     RecordarCuenta.Picture = General_Load_Picture_From_Resource_Ex("recordarsi")
     RecordarCuentaIni = True
     RecordarCuenta.Tag = 1
     MsgBox ("Tu cuenta fue guardada con exito")
     Else
     RecordarCuenta.Picture = General_Load_Picture_From_Resource_Ex("recordarno")
     RecordarCuentaIni = False
     RecordarCuenta.Tag = 0
     End If
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub txtPasswd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdButton_MouseDown(0, 0, 0, 0, 0)
        Call cmdbutton_MouseUp(0, 0, 0, 0, 0)
    End If
End Sub
Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdButton_MouseDown(0, 0, 0, 0, 0)
        Call cmdbutton_MouseUp(0, 0, 0, 0, 0)
    End If
End Sub
Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub WebLink_Click()
Call ShellExecute(0, "Open", "http://www.Link-AO.com.ar/", "", App.Path, SW_SHOWNORMAL)
End Sub
