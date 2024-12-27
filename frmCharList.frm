VERSION 5.00
Begin VB.Form frmCharList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "LinkAO "
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
   Icon            =   "frmCharList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCharList.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Index           =   1
      Left            =   3960
      ScaleHeight     =   1320
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   1200
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   9
      Left            =   9780
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   5880
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1320
      Index           =   8
      Left            =   1080
      ScaleHeight     =   1320
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   5880
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1230
      Index           =   7
      Left            =   9765
      ScaleHeight     =   1230
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   3720
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   6
      Left            =   6900
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   3840
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1425
      Index           =   5
      Left            =   3960
      ScaleHeight     =   1425
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3720
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1425
      Index           =   4
      Left            =   1080
      ScaleHeight     =   1425
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   3720
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   3
      Left            =   9795
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   1200
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   2
      Left            =   6900
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   1200
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Index           =   0
      Left            =   1080
      ScaleHeight     =   1320
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   960
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Index           =   1
      Left            =   7800
      TabIndex        =   22
      Top             =   7200
      Width           =   675
   End
   Begin VB.Image WebLink 
      Height          =   345
      Left            =   9720
      Top             =   8640
      Width           =   2205
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   6
      Left            =   11655
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   5
      Left            =   11325
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgAccion 
      Height          =   375
      Index           =   4
      Left            =   5040
      Tag             =   "0"
      Top             =   6840
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   3
      Left            =   12000
      Tag             =   "0"
      Top             =   7200
      Width           =   915
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   2
      Left            =   7800
      Tag             =   "0"
      Top             =   8040
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   495
      Index           =   1
      Left            =   2160
      Tag             =   "0"
      Top             =   8160
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   375
      Index           =   0
      Left            =   4920
      Tag             =   "0"
      Top             =   8280
      Width           =   1755
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   6840
      TabIndex        =   13
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   9720
      TabIndex        =   14
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   840
      TabIndex        =   15
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   3960
      TabIndex        =   16
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   6720
      TabIndex        =   17
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   9720
      TabIndex        =   18
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   960
      TabIndex        =   19
      Top             =   7440
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   9600
      TabIndex        =   20
      Top             =   7560
      Width           =   1365
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   9720
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   960
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   9600
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   6720
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   3840
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   960
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   9600
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   6720
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   3840
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00C0C000&
      Height          =   165
      Index           =   0
      Left            =   7800
      TabIndex        =   21
      Top             =   6840
      Width           =   405
   End
   Begin VB.Label lblAccData 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   480
      Width           =   3705
   End
End
Attribute VB_Name = "frmCharList"
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

Public intSelChar As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
frmMain.Socket1.Disconnect
Call FormParser.Parse_Form(frmConnect)
frmConnect.Visible = True
Me.Visible = False
End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub
Private Sub Form_Load()
    Me.Caption = Form_Caption
   ' Me.Picture = General_Load_Picture_From_Resource_Ex("_29")

    Call FormParser.Parse_Form(Me)
    
 End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

For i = 0 To imgAccion.UBound
    If imgAccion(i).Tag = "0" Then
        If (i <> 3 And i <> 4) Or (intSelChar > 0 And LenB(lblAccData(intSelChar)) > 0) Then
            imgAccion(i).Picture = Nothing
            imgAccion(i).Tag = "1"
        End If
    End If
Next i
End Sub
 
Private Sub imgAcc_Click(Index As Integer)

On Error Resume Next
Call DatosPersonaje(Index, 1)
 
End Sub
 
Private Sub lblAccData_Click(Index As Integer)

On Error Resume Next
Call DatosPersonaje(Index, 0)

End Sub
Private Sub picChar_Click(Index As Integer)
Call FormParser.Parse_Form(frmCrearPersonaje)
On Error Resume Next
Call DatosPersonaje(Index, 1)

End Sub
Private Sub imgAccion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Index = 3 Or Index = 4) And (intSelChar <= 0 Or lblAccData(intSelChar).Caption = "") Then Exit Sub
     
    Call imgAccionRestaurar
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
    
    Case 0 'Cambiar PW
        If IntervaloPermiteConectar Then
        
            Call FormParser.Parse_Form(Me, E_WAIT)
            
            If frmMain.Socket1.Connected Then
                frmCambiarContraseña.Show vbModal, frmCharList
                Exit Sub
            Else
                EstadoLogin = E_MODO.RecuperarCuenta
                frmMain.Socket1.HostName = CurServerIP
                frmMain.Socket1.RemotePort = CurServerPort
                frmMain.Socket1.Connect
                
            End If
            
        End If
    
    Case 1 'Desconectar
        frmMain.Socket1.Disconnect
        Call FormParser.Parse_Form(frmConnect)
        frmConnect.Visible = True
        Me.Visible = False
    Call frmCharList.LimpiarPersonajes
    Case 2 'Crear PJ
        If CheckDataCrearPJ(0, False) = True Then
            Call AbrirFormCrearPersonaje
            Call frmCharList.LimpiarPersonajes
        Else
            frmMensaje.msg.Caption = Locale_Error(10) 'No puedes crear mas personajes, has llegado a tu límite de diez personajes.
            frmMensaje.Show , frmCharList
            
        End If
    
    Case 3 'Borrar
        frmPregunta.SetAccion 4, lblAccData(intSelChar).Caption
        If frmCharList.Visible Then frmPregunta.Show , frmCharList

    Case 4 'Conectar
    
        Call LogearPersonaje

    Case 5 'Minimizar
        Me.WindowState = 1
    
    Case 6 'Cerrar
        CerrarJuego
    
    End Select
  
End Sub
Private Function CheckDataConnect() As Boolean

    On Error GoTo errorhandler
    
    CheckDataConnect = False
    
    If LenB(lblAccData(intSelChar).Caption) <= 0 Then
        frmMensaje.msg.Caption = Locale_Error(73) 'Debes seleccionar un personaje para continuar.
        frmMensaje.Show , frmCharList
        Exit Function
    End If
 
    If Len(Cuenta.UserName) <= 0 Then
        frmMensaje.msg.Caption = Locale_Error(64)  'Nombre invalido
        frmMensaje.Show , frmCharList
        Exit Function
    End If
    
    If Not AsciiValidos(Cuenta.UserName) Then
        frmMensaje.msg.Caption = Locale_Error(64)  'Nombre invalido
        frmMensaje.Show , frmCharList
        Exit Function
    End If
            
    CheckDataConnect = True
    
    Exit Function

errorhandler:
    CheckDataConnect = False
    frmMensaje.msg.Caption = Locale_Error(36)  'Msj error
    frmMensaje.Show , frmCharList
    
    Call RegistrarError(Err.number, Err.Description, "frmConnect.CheckDataConnect", Erl)
    Resume Next
End Function
Private Sub imgAccion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 'Apretamos el botón
 
 Select Case Index
    
     Case 0
        'imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("_22")
        imgAccion(0).Tag = "0"
     Case 1
        'imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("_23")
        imgAccion(1).Tag = "0"
     Case 2
       ' imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("_24")
        imgAccion(2).Tag = "0"
     Case 3
     If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
       ' imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("_25")
        imgAccion(3).Tag = "0"
     Case 4
     If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
        'imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("_26")
        imgAccion(4).Tag = "0"
     Case 5
       ' imgAccion(5).Picture = General_Load_Picture_From_Resource_Ex("_27")
        imgAccion(5).Tag = "0"
     Case 6
       ' imgAccion(6).Picture = General_Load_Picture_From_Resource_Ex("_28")
        imgAccion(6).Tag = "0"
 End Select
 
End Sub

Private Sub imgAccion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 'Pasamos el mouse por arriba del botón
 
 Select Case Index
    
     Case 0
        If imgAccion(0).Tag = "1" Then
       ' imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("_15")
        imgAccion(0).Tag = "0"
        End If
     
     Case 1
        If imgAccion(1).Tag = "1" Then
       ' imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("_16")
        imgAccion(1).Tag = "0"
        End If
      
     Case 2
        If imgAccion(2).Tag = "1" Then
       ' imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("_17")
        imgAccion(2).Tag = "0"
        End If
     
     Case 3
     
     If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
     
        If imgAccion(3).Tag = "1" Then
        'imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("_18")
        imgAccion(3).Tag = "0"
        End If
        
     Case 4
     If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
     
        If imgAccion(4).Tag = "1" Then
        'imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("_19")
        imgAccion(4).Tag = "0"
        End If
        
     Case 5
        If imgAccion(5).Tag = "1" Then
       ' imgAccion(5).Picture = General_Load_Picture_From_Resource_Ex("_20")
        imgAccion(5).Tag = "0"
        End If
     
     Case 6
        If imgAccion(6).Tag = "1" Then
        'imgAccion(6).Picture = General_Load_Picture_From_Resource_Ex("_21")
        imgAccion(6).Tag = "0"
        End If
        
 End Select
 
Call imgAccionRestaurar(Index)
End Sub

Public Sub imgAccionRestaurar(Optional ByVal NoIndex As Integer = 1000)

Dim i As Integer

For i = 0 To imgAccion.UBound
    If i <> NoIndex Then
         If (i <> 3 And i <> 4) Or (intSelChar > 0 And LenB(lblAccData(intSelChar)) > 0) Then
            imgAccion(i).Picture = Nothing
            imgAccion(i).Tag = "1"
        End If
    End If
Next i

End Sub
Private Sub picChar_DblClick(Index As Integer)
    
    intSelChar = Index + 1
    
    If CheckDataCrearPJ(intSelChar, True) Then
        LogearPersonaje
    Else
        Call AbrirFormCrearPersonaje
    End If
End Sub
 
Private Sub lblAccData_DblClick(Index As Integer)

    intSelChar = Index
    
    If CheckDataCrearPJ(intSelChar, True) Then
        LogearPersonaje
    Else
        Call AbrirFormCrearPersonaje
    End If

End Sub
Private Sub imgAcc_DblClick(Index As Integer)
    
    intSelChar = Index + 1
    
    If CheckDataCrearPJ(intSelChar, True) Then
        LogearPersonaje
    Else
        Call AbrirFormCrearPersonaje
    End If
   
End Sub
Private Sub LogearPersonaje()
    
    On Error GoTo LogearPersonaje_Err
    
    Cuenta.UserName = lblAccData(intSelChar).Caption
    
    If Not CheckDataConnect Then Exit Sub
        
    If IntervaloPermiteConectar Then
        
        Call FormParser.Parse_Form(Me, E_WAIT)
        
        Call Protocol.ReConnect(E_MODO.ConectarPersonaje)
        
    End If

    Exit Sub

LogearPersonaje_Err:
    Call RegistrarError(Err.number, Err.Description, "frmConnect.LogearPersonaje", Erl)
    Resume Next
    
End Sub


Private Sub WebLink_Click()
ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.Link-AO.com.ar/" & Chr$(34), vbNullString, vbNullString, 1
End Sub
Public Sub LimpiarPersonajes()

Dim i As Long

For i = 0 To 9
    frmCharList.lblAccData(i + 1).Caption = vbNullString
    frmCharList.picChar(i).Picture = Nothing
    frmCharList.imgAcc(i).Picture = Nothing
Next i

frmCharList.lblCharData(0) = vbNullString
frmCharList.lblCharData(1) = vbNullString

frmCharList.imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardes") 'Botones desactivados
frmCharList.imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccondes") 'Botones desactivados

frmCharList.intSelChar = 0
frmCharList.imgAccionRestaurar

If frmCharList.lblAccData(10).Caption <> "" Then 'Si tiene más de 10 personajes desactivamos boton de crear
    frmCharList.imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("accCrearPersonaje")
End If
 
If frmMensaje.Visible Then
    Unload frmMensaje
End If


End Sub

Public Function CheckDataCrearPJ(Optional Index As Integer = 0, Optional ByVal CrearPJ As Boolean = False) As Boolean

    CheckDataCrearPJ = False

    If CrearPJ = False Then 'Crea PJ desde Boton 'Crear PJ'
    
        Dim loopc As Long
        For loopc = 1 To 10
            If LenB(lblAccData(loopc).Caption) = 0 Then
                CheckDataCrearPJ = True
                Exit Function
            End If
        Next loopc
    Else 'Crea PJ desde Slots
    
        If LenB(lblAccData(Index)) <> 0 Then
            CheckDataCrearPJ = True
            Exit Function
        End If

    End If

    CheckDataCrearPJ = True
    
End Function
 
Private Function DatosPersonaje(Index As Integer, ByVal Opcion As Boolean)

Dim Indexx As Integer

Select Case Opcion
Case 0
Index = Index
Indexx = Index - 1
Case 1 'Opcion: 1 = imgAcc, PicAcc
Indexx = Index
Index = Index + 1
End Select

If Index <> intSelChar Then
  If intSelChar > 0 Then imgAcc(intSelChar - 1) = Nothing
    intSelChar = Index
    imgAcc(Indexx).Picture = General_Load_Picture_From_Resource_Ex("s" & intSelChar)
     
    If LenB(lblAccData(intSelChar)) > 0 Then
        lblCharData(0) = ListaClases(cPJ(intSelChar).Clase) & " Nivel " & cPJ(intSelChar).Nivel
        lblCharData(1) = Map_NameLoad(cPJ(intSelChar).Mapa) 'Ubicacion
        imgAccionRestaurar
    Else
        lblCharData(0) = vbNullString
        lblCharData(1) = vbNullString
        imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardes")
        imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccrearpersonaje")
    End If
  End If

If frmPregunta.Visible Then frmPregunta.Visible = False

End Function
Private Function AbrirFormCrearPersonaje()
        Me.Visible = False
        frmCrearPersonaje.Show
End Function

