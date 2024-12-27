VERSION 5.00
Begin VB.Form frmCustomKeys 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$396"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   5550
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustumKeys.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ReiniciarMacros 
      Caption         =   "$344"
      Height          =   315
      Left            =   8040
      TabIndex        =   77
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ComboBox SelectAutoUsar 
      Height          =   315
      ItemData        =   "frmCustumKeys.frx":000C
      Left            =   8040
      List            =   "frmCustumKeys.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   76
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox Accion2 
      Height          =   315
      ItemData        =   "frmCustumKeys.frx":0031
      Left            =   8040
      List            =   "frmCustumKeys.frx":0047
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   3270
      Width           =   2415
   End
   Begin VB.ComboBox Accion1 
      Height          =   315
      ItemData        =   "frmCustumKeys.frx":00F3
      Left            =   8040
      List            =   "frmCustumKeys.frx":0109
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   32
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   31
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   390
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   30
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   6990
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   29
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   6270
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   28
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   27
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   4710
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   25
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   6990
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   24
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   6270
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   23
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   6990
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   22
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   6270
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   21
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox txtMSens 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10080
      TabIndex        =   46
      Text            =   "10"
      Top             =   5160
      Width           =   375
   End
   Begin VB.HScrollBar scrSens 
      Height          =   345
      LargeChange     =   15
      Left            =   8040
      Max             =   20
      Min             =   1
      TabIndex        =   44
      Top             =   5160
      Value           =   10
      Width           =   2025
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   2730
      TabIndex        =   42
      Top             =   5460
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   15
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4710
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   5460
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$421"
      Height          =   315
      Left            =   8070
      TabIndex        =   37
      Top             =   6270
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   20
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2550
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   19
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   390
      Width           =   2415
   End
   Begin VB.CommandButton cmdAcciones 
      Caption         =   "$25"
      Height          =   315
      Left            =   8070
      TabIndex        =   30
      Top             =   6930
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "$420"
      Height          =   315
      Left            =   8070
      TabIndex        =   34
      Top             =   5940
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "$419"
      Height          =   315
      Left            =   8070
      TabIndex        =   32
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   14
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3990
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   12
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2550
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   390
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4710
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3990
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2550
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   390
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   26
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3990
      Width           =   2415
   End
   Begin VB.Label lblAutoUsar 
      Caption         =   "$576"
      Height          =   255
      Left            =   8040
      TabIndex        =   75
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$578"
      Height          =   195
      Index           =   34
      Left            =   8040
      TabIndex        =   74
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$577"
      Height          =   195
      Index           =   33
      Left            =   8040
      TabIndex        =   73
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$569"
      Height          =   195
      Index           =   32
      Left            =   5400
      TabIndex        =   70
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$570"
      Height          =   195
      Index           =   31
      Left            =   5400
      TabIndex        =   69
      Top             =   4410
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$571"
      Height          =   195
      Index           =   30
      Left            =   5400
      TabIndex        =   68
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$572"
      Height          =   195
      Index           =   29
      Left            =   5400
      TabIndex        =   67
      Top             =   5970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$573"
      Height          =   195
      Index           =   28
      Left            =   5400
      TabIndex        =   66
      Top             =   6720
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$574"
      Height          =   195
      Index           =   27
      Left            =   8040
      TabIndex        =   65
      Top             =   90
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$575"
      Height          =   195
      Index           =   26
      Left            =   8040
      TabIndex        =   64
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$568"
      Height          =   195
      Index           =   25
      Left            =   2760
      LinkItem        =   "0"
      TabIndex        =   63
      Top             =   6690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$567"
      Height          =   195
      Index           =   24
      Left            =   2760
      TabIndex        =   62
      Top             =   5970
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$566"
      Height          =   195
      Index           =   23
      Left            =   120
      TabIndex        =   61
      Top             =   6690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$565"
      Height          =   195
      Index           =   22
      Left            =   120
      TabIndex        =   60
      Top             =   6000
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$579"
      Height          =   195
      Index           =   21
      Left            =   5400
      TabIndex        =   48
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$418"
      Height          =   195
      Index           =   20
      Left            =   8040
      TabIndex        =   45
      Top             =   4890
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$404"
      Height          =   195
      Index           =   19
      Left            =   2730
      TabIndex        =   43
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$402"
      Height          =   195
      Index           =   18
      Left            =   2730
      TabIndex        =   41
      Top             =   4410
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$399"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   39
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$409"
      Height          =   195
      Index           =   16
      Left            =   5400
      TabIndex        =   36
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$401"
      Height          =   195
      Index           =   15
      Left            =   5400
      TabIndex        =   35
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$411"
      Height          =   195
      Index           =   14
      Left            =   5400
      TabIndex        =   33
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$417"
      Height          =   195
      Index           =   13
      Left            =   5400
      TabIndex        =   31
      Top             =   90
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$408"
      Height          =   195
      Index           =   12
      Left            =   2760
      TabIndex        =   25
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$407"
      Height          =   195
      Index           =   11
      Left            =   2760
      TabIndex        =   23
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$406"
      Height          =   195
      Index           =   10
      Left            =   2760
      TabIndex        =   21
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$405"
      Height          =   195
      Index           =   9
      Left            =   2760
      TabIndex        =   19
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$403"
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   17
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$400"
      Height          =   195
      Index           =   7
      Left            =   2760
      TabIndex        =   15
      Top             =   90
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$398"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   4410
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$397"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$410"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$416"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$415"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$414"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$413"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   360
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ReiniciarMacros_Click()
If MsgBox("¡Está seguro que desea reiniciar los macros?", vbYesNo) = vbYes Then
Dim i As Byte
Kill App.Path & "\Init\Macros\" & UCase$(Cuenta.UserName) & ".Mac"
'Actualizamos los macros
Call LoadMacros(Cuenta.UserName)

For i = 1 To 11
Next i

Call UpdateMacroLabels(0)
End If
End Sub

Private Sub SelectAutoUsar_Click()

 AutoUsarActivado = SelectAutoUsar.ListIndex
 
 If AutoUsarActivado = 1 Then
   MouseUno = "Desactivado"
 ElseIf AutoUsarActivado = 2 Then
   MouseUno = "Activado"
 End If
 
 
End Sub
 
Private Sub Accion1_click()
accionMouseUno = Accion1.ListIndex

If Accion2.ListIndex = 2 And Accion1.ListIndex = 2 Then
 Accion2.ListIndex = 0
End If
  
If Accion2.ListIndex = 1 And Accion1.ListIndex = 1 Then
 Accion2.ListIndex = 0
End If
  
If Accion2.ListIndex = 3 And Accion1.ListIndex = 3 Then
 Accion2.ListIndex = 0
End If

If (accionMouseUno = 0) Then
MouseUno = "Lanzar Hechizo/Seleccionar/Inspeccionar"

ElseIf (accionMouseUno = 1) Then
MouseUno = "Accionar/Tomar objeto"

ElseIf (accionMouseUno = 2) Then
MouseUno = "Atacar / Lanzar hechizos (click derecho)"
 
ElseIf (accionMouseUno = 3) Then
MouseUno = "Usar objeto seleccionado"

ElseIf (accionMouseUno = 4) Then
MouseUno = "Acción rápida 8"

ElseIf (accionMouseUno = 5) Then
MouseUno = "Acción rápida 9"
End If
 

End Sub

Private Sub Accion2_click()
accionMousedos = Accion2.ListIndex
  If Accion1.ListIndex = 2 And Accion2.ListIndex = 2 Then
 Accion1.ListIndex = 0
 End If
   If Accion1.ListIndex = 1 And Accion2.ListIndex = 1 Then
 Accion1.ListIndex = 0
 End If
   If Accion1.ListIndex = 3 And Accion2.ListIndex = 3 Then
 Accion1.ListIndex = 0
 End If
If (accionMousedos = 0) Then
MouseDos = "Lanzar Hechizo/Seleccionar/Inspeccionar"

ElseIf (accionMousedos = 1) Then
MouseDos = "Accionar/Tomar objeto"

ElseIf (accionMousedos = 2) Then
MouseDos = "Atacar / Lanzar hechizos (click izquierdo)"

ElseIf (accionMousedos = 3) Then
MouseDos = "Usar objeto seleccionado"

ElseIf (accionMousedos = 4) Then
MouseDos = "Acción rápida 8"

ElseIf (accionMousedos = 5) Then
MouseDos = "Acción rápida 9"
 End If
 
End Sub

Private Sub cmdAccion_Click()
Dim i As Long

For i = 22 To 32
    Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub cmdAcciones_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call frmElegirTeclas.Show(vbModal)
    Accion1.ListIndex = 0
        Accion2.ListIndex = 0
 
    txtMSens.Text = 10
End Sub

Private Sub Command2_Click()

   Call WriteVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "AutoUsarActivado", AutoUsarActivado)
   
   Call WriteVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "AccionMouseDos", accionMousedos)

   Call WriteVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "AccionMouseUno", accionMouseUno)
   
    Call WriteVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "MouseSpeed", MouseSpeed)
Dim i As Long

For i = 1 To CustomKeys.Count
    If LenB(Text1(i).Text) = 0 Then
        Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
        Exit Sub
    End If
Next i

Call CustomKeys.SaveCustomKeys
 
End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)


    AutoUsarActivado = GetVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "AutoUsarActivado")
    CargamosMouse = GetVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "AccionMouseUno")
    CargamosMouseDos = GetVar(App.Path & "\Init\CovAoInit.ini", "CONFIG", "AccionMouseDos")
    
    Accion1.ListIndex = CargamosMouse
    Accion2.ListIndex = CargamosMouseDos
    SelectAutoUsar.ListIndex = AutoUsarActivado
    
    scrSens = MouseSpeed
    txtMSens.Text = MouseSpeed
     
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub
 

Private Sub scrSens_Change()

 
    txtMSens.Text = scrSens.Value
    MouseSpeed = scrSens.Value
    Call General_Set_Mouse_Speed(MouseSpeed)
 
 
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.Count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub texT1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub

