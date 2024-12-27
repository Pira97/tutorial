VERSION 5.00
Begin VB.Form FrmInterv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Intervalos"
   ClientHeight    =   4710
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   74
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   64
      Top             =   2160
      Width           =   2655
      Begin VB.Frame Frame10 
         Caption         =   "Duracion Spells"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   135
         TabIndex        =   65
         Top             =   270
         Width           =   2400
         Begin VB.TextBox txtIntervaloVeneno 
            Height          =   300
            Left            =   195
            TabIndex        =   69
            Text            =   "0"
            Top             =   510
            Width           =   795
         End
         Begin VB.TextBox txtIntervaloParalizado 
            Height          =   300
            Left            =   195
            TabIndex        =   68
            Text            =   "0"
            Top             =   1170
            Width           =   795
         End
         Begin VB.TextBox txtIntervaloInvisible 
            Height          =   300
            Left            =   1170
            TabIndex        =   67
            Text            =   "0"
            Top             =   495
            Width           =   900
         End
         Begin VB.TextBox txtInvocacion 
            Height          =   300
            Left            =   1170
            TabIndex        =   66
            Text            =   "0"
            Top             =   1170
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Veneno"
            Height          =   180
            Left            =   225
            TabIndex        =   73
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Paralizado"
            Height          =   195
            Left            =   225
            TabIndex        =   72
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Invisible"
            Height          =   195
            Left            =   1170
            TabIndex        =   71
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Invocacion"
            Height          =   195
            Left            =   1170
            TabIndex        =   70
            Top             =   960
            Width           =   795
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame2 
         Caption         =   "Stamina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   3015
         TabIndex        =   59
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtStaminaIntervaloDescansar 
            Height          =   285
            Left            =   165
            TabIndex        =   61
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtStaminaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   60
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   62
            Top             =   930
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sanar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   4470
         TabIndex        =   54
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtSanaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   56
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtSanaIntervaloDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   55
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   58
            Top             =   930
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   57
            Top             =   255
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hambre y sed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5925
         TabIndex        =   49
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloSed 
            Height          =   285
            Left            =   150
            TabIndex        =   51
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloHambre 
            Height          =   285
            Left            =   150
            TabIndex        =   50
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sed"
            Height          =   195
            Left            =   165
            TabIndex        =   53
            Top             =   930
            Width           =   285
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hambre"
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   255
            Width           =   555
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Combate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   1545
         TabIndex        =   44
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloLanzaHechizo 
            Height          =   300
            Left            =   150
            TabIndex        =   46
            Text            =   "0"
            Top             =   525
            Width           =   930
         End
         Begin VB.TextBox txtPuedeAtacar 
            Height          =   300
            Left            =   135
            TabIndex        =   45
            Text            =   "0"
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Lanza Spell"
            Height          =   195
            Left            =   150
            TabIndex        =   48
            Top             =   285
            Width           =   825
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Puede Atacar"
            Height          =   195
            Left            =   135
            TabIndex        =   47
            Top             =   930
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   90
         TabIndex        =   39
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtTrabajo 
            Height          =   300
            Left            =   60
            TabIndex        =   41
            Text            =   "0"
            Top             =   1020
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloParaConexion 
            Height          =   300
            Left            =   45
            TabIndex        =   40
            Text            =   "0"
            Top             =   495
            Width           =   930
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Trabajo"
            Height          =   195
            Left            =   165
            TabIndex        =   43
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "IntervaloCon"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "NuevosINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   7440
         TabIndex        =   34
         Top             =   240
         Width           =   1410
         Begin VB.TextBox txtIntervaloIncinerado 
            Height          =   285
            Left            =   150
            TabIndex        =   36
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloAtacable 
            Height          =   285
            Left            =   150
            TabIndex        =   35
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Incinerado"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   255
            Width           =   750
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Atacable"
            Height          =   195
            Left            =   165
            TabIndex        =   37
            Top             =   930
            Width           =   630
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "NuevosINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   8880
         TabIndex        =   29
         Top             =   240
         Width           =   1290
         Begin VB.TextBox txtIntervaloOwnedNpc 
            Height          =   285
            Left            =   150
            TabIndex        =   31
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloPuedeSerAtacado 
            Height          =   285
            Left            =   150
            TabIndex        =   30
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "OwnedNpc"
            Height          =   195
            Left            =   165
            TabIndex        =   33
            Top             =   930
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Puede ser atacado"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   255
            Width           =   1350
         End
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Clima && Ambiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4680
      TabIndex        =   18
      Top             =   2160
      Width           =   2865
      Begin VB.Frame Frame7 
         Caption         =   "Frio y Fx Ambientales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2625
         Begin VB.TextBox txtIntervaloFrio 
            Height          =   285
            Left            =   180
            TabIndex        =   23
            Text            =   "0"
            Top             =   1080
            Width           =   915
         End
         Begin VB.TextBox txtIntervaloWAVFX 
            Height          =   300
            Left            =   150
            TabIndex        =   22
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloUserPuedeUsar 
            Height          =   300
            Left            =   1320
            TabIndex        =   21
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloFlechasCazadores 
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Text            =   "0"
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Frio"
            Height          =   195
            Left            =   195
            TabIndex        =   27
            Top             =   810
            Width           =   255
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "FxS"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   270
            Width           =   270
         End
         Begin VB.Label Vacio14 
            AutoSize        =   -1  'True
            Caption         =   "Puede usar"
            Height          =   195
            Left            =   1350
            TabIndex        =   25
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Vacio15 
            AutoSize        =   -1  'True
            Caption         =   "Flechas caza"
            Height          =   195
            Left            =   1320
            TabIndex        =   24
            Top             =   840
            Width           =   945
         End
      End
   End
   Begin VB.Frame Vacio6 
      Caption         =   "Nuevos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2880
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
      Begin VB.Frame Vacio5 
         Caption         =   "Nuevos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   1365
         Begin VB.TextBox txtIntervaloWS 
            Height          =   285
            Left            =   135
            TabIndex        =   15
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloGP 
            Height          =   285
            Left            =   150
            TabIndex        =   14
            Text            =   "0"
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Vacio4 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo WS"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   255
            Width           =   930
         End
         Begin VB.Label Vacio3 
            AutoSize        =   -1  'True
            Caption         =   "Guardar PJs"
            Height          =   195
            Left            =   165
            TabIndex        =   16
            Top             =   840
            Width           =   870
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Intervalos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Frame Frame14 
      Caption         =   "Intervalos nuevos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7560
      TabIndex        =   0
      Top             =   2160
      Width           =   2865
      Begin VB.Frame Frame15 
         Caption         =   "Intervalos nuevos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2625
         Begin VB.TextBox txtIntervaloOculto 
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Text            =   "0"
            Top             =   1110
            Width           =   915
         End
         Begin VB.TextBox txtIntervaloGolpeUsar 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloMagiaGolpe 
            Height          =   300
            Left            =   150
            TabIndex        =   3
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloGolpeMagia 
            Height          =   285
            Left            =   180
            TabIndex        =   2
            Text            =   "0"
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo oculto"
            Height          =   195
            Left            =   1320
            TabIndex        =   9
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Golpe Usar"
            Height          =   195
            Left            =   1350
            TabIndex        =   8
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Magia Golpe"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Golpe Magia"
            Height          =   195
            Left            =   195
            TabIndex        =   6
            Top             =   810
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "FrmInterv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub AplicarIntervalos()
    
    On Error GoTo AplicarIntervalos_Err
    
    '?????????? Intervalos del main loop ????????
    
100    IntervaloParaConexion = val(txtIntervaloParaConexion.Text)
102    IntervaloUserPuedeTrabajar = val(txtTrabajo.Text)
104    IntervaloUserPuedeCastear = val(txtIntervaloLanzaHechizo.Text)
106    IntervaloUserPuedeAtacar = val(txtPuedeAtacar.Text)
108    StaminaIntervaloDescansar = val(txtStaminaIntervaloDescansar.Text)
110    StaminaIntervaloSinDescansar = val(txtStaminaIntervaloSinDescansar.Text)
112    SanaIntervaloDescansar = val(txtSanaIntervaloDescansar.Text)
114    SanaIntervaloSinDescansar = val(txtSanaIntervaloSinDescansar.Text)
116    IntervaloHambre = val(txtIntervaloHambre.Text)
118    IntervaloSed = val(txtIntervaloSed.Text)
120    IntervaloIncinerado = val(txtIntervaloIncinerado.Text)
122    IntervaloVeneno = val(txtIntervaloVeneno.Text)
124    IntervaloParalizado = val(txtIntervaloParalizado.Text)
126    IntervaloInvisible = val(txtIntervaloInvisible.Text)
128    IntervaloInvocacion = val(txtInvocacion.Text)
130    IntervaloWavFx = val(txtIntervaloWAVFX.Text)
132    IntervaloFrio = val(txtIntervaloFrio.Text)
     
     'NEW
134    IntervaloMagiaGolpe = val(txtIntervaloMagiaGolpe.Text)
136    IntervaloGolpeMagia = val(txtIntervaloGolpeMagia.Text)
138    IntervaloGolpeUsar = val(txtIntervaloGolpeUsar.Text)
140    IntervaloOculto = val(txtIntervaloOculto.Text)
144    MinutosGuardarUsuarios = val(txtIntervaloGP.Text)
146    IntervaloPuedeSerAtacado = val(txtIntervaloPuedeSerAtacado.Text)
148    IntervaloAtacable = val(txtIntervaloAtacable.Text)
150    IntervaloOwnedNpc = val(txtIntervaloOwnedNpc.Text)
152    IntervaloUserPuedeUsar = val(txtIntervaloUserPuedeUsar.Text)
154    IntervaloFlechasCazadores = val(txtIntervaloFlechasCazadores.Text)
    
        Exit Sub

AplicarIntervalos_Err:
156     Call RegistrarError(Err.Number, Err.description, "FrmInterv.AplicarIntervalos", Erl)
158     Resume Next
End Sub


Private Sub Command1_Click()

    On Error GoTo Command1_Click_Err


100    Call AplicarIntervalos
    
       Exit Sub

Command1_Click_Err:
102  Call RegistrarError(Err.Number, Err.description, "FrmInterv.Command1_Click", Erl)
     Resume Next
        
End Sub
Private Sub Command2_Click()

    On Error GoTo Err

    'Intervalos
100    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion", str(IntervaloParaConexion))
102    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo", str(IntervaloUserPuedeTrabajar))
104    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", str(IntervaloUserPuedeCastear))
106    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", str(IntervaloUserPuedeAtacar))
108    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar", str(StaminaIntervaloDescansar))
110    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar", str(StaminaIntervaloSinDescansar))
112    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar", str(SanaIntervaloDescansar))
114    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar", str(SanaIntervaloSinDescansar))
116    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre", str(IntervaloHambre))
118    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed", str(IntervaloSed))
120    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloIncinerado", str(IntervaloIncinerado))
122    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno", str(IntervaloVeneno))
124    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado", str(IntervaloParalizado))
126    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible", str(IntervaloInvisible))
128    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion", str(IntervaloInvocacion))
130    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX", str(IntervaloWavFx))
132    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio", str(IntervaloFrio))
 
    'New
134    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe", str(IntervaloMagiaGolpe))
136    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia", str(IntervaloGolpeMagia))
138    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar", str(IntervaloGolpeUsar))
140    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto", str(IntervaloOculto))
144    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGuardarUsuarios", str(MinutosGuardarUsuarios))
    
146    Call WriteVar(IniPath & "Server.ini", "TIMERS", "IntervaloPuedeSerAtacado", str(IntervaloPuedeSerAtacado))
148    Call WriteVar(IniPath & "Server.ini", "TIMERS", "IntervaloAtacable", str(IntervaloAtacable))
150    Call WriteVar(IniPath & "Server.ini", "TIMERS", "IntervaloOwnedNpc", str(IntervaloOwnedNpc))
    
152    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar", str(IntervaloUserPuedeUsar))
154    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores", str(IntervaloFlechasCazadores))

156    MsgBox "Los intervalos se han guardado sin problemas."

       Exit Sub
Err:
158         MsgBox "Error al intentar grabar los intervalos"
160         Call RegistrarError(Err.Number, Err.description, "FrmInterv.AplicarIntervalos", Erl)
End Sub
Private Sub Form_Load()

    On Error GoTo Err
    
    With Me
    
100        .txtIntervaloParaConexion.Text = IntervaloParaConexion
102        .txtTrabajo.Text = IntervaloUserPuedeTrabajar
104        .txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
106        .txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
108        .txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
110        .txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
112        .txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
114        .txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
116        .txtIntervaloHambre.Text = IntervaloHambre
118        .txtIntervaloSed.Text = IntervaloSed
120        .txtIntervaloIncinerado.Text = IntervaloIncinerado
        
122        .txtIntervaloVeneno.Text = IntervaloVeneno
124        .txtIntervaloParalizado.Text = IntervaloParalizado
        
126        .txtIntervaloInvisible.Text = IntervaloInvisible
128        .txtInvocacion.Text = IntervaloInvocacion
        
        
130        .txtIntervaloWAVFX.Text = IntervaloWavFx
132        .txtIntervaloFrio.Text = IntervaloFrio
         
        'New
134        .txtIntervaloMagiaGolpe.Text = IntervaloMagiaGolpe
136        .txtIntervaloGolpeMagia.Text = IntervaloGolpeMagia
138        .txtIntervaloGolpeUsar.Text = IntervaloGolpeUsar
        
140        .txtIntervaloOculto.Text = IntervaloOculto
144        .txtIntervaloGP.Text = MinutosGuardarUsuarios
        
146        .txtIntervaloPuedeSerAtacado = IntervaloPuedeSerAtacado
148        .txtIntervaloAtacable = IntervaloAtacable
150        .txtIntervaloOwnedNpc = IntervaloOwnedNpc
        
152        .txtIntervaloUserPuedeUsar = IntervaloUserPuedeUsar
154        .txtIntervaloFlechasCazadores = IntervaloFlechasCazadores
    End With
    
    Exit Sub

Err:
156     Call RegistrarError(Err.Number, Err.description, "FrmInterv.ok_Click", Erl)
158     Resume Next
End Sub

Private Sub ok_Click()
On Error GoTo ok_Click_Err

100    Me.Visible = False
        Exit Sub

ok_Click_Err:
102     Call RegistrarError(Err.Number, Err.description, "FrmInterv.ok_Click", Erl)
104     Resume Next
        
End Sub

