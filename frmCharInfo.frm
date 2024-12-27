VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "$440"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame charinfo 
      Caption         =   "$68"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5355
      Begin VB.Label Nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   5040
      End
      Begin VB.Label Nivel 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   3105
      End
      Begin VB.Label Clase 
         Caption         =   "Clase:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   3270
      End
      Begin VB.Label Raza 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Genero 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   4845
      End
      Begin VB.Label Banco 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   2985
      End
      Begin VB.Label status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   2760
      End
      Begin VB.Label guildactual 
         Caption         =   "Clan Actual:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   2880
      End
      Begin VB.Label ejercito 
         Caption         =   "Faccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   2880
      End
      Begin VB.Label Ciudadanos 
         Caption         =   "Ciudadanos asesinados:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   2850
      End
      Begin VB.Label criminales 
         Caption         =   "Criminales asesinados:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label reputacion 
         Caption         =   "Reputacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   2445
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$34"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   135
      TabIndex        =   20
      Top             =   4560
      Width           =   5355
      Begin VB.TextBox txtPeticiones 
         Enabled         =   0   'False
         Height          =   855
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   480
         Width           =   5070
      End
      Begin VB.TextBox txtMiembro 
         Enabled         =   0   'False
         Height          =   975
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1800
         Width           =   5070
      End
      Begin VB.Label lblSolicitado 
         Caption         =   "Ultimas membresías solicitadas:"
         Height          =   255
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   2985
      End
      Begin VB.Label lblMiembro 
         Caption         =   "Ultimos clanes en los que participó:"
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   1620
         Width           =   2985
      End
   End
   Begin VB.CommandButton imgCerrar 
      Cancel          =   -1  'True
      Caption         =   "$2"
      Height          =   495
      Left            =   135
      MouseIcon       =   "frmCharInfo.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   7905
      Width           =   1000
   End
   Begin VB.CommandButton imgRechazar 
      Caption         =   "$21"
      Height          =   495
      Left            =   3360
      MouseIcon       =   "frmCharInfo.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   7905
      Width           =   1000
   End
   Begin VB.CommandButton imgAceptar 
      Caption         =   "$1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MouseIcon       =   "frmCharInfo.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   7905
      Width           =   1000
   End
   Begin VB.CommandButton imgEchar 
      Caption         =   "$442"
      Height          =   495
      Left            =   1200
      MouseIcon       =   "frmCharInfo.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   7905
      Width           =   1000
   End
   Begin VB.CommandButton imgPeticion 
      Caption         =   "$441"
      Height          =   495
      Left            =   2280
      MouseIcon       =   "frmCharInfo.frx":0554
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   7905
      Width           =   1000
   End
   Begin VB.TextBox txtPeticiones1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4050
      Width           =   5730
   End
   Begin VB.TextBox txtMiembro1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5535
      Width           =   5730
   End
   Begin VB.Image imgAceptar1 
      Height          =   510
      Left            =   11640
      Tag             =   "1"
      Top             =   6795
      Width           =   1020
   End
   Begin VB.Image imgRechazar1 
      Height          =   510
      Left            =   10320
      Tag             =   "1"
      Top             =   6795
      Width           =   1020
   End
   Begin VB.Image imgPeticion1 
      Height          =   510
      Left            =   9120
      Tag             =   "1"
      Top             =   6795
      Width           =   1020
   End
   Begin VB.Image imgEchar1 
      Height          =   510
      Left            =   7920
      Tag             =   "1"
      Top             =   6795
      Width           =   1020
   End
   Begin VB.Image imgCerrar1 
      Height          =   510
      Left            =   6600
      Tag             =   "1"
      Top             =   6795
      Width           =   1020
   End
   Begin VB.Label status1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Label Nombre1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   1545
      Width           =   1440
   End
   Begin VB.Label Nivel1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   12
      Top             =   2595
      Width           =   1185
   End
   Begin VB.Label Clase1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   2070
      Width           =   1575
   End
   Begin VB.Label Raza1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Label Genero1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   9
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label Oro1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   2850
      Width           =   1365
   End
   Begin VB.Label Banco1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   3090
      Width           =   1425
   End
   Begin VB.Label guildactual1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   6
      Top             =   1800
      Width           =   2265
   End
   Begin VB.Label ejercito1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   5
      Top             =   2070
      Width           =   1785
   End
   Begin VB.Label Ciudadanos1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11385
      TabIndex        =   4
      Top             =   2340
      Width           =   1185
   End
   Begin VB.Label criminales1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11400
      TabIndex        =   3
      Top             =   2610
      Width           =   1185
   End
   Begin VB.Label reputacion1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11385
      TabIndex        =   2
      Top             =   2880
      Width           =   1185
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'La Plata - Pcia, Buenos Aires
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit


Public Enum CharInfoFrmType

    frmMembers
    frmMembershipRequests

End Enum

Public frmType As CharInfoFrmType

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
     
    
    'Me.Picture = LoadPicture(App.path & "\Graficos\VentanaInfoPj.jpg")
    Call FormParser.Parse_Form(Me)

     
End Sub
 
Private Sub imgAceptar_Click()
    Call WriteGuildAcceptNewMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me

End Sub

Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgEchar_Click()
    Call WriteGuildKickMember(Nombre)
    Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
    Unload Me

End Sub

Private Sub imgPeticion_Click()
    Call WriteGuildRequestJoinerInfo(Nombre)

End Sub

Private Sub imgRechazar_Click()
    frmCommet.t = RECHAZOPJ
    frmCommet.Nombre = Nombre.Caption
    frmCommet.Show vbModeless, frmCharInfo

End Sub

Private Sub txtMiembro_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
 
End Sub

