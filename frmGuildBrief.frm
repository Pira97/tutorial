VERSION 5.00
Begin VB.Form frmGuildBrief 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   -30
   ClientWidth     =   7455
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
   Icon            =   "frmGuildBrief.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "&H8000000A&"
   Begin VB.Frame Frame1 
      Caption         =   "Info del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      Begin VB.Label eleccion 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   38
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Miembros 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   37
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblAlineacion 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label creacion 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Aliados 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Enemigos 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label antifaccion 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label web 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lider 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   520
         TabIndex        =   30
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label fundador 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   870
         TabIndex        =   29
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label nombre 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   750
         TabIndex        =   28
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label11 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label Label10 
         Caption         =   "Fundador:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha de creacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label Label8 
         Caption         =   "Lider:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label Label7 
         Caption         =   "Web site:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Label Label6 
         Caption         =   "Miembros:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label Label5 
         Caption         =   "Elecciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   6975
      End
      Begin VB.Label Label4 
         Caption         =   "Alineacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label Label3 
         Caption         =   "Clanes Enemigos:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   6975
      End
      Begin VB.Label Label2 
         Caption         =   "Clanes Aliados:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   6975
      End
      Begin VB.Label Label1 
         Caption         =   "Puntos Antifaccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   3105
      Width           =   7215
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Codex 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   5535
      Width           =   7215
      Begin VB.TextBox Desc 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.CommandButton imgCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildBrief.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6975
      Width           =   1455
   End
   Begin VB.CommandButton imgDeclararGuerra 
      Caption         =   "Declarar Guerra"
      Height          =   375
      Left            =   4440
      MouseIcon       =   "frmGuildBrief.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6975
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton imgOfrecerAlianza 
      Caption         =   "Ofrecer Alianza"
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmGuildBrief.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6975
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton imgOfrecerPaz 
      Caption         =   "Ofrecer Paz"
      Height          =   375
      Left            =   1800
      MouseIcon       =   "frmGuildBrief.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6975
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton imgSolicitarIngreso 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   6000
      MouseIcon       =   "frmGuildBrief.frx":0554
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6960
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildBrief"
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
 

Public EsLeader                As Boolean

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
 
    'Me.Picture = LoadPicture(App.path & "\graficos\VentanaDetallesClan.jpg")
    
     
End Sub
 
Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgDeclararGuerra_Click()
    Call WriteGuildDeclareWar(Nombre.Caption)
    Unload Me

End Sub

Private Sub imgOfrecerAlianza_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.t = TIPO.ALIANZA
    Call frmCommet.Show(vbModal, frmGuildBrief)

End Sub

Private Sub imgOfrecerPaz_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.t = TIPO.PAZ
    Call frmCommet.Show(vbModal, frmGuildBrief)

End Sub

Private Sub imgSolicitarIngreso_Click()
    Call frmGuildSol.RecieveSolicitud(Nombre.Caption)
    Call frmGuildSol.Show(vbModal, frmGuildBrief)

End Sub

