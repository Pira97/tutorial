VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas del personaje"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   -90
   ClientWidth     =   6450
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
   Icon            =   "frmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEstadisticas.frx":000C
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   6
      Left            =   930
      TabIndex        =   49
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   5
      Left            =   930
      TabIndex        =   48
      Top             =   3090
      Width           =   975
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5850
      TabIndex        =   47
      Top             =   3750
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   41
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":903C6
      Top             =   2100
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   40
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":90518
      Top             =   2010
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   39
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":9066A
      Top             =   1890
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   38
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":907BC
      Top             =   1770
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   37
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":9090E
      Top             =   1650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":90A60
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   35
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":90BB2
      Top             =   1440
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   34
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":90D04
      Top             =   1350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   33
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":90E56
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   32
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":90FA8
      Top             =   1110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   31
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":910FA
      Top             =   990
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   30
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":9124C
      Top             =   900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   29
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":9139E
      Top             =   750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   28
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":914F0
      Top             =   660
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   42
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91642
      Top             =   2250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   43
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91794
      Top             =   2370
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   44
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":918E6
      Top             =   2460
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   45
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91A38
      Top             =   2580
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   46
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91B8A
      Top             =   2700
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   47
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91CDC
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   48
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91E2E
      Top             =   2910
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   49
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":91F80
      Top             =   3000
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   50
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":920D2
      Top             =   3150
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   51
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":92224
      Top             =   3240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   52
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":92376
      Top             =   3360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   53
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":924C8
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   26
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":9261A
      Top             =   3600
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   24
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":9276C
      Top             =   3360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   22
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":928BE
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   20
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":92A10
      Top             =   2910
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   18
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":92B62
      Top             =   2700
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   16
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":92CB4
      Top             =   2460
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   14
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":92E06
      Top             =   2250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   12
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":92F58
      Top             =   2010
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   10
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":930AA
      Top             =   1800
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   8
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":931FC
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   6
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":9334E
      Top             =   1320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   4
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":934A0
      Top             =   1110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   2
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":935F2
      Top             =   870
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   0
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":93744
      Top             =   660
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   1
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":93896
      Top             =   750
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   27
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":939E8
      Top             =   3690
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   25
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":93B3A
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   23
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":93C8C
      Top             =   3210
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   21
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":93DDE
      Top             =   3000
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   19
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":93F30
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   17
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":94082
      Top             =   2550
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   15
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":941D4
      Top             =   2340
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   13
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":94326
      Top             =   2130
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   11
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":94478
      Top             =   1890
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   9
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":945CA
      Top             =   1680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   7
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":9471C
      Top             =   1440
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   5
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":9486E
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":949C0
      Top             =   960
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   1
      Left            =   1230
      TabIndex        =   46
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   2
      Left            =   1200
      TabIndex        =   45
      Top             =   4680
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   3
      Left            =   1230
      TabIndex        =   44
      Top             =   4890
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   4
      Left            =   1230
      TabIndex        =   43
      Top             =   5250
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   5
      Left            =   1230
      TabIndex        =   42
      Top             =   5490
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   6
      Left            =   1230
      TabIndex        =   41
      Top             =   5700
      Width           =   600
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4050
      TabIndex        =   40
      Top             =   900
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4050
      TabIndex        =   39
      Top             =   690
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4050
      TabIndex        =   38
      Top             =   1110
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4050
      TabIndex        =   37
      Top             =   1350
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4050
      TabIndex        =   36
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4050
      TabIndex        =   35
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4050
      TabIndex        =   34
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4050
      TabIndex        =   33
      Top             =   2250
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4050
      TabIndex        =   32
      Top             =   2490
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4050
      TabIndex        =   31
      Top             =   2700
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4050
      TabIndex        =   30
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4050
      TabIndex        =   29
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4050
      TabIndex        =   28
      Top             =   3390
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4050
      TabIndex        =   27
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   5700
      TabIndex        =   26
      Top             =   690
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   5700
      TabIndex        =   25
      Top             =   930
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   5700
      TabIndex        =   24
      Top             =   1140
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   5700
      TabIndex        =   23
      Top             =   1350
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   5700
      TabIndex        =   22
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   5700
      TabIndex        =   21
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5700
      TabIndex        =   20
      Top             =   1590
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   5700
      TabIndex        =   19
      Top             =   2250
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   5700
      TabIndex        =   18
      Top             =   2490
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   5700
      TabIndex        =   17
      Top             =   2700
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   5700
      TabIndex        =   16
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   5700
      TabIndex        =   15
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   5700
      TabIndex        =   14
      Top             =   3390
      Width           =   255
   End
   Begin VB.Image imgEstado 
      Height          =   315
      Left            =   525
      Top             =   6315
      Width           =   1110
   End
   Begin VB.Image imgFami 
      Height          =   1680
      Left            =   6555
      Top             =   3780
      Width           =   2265
   End
   Begin VB.Image cmdGuardar 
      Height          =   480
      Left            =   3780
      Tag             =   "1"
      Top             =   3900
      Width           =   1050
   End
   Begin VB.Image iEx 
      Height          =   450
      Left            =   6120
      Tag             =   "1"
      Top             =   0
      Width           =   390
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Veces muerto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   13
      Top             =   5550
      Width           =   1665
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acá van las habilidades especiales del familiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   6630
      TabIndex        =   12
      Top             =   5070
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas matadas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2250
      TabIndex        =   11
      Top             =   6180
      Width           =   1665
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   0
      Left            =   930
      TabIndex        =   10
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   4
      Left            =   7230
      TabIndex        =   9
      Top             =   4590
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6570
      TabIndex        =   8
      Top             =   3750
      Width           =   2220
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   8310
      TabIndex        =   7
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   5
      Left            =   1590
      TabIndex        =   4
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   4
      Left            =   1590
      TabIndex        =   3
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   3
      Left            =   1590
      TabIndex        =   2
      Top             =   1260
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   2
      Left            =   1590
      TabIndex        =   1
      Top             =   975
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   1
      Left            =   1590
      TabIndex        =   0
      Top             =   720
      Width           =   105
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   6735
      TabIndex        =   6
      Top             =   4230
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Shape fExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   6735
      Top             =   4260
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   3
      Left            =   8010
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Shape fHPShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   8010
      Top             =   4710
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmEstadisticas"
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
Private LibresOrig As Integer
Private RealizoCambios As Boolean
Private NewSkills(1 To NUMSKILLS) As Byte

Private Sub iEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave(SND_CLICK)
'iEx.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrar-est-down")
End Sub

Private Sub iEx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If iEx.Tag = "0" Then
        iEx.Tag = "1"
        iEx.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrar-est-over")
    End If
 End Sub
Private Sub iEx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub
Private Sub cmdGuardar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave(SND_CLICK)
cmdGuardar.Picture = General_Load_Skin_Picture_From_Resource_Ex("guardar-down")
End Sub

Private Sub cmdGuardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If cmdGuardar.Tag = "0" Then
    cmdGuardar.Tag = "1"
    cmdGuardar.Picture = General_Load_Skin_Picture_From_Resource_Ex("guardar-over")
End If
  
End Sub
 
Private Sub cmdGuardar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Long, Cambio As Integer
Dim Cad As String

Call Form_MouseMove(Button, Shift, X, Y)
 

         For i = 1 To NUMSKILLS
            NewSkills(i) = CByte(Skill(i - 1).Caption) - CurrentUser.UserSkills(i)
            CurrentUser.UserSkills(i) = val(Skill(i - 1).Caption)
        Next i
        Call WriteModifySkills(NewSkills())
        
Unload Me

End Sub
Private Sub Command1_Click(Index As Integer)
Dim Indice As Byte

If (Index And &H1) = 0 Then
    Indice = Index \ 2
    If SkillPoints > 0 And val(Skill(Indice).Caption) < 100 Then
        Skill(Indice).Caption = val(Skill(Indice).Caption) + 1
        SkillPoints = SkillPoints - 1
         Flags(Indice) = Flags(Indice) + 1
    End If
Else
    Indice = Index \ 2
   If Alocados < SkillPoints Then
    If val(Skill(Indice).Caption) > 0 And Flags(Indice) > 0 Then
         Skill(Indice).Caption = val(Skill(Indice).Caption) - 1
     Flags(Indice) = Flags(Indice) - 1
        SkillPoints = SkillPoints + 1
        Alocados = Alocados + 1
    End If
End If
End If

Puntos.Caption = SkillPoints
RealizoCambios = (SkillPoints <> LibresOrig)
'Skill(indice).ForeColor = IIf(CurrentUser.UserSkills(SkillRealToIndex(indice + 1)) = SkillsOrig(SkillRealToIndex(indice + 1)), vbWhite, vbRed)


End Sub
 Private Sub Form_Load()
' Me.Picture = General_Load_Skin_Picture_From_Resource_Ex("stats")
 imgFami.Picture = General_Load_Skin_Picture_From_Resource_Ex("fmnodisp")
Call FormParser.Parse_Form(Me)
 
 
ReDim Flags(1 To NUMSKILLS)
End Sub
    
Public Sub Iniciar_Labels()

On Error Resume Next

'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = CurrentUser.UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    Skill(i - 1).Caption = CurrentUser.UserSkills(i)
Next

With UserEstadisticas
    Label4(1).Caption = .CiudadanosMatados
    Label4(3).Caption = .CriminalesMatados
    label6(3).Caption = .NpcsMatados
    Label4(2).Caption = .RepublicanosMatados
    Label4(4).Caption = .ArmadasRealesMatados
    Label4(5).Caption = .MiliciasMatados
    Label4(6).Caption = .CaosMatados
 
    est(0).Caption = .Clase
    est(5).Caption = IIf(.Genero = 1, "Masculino", "Femenino")
    est(6).Caption = ListaRazas(.Raza)
    est(4).Caption = .MuertesUsuario
    Select Case .status
    

    Case 1: imgEstado.Picture = General_Load_Skin_Picture_From_Resource_Ex("renegado")
    Case 2: imgEstado.Picture = General_Load_Skin_Picture_From_Resource_Ex("imperial")
    Case 3: imgEstado.Picture = General_Load_Skin_Picture_From_Resource_Ex("republicano")
    Case 4: imgEstado.Picture = General_Load_Skin_Picture_From_Resource_Ex("caos")
    Case 5: imgEstado.Picture = General_Load_Skin_Picture_From_Resource_Ex("armada")
    Case 6: imgEstado.Picture = General_Load_Skin_Picture_From_Resource_Ex("miliciano")
    End Select
 
End With

 
 
LibresOrig = SkillPoints

Puntos.Caption = SkillPoints
RealizoCambios = False

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If iEx.Tag = "1" Then
    iEx.Tag = "0"
    iEx.Picture = Nothing
End If

If cmdGuardar.Tag = "1" Then
    cmdGuardar.Tag = "0"
    cmdGuardar.Picture = Nothing
End If

End Sub
 
    
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload Me
End Sub

Public Function SkillRealToIndex(ByVal SkillIndex As Integer) As Integer

Select Case SkillIndex
    Case 1
        SkillRealToIndex = 4
    Case 2
        SkillRealToIndex = 5
    Case 3
        SkillRealToIndex = 20
    Case 4
        SkillRealToIndex = 7
    Case 5
        SkillRealToIndex = 23
    Case 6
        SkillRealToIndex = 19
    Case 7
        SkillRealToIndex = 12
    Case 8
        SkillRealToIndex = 2
    Case 9
        SkillRealToIndex = 22
    Case 10
        SkillRealToIndex = 6
    Case 11
        SkillRealToIndex = 8
    Case 12
        SkillRealToIndex = 18
    Case 13
        SkillRealToIndex = 1
    Case 14
        SkillRealToIndex = 3
    Case 15
        SkillRealToIndex = 11
    Case 16
        SkillRealToIndex = 9
    Case 17
        SkillRealToIndex = 17
    Case 18
        SkillRealToIndex = 13
    Case 19
        SkillRealToIndex = 14
    Case 20
        SkillRealToIndex = 10
    Case 21
        SkillRealToIndex = 26
    Case 22
        SkillRealToIndex = 16
    Case 23
        SkillRealToIndex = 15
    Case 24
        SkillRealToIndex = 24
    Case 25
        SkillRealToIndex = 25
    Case 26
        SkillRealToIndex = 21
    Case 27
        SkillRealToIndex = 27
End Select

End Function
Public Function RealSkillToIndex(ByVal Skill As Integer) As Integer

Select Case Skill
    Case 4
        RealSkillToIndex = 1
    Case 5
        RealSkillToIndex = 2
    Case 20
        RealSkillToIndex = 3
    Case 7
        RealSkillToIndex = 4
    Case 23
        RealSkillToIndex = 5
    Case 19
        RealSkillToIndex = 6
    Case 12
        RealSkillToIndex = 7
    Case 2
        RealSkillToIndex = 8
    Case 22
        RealSkillToIndex = 9
    Case 6
        RealSkillToIndex = 10
    Case 8
        RealSkillToIndex = 11
    Case 18
        RealSkillToIndex = 12
    Case 1
        RealSkillToIndex = 13
    Case 3
        RealSkillToIndex = 14
    Case 11
        RealSkillToIndex = 15
    Case 9
        RealSkillToIndex = 16
    Case 17
        RealSkillToIndex = 17
    Case 13
        RealSkillToIndex = 18
    Case 14
        RealSkillToIndex = 19
    Case 10
        RealSkillToIndex = 20
    Case 26
        RealSkillToIndex = 21
    Case 16
        RealSkillToIndex = 22
    Case 15
        RealSkillToIndex = 23
    Case 24
        RealSkillToIndex = 24
    Case 25
        RealSkillToIndex = 25
    Case 21
        RealSkillToIndex = 26
    Case 27
        RealSkillToIndex = 27
End Select

End Function

 
 

