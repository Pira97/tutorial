VERSION 5.00
Begin VB.Form frmDruida 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$439"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlquimia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "$2"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdFabricar 
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
      Height          =   375
      Left            =   2550
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "1"
      Top             =   2610
      Width           =   4215
   End
   Begin VB.ListBox lstPociones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   2370
      Width           =   4215
   End
End
Attribute VB_Name = "frmDruida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFabricar_Click()
     On Error Resume Next

    Call WriteCraftalquimia(ObjAlquimia(lstPociones.ListIndex + 1), val(txtCantidad.Text))

    Unload Me
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub
Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub txtCantidad_Change()
If val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

If val(txtCantidad.Text) > 1000 Then
    txtCantidad.Text = 1
End If
End Sub
