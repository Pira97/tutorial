VERSION 5.00
Begin VB.Form frmGMAyuda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$296"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
   Icon            =   "frmGM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSoporte 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   180
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "$28"
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
      Left            =   180
      TabIndex        =   6
      Top             =   5160
      Width           =   4215
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "$294"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   1360
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "$290"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "$291"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1360
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "$293"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   2
      Top             =   1650
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "$292"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   1
      Top             =   1650
      Width           =   1455
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "$289"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$26"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$27"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   180
      TabIndex        =   8
      Top             =   4320
      Width           =   4215
   End
End
Attribute VB_Name = "frmGMAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private opt As Long

Private Sub cmdSend_Click()

If TxtSoporte.Text = vbNullString Then
    Call MensajeAdvertencia(Locale_GUI_Frase(264))
    Exit Sub
ElseIf opt < 0 Then
        Call MensajeAdvertencia(Locale_GUI_Frase(265))
Exit Sub

Else
    Call WriteGMRequest(opt, TxtSoporte.Text)
    Unload Me
End If

End Sub
Private Sub Form_Load()
    opt = -1
    Call FormParser.Parse_Form(Me)
End Sub
Private Sub Label1_Click()
frmHlp.Show vbModeless, frmGMAyuda
End Sub

Private Sub optConsulta_Click(Index As Integer)

opt = Index

Select Case Index

    Case 0
        Label2.Caption = Locale_GUI_Frase(204)
        
    Case 1
        Label2.Caption = Locale_GUI_Frase(200)
    
    Case 2, 3
        Label2.Caption = Locale_GUI_Frase(201)

    Case 4
        Label2.Caption = Locale_GUI_Frase(653)
        
    Case 5
        Label2.Caption = Locale_GUI_Frase(199)
        
End Select

End Sub
