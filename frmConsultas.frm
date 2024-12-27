VERSION 5.00
Begin VB.Form frmConsultas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración"
   ClientHeight    =   6420
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Height          =   2235
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3240
      Width           =   6135
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   5595
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Responder consulta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje                    Tipo                   Estado                   Enviado hace"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   6135
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   6240
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu cmdAccion 
         Caption         =   "Sumonear"
         Index           =   0
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir a usuario"
         Index           =   1
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Salir"
         Index           =   2
      End
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
Me.Visible = False
List1.Clear
End Sub

Private Sub Form_Load()
List1.Clear
Call FormParser.Parse_Form(Me)
End Sub
Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
End Sub

Private Sub mnutraer_Click()

    'Pablo (ToxicWaste)
    Dim Aux As String

    Aux = mid$(ReadField(1, List1.list(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.list(List1.ListIndex), Asc("-"))))
    Call WriteSummonChar(Aux)

    'Pablo (ToxicWaste)
    Call WriteSummonChar(ReadField(1, List1.list(List1.ListIndex), Asc("-")))
End Sub


Private Sub mnuIR_Click()

    'Pablo (ToxicWaste)
    Dim Aux As String

    Aux = mid$(ReadField(1, List1.list(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.list(List1.ListIndex), Asc("-"))))
    Call WriteGoToChar(Aux)
    '/Pablo (ToxicWaste)
    Call WriteGoToChar(ReadField(1, List1.list(List1.ListIndex), Asc("-")))
    
End Sub

Private Sub mnuBorrar_Click()

    If List1.ListIndex < 0 Then Exit Sub

    'Pablo (ToxicWaste)
    Dim Aux As String

    Aux = mid$(ReadField(1, List1.list(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.list(List1.ListIndex), Asc("-"))))
    Call WriteSOSRemove(Aux)
    '/Pablo (ToxicWaste)
Call WriteSOSRemove(List1.list(List1.ListIndex))
    
    List1.RemoveItem List1.ListIndex

End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu menU_usuario

    End If

End Sub

