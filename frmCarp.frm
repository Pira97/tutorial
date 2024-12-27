VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$444"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmCarp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Text            =   "1"
      Top             =   2280
      Width           =   4035
   End
   Begin VB.CommandButton Command3 
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
      Height          =   435
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2670
      Width           =   1710
   End
   Begin VB.ListBox lstArmas 
      Height          =   1815
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4080
   End
   Begin VB.CommandButton Command4 
      Caption         =   "$2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2670
      Width           =   1710
   End
   Begin VB.Label lblCantidad 
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
      Height          =   225
      Left            =   180
      TabIndex        =   3
      Top             =   2040
      Width           =   1395
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    On Error Resume Next
    If MainTimer.Check(TimersIndex.Attack) Then Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1), val(txtCantidad.Text))
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub txtCantidad_Change()
If val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

If val(txtCantidad.Text) > 1000 Then
    txtCantidad.Text = 1
End If

End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
End Sub
