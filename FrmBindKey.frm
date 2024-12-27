VERSION 5.00
Begin VB.Form FrmBindKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar Acción"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBindKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optAccion 
      Caption         =   "$8"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "$9"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2430
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "$10"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2700
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "$11"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2970
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   390
      TabIndex        =   4
      Top             =   2070
      Width           =   2655
   End
   Begin VB.CommandButton cmdAccept 
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
      Left            =   1800
      TabIndex        =   9
      Top             =   3270
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "$2"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "/"
      Height          =   255
      Left            =   270
      TabIndex        =   3
      Top             =   2070
      Width           =   105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "$7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblTecla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$205"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "FrmBindKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()

On Error Resume Next

Dim i As Integer
    For i = optAccion.LBound To optAccion.UBound
        If optAccion(i).Value = True Then
            MacroList(MacroIndex).mTipe = i + 1
            Exit For
        End If
    Next i

    Select Case MacroList(MacroIndex).mTipe
    
    Case 1
    
    If LenB(Text1.Text) = 0 Then
    MacroList(MacroIndex).mTipe = 0
    MensajeAdvertencia (Locale_GUI_Frase(266))
    Exit Sub
    End If
    
    MacroList(MacroIndex).mTipe = eMacros.aComando
    MacroList(MacroIndex).Grh = 17506
    MacroList(MacroIndex).Nombre = UCase$(Text1.Text)
            
    Case 2
    
    If frmMain.hlst.list(frmMain.hlst.ListIndex) = Locale_GUI_Frase(269) Or _
    frmMain.hlst.ListIndex = -1 Then
    MacroList(MacroIndex).mTipe = 0
    Exit Sub
    End If
    
    MacroList(MacroIndex).mTipe = eMacros.aLanzar
    MacroList(MacroIndex).Grh = 609
    MacroList(MacroIndex).Nombre = frmMain.hlst.list(frmMain.hlst.ListIndex)
    MacroList(MacroIndex).SpellSlot = frmMain.hlst.ListIndex + 1
    
   Case 3
   
    If Inventario.SelectedItem = 0 Then
    MacroList(MacroIndex).mTipe = 0
    Unload Me
    Exit Sub
    End If
    
    MacroList(MacroIndex).mTipe = eMacros.aUsar
    MacroList(MacroIndex).Grh = Inventario.GrhIndex(Inventario.SelectedItem)
    MacroList(MacroIndex).Nombre = Inventario.ItemName(Inventario.SelectedItem)
    MacroList(MacroIndex).OBJIndex = Inventario.OBJIndex(Inventario.SelectedItem)
    MacroList(MacroIndex).Slot = Inventario.SelectedItem
            
    Case 4
    
    If Inventario.SelectedItem = 0 Then
    MacroList(MacroIndex).mTipe = 0
    Unload Me
    Exit Sub
    End If
    
    MacroList(MacroIndex).mTipe = eMacros.aEquipar
    MacroList(MacroIndex).Grh = Inventario.GrhIndex(Inventario.SelectedItem)
    MacroList(MacroIndex).Nombre = Inventario.ItemName(Inventario.SelectedItem)
    MacroList(MacroIndex).OBJIndex = Inventario.OBJIndex(Inventario.SelectedItem)
    MacroList(MacroIndex).Slot = Inventario.SelectedItem
 
    End Select
    
    Unload Me
    
    
    'If CantidadEnMacros Then Call UpdateMacroLabels(1)
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()

Call FormParser.Parse_Form(Me)

If MacroList(MacroIndex).mTipe <> 0 Then
    Select Case MacroList(MacroIndex).mTipe
        Case 1 'Envia
            optAccion(1).Value = True
            Text1.Text = MacroList(MacroIndex).Nombre
            Text1.Enabled = True
    End Select
End If
End Sub

Private Sub optAccion_Click(Index As Integer)

If Index = 0 Then
    Text1.Enabled = True
Else
    Text1.Enabled = False
End If

End Sub
