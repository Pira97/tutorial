VERSION 5.00
Begin VB.Form frmBuscar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "asd"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "frmBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   80
      Width           =   480
   End
   Begin VB.CommandButton BuscarUno 
      Caption         =   "$558"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton BuscarUno 
      Caption         =   "$561"
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
      Index           =   0
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton Limpiarlista 
      Caption         =   "$482"
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
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   4935
   End
   Begin VB.ListBox Resultados 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "frmBuscar.frx":000C
      Left            =   120
      List            =   "frmBuscar.frx":000E
      TabIndex        =   3
      Top             =   2160
      Width           =   4935
   End
   Begin VB.CommandButton Buscar 
      Caption         =   "$559"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Buscar 
      Caption         =   "$560"
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
      Index           =   0
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Busqueda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$562"
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
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$258"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$504"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuCrearO 
      Caption         =   "Crear Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear 1"
         Index           =   0
      End
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear 10"
         Index           =   1
      End
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear 100"
         Index           =   2
      End
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear N"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCrearN 
      Caption         =   "Crear NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearNPC 
         Caption         =   "Crear NPC"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessages Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Private Const LB_ITEMFROMPOINT = &H1A9
Private Const LB_FINDSTRING = &H18F
Private MensajeBusqueda As Boolean
Private BusquedaObjetos As Boolean

Private Sub Buscar_Click(Index As Integer)
    
    'Limpiamos la lista antes.
    Resultados.Clear
    
    Dim i As Integer
    
    Select Case Index
    
    Case 0
    
    BusquedaObjetos = True
    PicView.Visible = True
    For i = 1 To General_Locale_Obj(0, 4)
    Resultados.AddItem i & " - " & General_Locale_Obj(i, 0)
    Next i
    
    Buscar(1).Enabled = True
    
    Case 1
    
    BusquedaObjetos = False
    PicView.Visible = False
    For i = 1 To General_Locale_NPCs(0, 3)
    Resultados.AddItem i & " - " & General_Locale_NPCs(i, 0)
    Next i
    
    Buscar(0).Enabled = True
    
    End Select
    
    Buscar(Index).Enabled = False

    
End Sub
Private Sub BuscarUno_Click(Index As Integer)


If Len(Busqueda.Text) < 3 Then
MsgBox "Escribe algo más completo.", vbApplicationModal
Exit Sub
End If

''Limpiamos la lista antes.
Resultados.Clear

Dim i As Integer

Select Case Index

Case 0

BusquedaObjetos = True
PicView.Visible = True
For i = 1 To General_Locale_Obj(0, 4)
If InStr(1, Tilde(General_Locale_Obj(i, 0)), Tilde(Busqueda.Text)) Then
Resultados.AddItem i & " - " & General_Locale_Obj(i, 0)
End If
Next

Case 1

BusquedaObjetos = False
PicView.Visible = False
For i = 1 To General_Locale_NPCs(0, 3)
If InStr(1, Tilde(General_Locale_NPCs(i, 0)), Tilde(Busqueda.Text)) Then
Resultados.AddItem i & " - " & General_Locale_NPCs(i, 0)
End If
Next

End Select

End Sub

Private Sub Form_Load()
   Call FormParser.Parse_Form(Me)
   MensajeBusqueda = True
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Busqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If MensajeBusqueda Then
        Busqueda = vbNullString
        Busqueda.ForeColor = vbBlack
        MensajeBusqueda = False
    End If
End Sub

Private Sub Limpiarlista_Click()
       'Limpiamos la lista antes.
        Resultados.Clear
End Sub

Private Sub Resultados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Index As Long
    Dim PosX As Long, PosY As Long
    
    If BusquedaObjetos = True And PicView.Visible = True Then
    PicView.Cls
    Call DrawGrhtoHdc(PicView.hDC, CInt(General_Locale_Obj(val(Resultados.list(Resultados.ListIndex)), 3)), 0, 0)
    PicView.Refresh
    Else
    PicView.Visible = False
    End If
    
    
    ' Detectamos el clic derecho para simular la seleccion
    If Button = vbRightButton Then
        ' Convertir a pixeles
        PosX = CLng(X / Screen.TwipsPerPixelX)
        PosY = CLng(Y / Screen.TwipsPerPixelY)

        ' Mensaje directo al hWnd usando WinAPI
        Index = SendMessage(Resultados.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((PosY * 65536) + PosX))

        ' Si seleccionamos un item valido
        If Index < Resultados.ListCount Then
            Resultados.ListIndex = Index
        End If
    End If
End Sub

Private Sub Resultados_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
    
    Case vbKeyUp
    If Resultados.ListIndex = 0 Then Exit Sub
    If BusquedaObjetos = True And PicView.Visible = True Then
    PicView.Cls
    Call DrawGrhtoHdc(PicView.hDC, CInt(General_Locale_Obj(val(Resultados.list(Resultados.ListIndex - 1)), 3)), 0, 0)
    PicView.Refresh
    Else
    PicView.Visible = False
    End If
    
    Case vbKeyDown
    If Resultados.ListIndex = General_Locale_Obj(0, 4) - 1 Then Exit Sub
    
    If BusquedaObjetos = True And PicView.Visible = True Then
    PicView.Cls
    Call DrawGrhtoHdc(PicView.hDC, CInt(General_Locale_Obj(val(Resultados.list(Resultados.ListIndex + 1)), 3)), 0, 0)
    PicView.Refresh
    Else
    PicView.Visible = False
    End If
    
    End Select
 


End Sub
              
 
Private Sub Resultados_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Resultados.ListIndex >= 0 Then
        If BusquedaObjetos Then
            PopupMenu mnuCrearO
        Else
            PopupMenu mnuCrearN
        End If
    End If
End Sub

Private Sub mnuCrearObj_Click(Index As Integer)
On Error GoTo ErrHandler

    Dim Numero As Integer
    Dim Cantidad As Integer
    
    Select Case Index
        Case 0
            Cantidad = 1
        Case 1
            Cantidad = 10
        Case 2
            Cantidad = 100
        Case 3
            Cantidad = val(InputBox("Escribe la cantidad.", "Crear Objeto - Link-AO Staff"))
            
            If Cantidad <= 0 Then
                Exit Sub
            ElseIf Cantidad > MAX_INVENTORY_OBJS Then
                Cantidad = MAX_INVENTORY_OBJS
            End If
    End Select

    'Parche para evitar que no se seleccione un item y al querer crearlo explote el juego (Recox)
    If Resultados.ListIndex < 0 Then
        MsgBox "Seleccione objeto"
        Exit Sub
    End If
    
    Numero = val(Resultados.list(Resultados.ListIndex))
    
    If Numero > 0 Then
        Call WriteCreateItem(Numero, Cantidad)
    End If

    Exit Sub

ErrHandler:
    Cantidad = MAX_INVENTORY_OBJS
    Resume Next
End Sub

Private Sub mnuCrearNPC_Click(Index As Integer)
    Dim Numero As Integer

    'Parche para evitar que no se seleccione un item y al querer crearlo explote el juego (Recox)
    If Resultados.ListIndex < 0 Then
        MsgBox "Seleccione NPC"
        Exit Sub
    End If
    
    Numero = val(Resultados.list(Resultados.ListIndex))
    
    If Numero > 0 Then
        Call WriteCreateNPC(Numero)
    End If
End Sub

'Parche: Al cerrar el formulario tambien te desconecta hahahaha ^_^'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
Private Sub Busqueda_change()
Resultados.ListIndex = SendMessages(Resultados.hwnd, LB_FINDSTRING, -1, ByVal Busqueda.Text)
End Sub
