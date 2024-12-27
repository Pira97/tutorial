VERSION 5.00
Begin VB.Form frmGoliath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$297"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
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
   Icon            =   "frmBanco2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBanco 
      Height          =   840
      ItemData        =   "frmBanco2.frx":000C
      Left            =   90
      List            =   "frmBanco2.frx":001C
      TabIndex        =   1
      Top             =   1230
      Width           =   4395
   End
   Begin VB.TextBox txtDatos 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "$1"
      Height          =   345
      Left            =   3000
      TabIndex        =   5
      Top             =   3090
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "$2"
      Height          =   329
      Left            =   120
      TabIndex        =   4
      Top             =   3090
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBanco2.frx":007E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   987
      Left            =   90
      TabIndex        =   0
      Top             =   86
      Width           =   4395
   End
   Begin VB.Label lblDatos 
      Caption         =   "$29"
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
End
Attribute VB_Name = "frmGoliath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private EtapaTransferencia As Byte
Private Usuario As String
Private Cantidad As String
Private Oro As Long
Private Items As Long
Private Sub cmdClose_Click()
Call WriteBankEnd
    Unload Me
End Sub

Private Sub CmdOk_Click()
Call Audio.PlayWave(SND_CLICK)
Select Case lstBanco.ListIndex

    Case 0, -1  'Depositar
    
        'Si es negativo o cero jodete por pobre xD
        If val(txtDatos.Text) <= 0 Then
            lblDatos.Caption = "Cantidad inválida."
            Exit Sub
        End If
        
        If val(txtDatos.Text) > CurrentUser.UserGLD Then
            lblDatos.Caption = "No tienes esa cantidad. Escríbela nuevamente."
            Exit Sub
        Else
           ' Call Clienttcp.ParseUserCommand("/DEPOSITAR " & val(txtDatos.Text))
            lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & CurrentUser.UserGLD & " monedas de oro en tu billetera y en tu cuenta tienes " & Oro & " Monedas de oro. ¿Cómo te puedo ayudar?"
         End If
         Call WriteBankEnd
         Unload Me
         
    Case 1 'Retirar
    
        'Si es negativo o cero jodete por pobre xD
        If val(txtDatos.Text) <= 0 Then
            lblDatos.Caption = "Cantidad inválida."
            Exit Sub
        End If
        
       'Call Clienttcp.ParseUserCommand("/RETIRAR " & val(txtDatos.Text))
            lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & Items & " monedas de oro en tu billetera y en tu cuenta tienes " & CurrentUser.UserGLD & " Monedas de oro." & "¿Cómo te puedo ayudar?"
       Call WriteBankEnd
 
        Unload Me
        
    Case 2 'Bóveda
        Call WriteBankStart
        Unload Me
    
    Case 3 'trasferir oro Mermas :p
    On Local Error GoTo error
    
    If EtapaTransferencia = 0 Then
    
                'Negativos y ceros
        If val(txtDatos.Text) <= 0 Then
            lblDatos.Caption = "Cantidad inválida."
            txtDatos.Text = vbNullString
            Exit Sub
        End If
            
        If val(txtDatos.Text) >= Items Then
            Cantidad = val(txtDatos.Text)
            lblDatos.Caption = "¿A quién le deseas enviar " & Cantidad & " monedas de oro?"
            EtapaTransferencia = 1
            txtDatos.Text = vbNullString
        Else
            lblDatos.Caption = "No tienes esa cantidad"
            txtDatos.Text = vbNullString
        End If
        
    ElseIf EtapaTransferencia = 1 Then
        If LenB(txtDatos.Text) > 0 Then
            Usuario = txtDatos.Text
            lblDatos.Caption = Cantidad & " monedas de oro serán transferidas a " & Usuario & " , si es correcto presione aceptar."
            EtapaTransferencia = 2
        Else
            lblDatos.Caption = "No tienes esa cantidad"
            txtDatos.Text = vbNullString
        End If
        
    ElseIf EtapaTransferencia = 2 Then
    Call WriteTransferGold(Usuario, Cantidad)
    Call WriteBankEnd
    Unload Me
    End If
 
error:
    Exit Sub
    Unload Me
    End Select

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & CurrentUser.UserGLD & " monedas de oro en tu billetera y en tu cuenta tienes " & Oro & " Monedas de oro. y " & Items & " items en tu Boveda. ¿Cómo te puedo ayudar?"
End Sub

Private Sub lstBanco_Click()

Select Case lstBanco.ListIndex
    Case 0 'Depositar oro
        lblDatos.Caption = "¿Cuánto deseas depositar?"
        txtDatos.Visible = True
    Case 1 'Retirar oro
        lblDatos.Caption = "¿Cuánto deseas retirar?"
        txtDatos.Visible = True
    Case 2 'ver la Boveda
        lblDatos.Caption = "Presiona aceptar para ver tu boveda."
        txtDatos.Visible = False
    Case 3 'Transferir oro
        lblDatos.Caption = "¿Qué cantidad desea transferir?"
        txtDatos.Visible = True
         EtapaTransferencia = 0
End Select

End Sub

Public Sub ParseBancoInfo(ByVal GLD As Long, ByVal item As Long)

On Error GoTo Error_Handler

Oro = GLD
Items = item

Me.Show vbModeless, frmMain

Exit Sub

Error_Handler:
    'Error vite'

End Sub

