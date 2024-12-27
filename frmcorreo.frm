VERSION 5.00
Begin VB.Form frmCorreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$475"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   Icon            =   "frmcorreo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "$477"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   7215
      Begin VB.TextBox txCantidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   15
         Text            =   "1"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "$28"
         Height          =   495
         Left            =   5520
         TabIndex        =   14
         Top             =   2560
         Width           =   1575
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "$482"
         Height          =   495
         Left            =   5520
         TabIndex        =   13
         Top             =   1960
         Width           =   1575
      End
      Begin VB.CheckBox adjItem 
         Caption         =   "$480"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   160
         Width           =   1575
      End
      Begin VB.TextBox txSndMsg 
         Height          =   1875
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txTo 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.ListBox lstInv 
         Enabled         =   0   'False
         Height          =   2595
         ItemData        =   "frmcorreo.frx":000C
         Left            =   2880
         List            =   "frmcorreo.frx":0046
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.PictureBox picInvT 
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   5520
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lbCount 
         AutoSize        =   -1  'True
         Caption         =   "$206"
         Height          =   195
         Left            =   6120
         TabIndex        =   16
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "$481"
         Height          =   195
         Left            =   5520
         TabIndex        =   12
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "$206"
         Height          =   195
         Left            =   5520
         TabIndex        =   11
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "$479"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$478"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$486"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox picItem 
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   2640
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   19
         Top             =   1920
         Width           =   540
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "$483"
         Height          =   495
         Left            =   5640
         TabIndex        =   18
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "$484"
         Height          =   495
         Left            =   5640
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txMensaje 
         Enabled         =   0   'False
         Height          =   1575
         Left            =   2640
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstMsg 
         Height          =   2790
         ItemData        =   "frmcorreo.frx":00B6
         Left            =   120
         List            =   "frmcorreo.frx":00B8
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2640
         TabIndex        =   21
         Top             =   2520
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lbCant 
         AutoSize        =   -1  'True
         Caption         =   "$206"
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelectMsg As Byte

Public Sub ActualizarCorreo()

    PicItem.Cls
    PicItem.Refresh
    
    If lstMsg.list(lstMsg.ListIndex) <> "(" & Locale_GUI_Frase(269) & ")" And lstMsg.list(lstMsg.ListIndex) <> "" Then
    
        txMensaje.Text = Correos(SelectMsg).Mensaje
        
        If Correos(SelectMsg).Cantidad <> 0 Then
        
            PicItem.Visible = True
            cmdSave.Enabled = True
            lbCant.Visible = True
            lbCant.Caption = Locale_GUI_Frase(206) & ": " & Correos(SelectMsg).Cantidad
            lbCant.Caption = Correos(SelectMsg).Nombre & " (" & Correos(SelectMsg).Cantidad & ")"
            Label5.Visible = True
            Label5.Caption = General_Locale_Obj(Correos(SelectMsg).GrhIndex, 1)
            DrawGrhtoHdc PicItem.hDC, Correos(SelectMsg).GrhIndex, 0, 0
            
        Else
            PicItem.Visible = False
            lbCant.Caption = vbNullString
            Label5.Caption = vbNullString
            cmdSave.Enabled = False
        End If
        
    Else
        lbCant.Caption = vbNullString
        txMensaje.Text = vbNullString
        Label5.Caption = vbNullString
        PicItem.Visible = False
        cmdSave.Enabled = False
    End If
    
End Sub

Private Sub adjItem_Click()

    If adjItem.Value = vbChecked Then
        lstInv.Enabled = True
        txCantidad.Enabled = True
        Label4.Caption = Locale_GUI_Frase(206) & ":1750"
    Else
        txCantidad.Enabled = False
        lstInv.Enabled = False
        Label4.Caption = Locale_GUI_Frase(481) & ":"
    End If
End Sub

Private Sub cmdClean_Click()

    txTo.Text = vbNullString
    txSndMsg.Text = vbNullString
    txCantidad.Text = vbNullString
    
End Sub

Private Sub cmdDel_Click()

If MsgBox(Locale_GUI_Frase(544), vbYesNo + vbQuestion) = vbYes Then
    Call WritePacketsCorreo(1, lstMsg.ListIndex + 1)
    Unload Me
End If
    
End Sub

Private Sub cmdSave_Click()
Call WritePacketsCorreo(3, lstMsg.ListIndex + 1)
Unload Me
End Sub

Private Sub cmdSend_Click()

If txTo.Text = vbNullString Then Exit Sub
'No se puede auto enviar correos!
'If UCase(txTo.Text) = ucase(currentuser.username) Then
'   frmmensaje.msg.Caption = "No puedes enviarte correos a ti mismo."
'   frmmensaje.Show , Me
'   Exit Sub
 
'End If
'Shermie
 
If adjItem.Value = vbChecked Then
    If lstInv.ListIndex = -1 Then
        Call MsgBox(Locale_GUI_Frase(540), vbInformation, Locale_GUI_Frase(475))
    Else
        If Not txCantidad.Text = vbNullString Then
            Call WriteEnviarCorreo(txTo.Text, txSndMsg.Text, lstInv.ListIndex + 1, txCantidad.Text)
            Unload Me
        Else
            Call MsgBox(Locale_GUI_Frase(220), vbInformation, Locale_GUI_Frase(475))
        End If
    End If
Else
    If txSndMsg.Text = vbNullString And adjItem.Value = vbUnchecked Then
        Call MsgBox("Debes enviar un mensaje.", vbInformation, Locale_GUI_Frase(475))
    Else
        Call WriteEnviarCorreo(txTo.Text, txSndMsg.Text, 0, 0)
        Unload Me
    End If
End If

End Sub

Private Sub Form_Load()
 
Call FormParser.Parse_Form(Me)

Dim i As Integer

lstInv.Clear

For i = 1 To MAX_INVENTORY_SLOTS
    If Inventario.ItemName(i) <> "" Then
        lstInv.AddItem Inventario.ItemName(i)
    Else
        lstInv.AddItem Locale_GUI_Frase(269)
    End If
Next i

End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
lstInv.Visible = True
End Sub

Private Sub lstInv_Click()

picInvT.Cls
DrawGrhtoHdc picInvT.hDC, Inventario.GrhIndex(lstInv.ListIndex + 1), 0, 0
lbCount.Caption = Locale_GUI_Frase(206) & ":" & Inventario.Amount(lstInv.ListIndex + 1)

End Sub

Private Sub lstMsg_Click()

SelectMsg = lstMsg.ListIndex + 1

If SelectMsg = 0 Then Exit Sub

Call ActualizarCorreo
Call WritePacketsCorreo(2, SelectMsg)

If Correos(SelectMsg).Leido = 0 Then
    lstMsg.list(lstMsg.ListIndex) = Correos(SelectMsg).De
    Correos(SelectMsg).Leido = 1
End If

End Sub

Private Sub txCantidad_Change()

If val(txCantidad.Text) < 0 Then
    txCantidad.Text = 1
End If

If val(txCantidad.Text) > 10000 Then
    txCantidad.Text = 1
End If

End Sub

