VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "$436"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3525
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
   Icon            =   "frmGuildAdm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "$34"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.ListBox GuildsList 
         Height          =   2985
         ItemData        =   "frmGuildAdm.frx":000C
         Left            =   120
         List            =   "frmGuildAdm.frx":0013
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$35"
      Height          =   375
      Left            =   2280
      MouseIcon       =   "frmGuildAdm.frx":0023
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFundar 
      Caption         =   "$36"
      Height          =   375
      Left            =   1200
      MouseIcon       =   "frmGuildAdm.frx":0175
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "$25"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildAdm.frx":02C7
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtBuscar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   375
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   360
      Tag             =   "1"
      Top             =   5985
      Width           =   855
   End
End
Attribute VB_Name = "frmGuildAdm"
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

 
Private Sub cmdFundar_Click()
'Call Clienttcp.ParseUserCommand("/fundarclan")
Unload Me
End Sub

Private Sub Command1_Click()
frmGuildBrief.EsLeader = False

    Call WriteGuildRequestDetails(guildslist.list(guildslist.ListIndex))

End Sub

Private Sub Command3_Click()
Unload Me
frmMain.SetFocus
End Sub

Private Sub Form_Load()
    
   ' Me.Picture = LoadPicture(App.path & "\graficos\VentanaListaClanes.jpg")
    Call FormParser.Parse_Form(Me)

     
End Sub
 
Private Sub guildslist_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
 
End Sub


Private Sub txtBuscar_Change()
    Call FiltrarListaClanes(txtBuscar.Text)

End Sub

Private Sub txtBuscar_GotFocus()

    With txtBuscar
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    If UBound(GuildNames) <> 0 Then

        With guildslist
            'Limpio la lista
            .Clear
            
            .Visible = False
            
            ' Recorro los arrays
            For lIndex = 0 To UBound(GuildNames)

                ' Si coincide con los patrones
                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                    ' Lo agrego a la lista
                    .AddItem GuildNames(lIndex)

                End If

            Next lIndex
            
            .Visible = True

        End With

    End If

End Sub
