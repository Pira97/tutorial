VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración del Clan"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5880
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
   Icon            =   "frmGuildLeader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar Clan"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildLeader.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   22
      Top             =   3960
      Width           =   2895
      Begin VB.ListBox solicitudes 
         Height          =   840
         ItemData        =   "frmGuildLeader.frx":015E
         Left            =   120
         List            =   "frmGuildLeader.frx":0160
         TabIndex        =   24
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton imgDetallesSolicitudes 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0162
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1170
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1620
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   19
      Top             =   0
      Width           =   2895
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":02B4
         Left            =   120
         List            =   "frmGuildLeader.frx":02B6
         TabIndex        =   21
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton imgDetallesMiembros 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":02B8
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   5775
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton imgActualizar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":040A
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1080
         Width           =   5535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton imgDetallesClan 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":055C
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":06AE
         Left            =   120
         List            =   "frmGuildLeader.frx":06B0
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton imgEditarCodex 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":06B2
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3930
      Width           =   2775
   End
   Begin VB.CommandButton imgEditarURL 
      Caption         =   "Editar URL de la web del clan"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0804
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton imgPropuestasPaz 
      Caption         =   "Propuestas de paz"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0956
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   4950
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton imgCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0AA8
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5970
      Width           =   2775
   End
   Begin VB.CommandButton imgPropuestasAlianzas 
      Caption         =   "Propuestas de alianzas"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0BFA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   6840
      MouseIcon       =   "frmGuildLeader.frx":0D4C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   6360
      Width           =   2655
   End
   Begin VB.TextBox txtFiltrarMiembros1 
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
      Height          =   225
      Left            =   9795
      TabIndex        =   6
      Top             =   2220
      Width           =   2580
   End
   Begin VB.TextBox txtFiltrarClanes1 
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
      Left            =   6915
      TabIndex        =   5
      Top             =   2220
      Width           =   2580
   End
   Begin VB.TextBox txtguildnews1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   690
      Left            =   6915
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3315
      Width           =   5475
   End
   Begin VB.ListBox solicitudes1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":0E9E
      Left            =   6915
      List            =   "frmGuildLeader.frx":0EA5
      TabIndex        =   2
      Top             =   4980
      Width           =   2595
   End
   Begin VB.ListBox members1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0EB7
      Left            =   9780
      List            =   "frmGuildLeader.frx":0EBE
      TabIndex        =   1
      Top             =   420
      Width           =   2595
   End
   Begin VB.ListBox guildslist1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0ECC
      Left            =   6900
      List            =   "frmGuildLeader.frx":0ED3
      TabIndex        =   0
      Top             =   420
      Width           =   2595
   End
   Begin VB.Image imgCerrar1 
      Height          =   495
      Left            =   9720
      Tag             =   "1"
      Top             =   6585
      Width           =   2775
   End
   Begin VB.Image imgPropuestasAlianzas1 
      Height          =   495
      Left            =   9720
      Tag             =   "1"
      Top             =   6075
      Width           =   2775
   End
   Begin VB.Image imgPropuestasPaz1 
      Height          =   495
      Left            =   9720
      Tag             =   "1"
      Top             =   5565
      Width           =   2775
   End
   Begin VB.Image imgEditarURL1 
      Height          =   495
      Left            =   9720
      Tag             =   "1"
      Top             =   5055
      Width           =   2775
   End
   Begin VB.Image imgEditarCodex1 
      Height          =   495
      Left            =   9720
      Tag             =   "1"
      Top             =   4545
      Width           =   2775
   End
   Begin VB.Image imgActualizar1 
      Height          =   390
      Left            =   6870
      Tag             =   "1"
      Top             =   4110
      Width           =   5550
   End
   Begin VB.Image imgDetallesSolicitudes1 
      Height          =   375
      Left            =   6840
      Tag             =   "1"
      Top             =   5925
      Width           =   2655
   End
   Begin VB.Image imgDetallesMiembros1 
      Height          =   375
      Left            =   9780
      Tag             =   "1"
      Top             =   2580
      Width           =   2655
   End
   Begin VB.Image imgDetallesClan1 
      Height          =   375
      Left            =   6885
      Tag             =   "1"
      Top             =   2580
      Width           =   2655
   End
   Begin VB.Image imgElecciones1 
      Height          =   375
      Left            =   6840
      Tag             =   "1"
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   8535
      TabIndex        =   3
      Top             =   6390
      Width           =   255
   End
End
Attribute VB_Name = "frmGuildLeader"
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

Private Const MAX_NEWS_LENGTH     As Integer = 512

 

Private Sub Command1_Click()
Select Case MsgBox("¿Esta seguro que desea eliminar tu clan?, una vez eliminado no se podrá recuperar.", vbOKCancel Or vbInformation Or vbDefaultButton1, "Mensaje")

    Case vbOK
 Call writeCloseGuild
 Unload Me
         
    Case vbCancel

End Select

 
End Sub

Private Sub Form_Load()
 
   ' Me.Picture = LoadPicture(App.path & "\graficos\VentanaAdministrarClan.jpg")
    


End Sub
 
Private Sub guildslist_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)
End Sub

Private Sub imgActualizar_Click()

    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)

End Sub

Private Sub imgCerrar_Click()
    Unload Me
    frmMain.SetFocus

End Sub

Private Sub imgDetallesClan_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.list(guildslist.ListIndex))

End Sub

Private Sub imgDetallesMiembros_Click()

    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.list(members.ListIndex))

End Sub

Private Sub imgDetallesSolicitudes_Click()

    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.list(solicitudes.ListIndex))

End Sub

Private Sub imgEditarCodex_Click()
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)

End Sub

Private Sub imgEditarURL_Click()
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)

End Sub

Private Sub imgElecciones_Click()
    Call WriteGuildOpenElections
    Unload Me

End Sub

Private Sub imgPropuestasAlianzas_Click()
    Call WriteGuildAlliancePropList

End Sub

Private Sub imgPropuestasPaz_Click()
    Call WriteGuildPeacePropList

End Sub

Private Sub members_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)
 
End Sub

Private Sub solicitudes_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
 
End Sub

Private Sub txtguildnews_Change()

    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)

End Sub

Private Sub txtguildnews_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   y As Single)
 
End Sub

Private Sub txtFiltrarClanes_Change()
    'Call FiltrarListaClanes(txtFiltrarClanes.Text)

End Sub

Private Sub txtFiltrarClanes_GotFocus()

    'With txtFiltrarClanes
    '    .SelStart = 0
    '    .SelLength = Len(.Text)

   ' End With

End Sub

Private Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
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

End Sub

Private Sub txtFiltrarMiembros_Change()
   ' Call FiltrarListaMiembros(txtFiltrarMiembros.Text)

End Sub

Private Sub txtFiltrarMiembros_GotFocus()

    'With txtFiltrarMiembros
    '    .SelStart = 0
   '    .SelLength = Len(.Text)
'
    'End With

End Sub

Private Sub FiltrarListaMiembros(ByRef sCompare As String)

End Sub

 '   Dim lIndex As Long
    
 '   With members
        'Limpio la lista
 '       .Clear
        
 '       .Visible = False
        
        ' Recorro los arrays
 '       For lIndex = 0 To UBound(GuildMembers)

            ' Si coincide con los patrones
 '           If InStr(1, UCase$(GuildMembers(lIndex)), UCase$(sCompare)) Then
'                ' Lo agrego a la lista
'                .AddItem GuildMembers(lIndex)

'            End If

'        Next lIndex
        
'        .Visible = True

'    End With

'End Sub

