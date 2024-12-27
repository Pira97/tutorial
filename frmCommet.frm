VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   Icon            =   "frmCommet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton imgEnviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   2400
      MouseIcon       =   "frmCommet.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton imgCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmCommet.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

 
Private Const MAX_PROPOSAL_LENGTH As Integer = 520
 

Public Nombre                     As String

Public t                          As Tipo

Public Enum Tipo

    ALIANZA = 1
    PAZ = 2
    RECHAZOPJ = 3

End Enum

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)

End Sub

Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgEnviar_Click()

    If Text1 = "" Then
        If t = PAZ Or t = ALIANZA Then
            MsgBox "Debes redactar un mensaje solicitando la paz o alianza al l�der de " & Nombre
        Else
            MsgBox "Debes indicar el motivo por el cual rechazas la membres�a de " & Nombre

        End If
        
        Exit Sub

    End If
    
    If t = PAZ Then
        Call WriteGuildOfferPeace(Nombre, Replace(Text1, vbCrLf, "�"))
        
    ElseIf t = ALIANZA Then
        Call WriteGuildOfferAlliance(Nombre, Replace(Text1, vbCrLf, "�"))
        
    ElseIf t = RECHAZOPJ Then
        Call WriteGuildRejectNewMember(Nombre, Replace(Replace(Text1.Text, ",", " "), vbCrLf, " "))

        'Sacamos el char de la lista de aspirantes
        Dim i As Long
        
        For i = 0 To frmGuildLeader.solicitudes.ListCount - 1

            If frmGuildLeader.solicitudes.list(i) = Nombre Then
                frmGuildLeader.solicitudes.RemoveItem i
                Exit For

            End If

        Next i
        
        Me.Hide
        Unload frmCharInfo

    End If
    
    Unload Me

End Sub

Private Sub Text1_Change()

    If Len(Text1.Text) > MAX_PROPOSAL_LENGTH Then Text1.Text = Left$(Text1.Text, MAX_PROPOSAL_LENGTH)

End Sub
