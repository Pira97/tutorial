VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$438"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   3780
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
   Icon            =   "frmEntrenador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstCriaturas 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   480
      TabIndex        =   2
      Top             =   540
      Width           =   2550
   End
   Begin VB.CommandButton imgLuchar 
      Caption         =   "$24"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   1665
   End
   Begin VB.CommandButton imgSalir 
      Caption         =   "$25"
      Height          =   390
      Left            =   2160
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "$23"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1710
      TabIndex        =   3
      Top             =   105
      Width           =   375
   End
End
Attribute VB_Name = "frmEntrenador"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^,
'   le puse borde a la ventana y le cambié la letra a
'   una más linda :)
'
'[END]'

Option Explicit

 
 
Private Sub Form_Load()
Call FormParser.Parse_Form(Me)

End Sub

Private Sub imgLuchar_Click()
    Call WriteTrain(lstCriaturas.ListIndex + 1)
    Unload Me

End Sub

Private Sub imgSalir_Click()
    Unload Me

End Sub

Private Sub lstCriaturas_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
 
End Sub
