VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "LinkAO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrpres 
      Interval        =   3000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picLoad 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   1320
      ScaleHeight     =   90
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   8640
      Width           =   15
   End
   Begin InetCtlsObjects.Inet mainInet 
      Left            =   11250
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   15
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
' soporte@Link-AO.com.ar
'   - Relase Number 1
'*****************************************************************

Option Explicit
Dim f As Integer
Private porcentajeActual As Integer
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 499
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
Private Sub Form_Load()
Me.Caption = Form_Caption
'Me.Picture = General_Load_Picture_From_Resource_Ex("_40")
'picLoad.Picture = General_Load_Picture_From_Resource_Ex("_39")
    
Call FormParser.Parse_Form(Me, E_WAIT)
End Sub
Function Analizar()
            On Error Resume Next
           
            Dim iX As Integer
            Dim TX As Integer
            Dim DifX As Integer
           
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
'    iX = mainInet.OpenURL("http://coveronline.000webhostapp.com/cao/VEREXE.txt") 'Host
    TX = LeerInt(App.Path & "\INIT\Update.ini")
    DifX = iX - TX
 
            If Not (DifX = 0) Then
 MsgBox "Tu cliente no está actualizado, por favor, ejecuta el autoupdate y actualiza."
 
CerrarJuego
      
End If
End Function
Private Function LeerInt(ByVal Ruta As String) As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function
 
Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    f = FreeFile
    Open Ruta For Output As f
    Print #f, data
    Close #f
End Sub
Public Sub EstablecerProgreso(ByVal nuevoPorcentaje As Integer)
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    picLoad.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    picLoad.Width = DEFAULT_PROGRESS_WIDTH
Else
    picLoad.Width = 0
End If
porcentajeActual = nuevoPorcentaje
End Sub

Public Sub progresoConDelay(ByVal porcentaje As Integer)
If porcentaje = porcentajeActual Then Exit Sub
Dim step As Integer, stepInterval As Integer, Timer As Long, tickCount As Long
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
Do Until CompararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (Timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    Timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call EstablecerProgreso(porcentajeActual)
Loop
End Sub
Private Function CompararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    CompararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    CompararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
End Function


