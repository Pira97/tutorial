VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   2220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCantidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   300
      MaxLength       =   6
      TabIndex        =   0
      Top             =   540
      Width           =   1470
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   1905
      Tag             =   "0"
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgMas 
      Height          =   135
      Left            =   1800
      Top             =   510
      Width           =   195
   End
   Begin VB.Image imgMenos 
      Height          =   135
      Left            =   1800
      Top             =   630
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   1125
      Tag             =   "0"
      Top             =   840
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   150
      Tag             =   "0"
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************f
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free  Foundation version 2.1 of
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
Private Sub Form_Deactivate()
Unload Me
End Sub
Private Sub Form_Load()
    Me.Picture = General_Load_Skin_Picture_From_Resource_Ex("cantidad")
    Make_Transparent_Form Me.hwnd, 210
    Call FormParser.Parse_Form(Me)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Tag = "1" Then
    Image1.Picture = Nothing
    Image1.Tag = "0"
End If

If Image2.Tag = "1" Then
    Image2.Picture = Nothing
    Image2.Tag = "0"
End If

If Image3.Tag = "1" Then
    Image3.Picture = Nothing
    Image3.Tag = "0"
End If
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrarcantdown")
End Sub
Private Sub image3_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Tag = "1" Then
    Image1.Picture = Nothing
    Image1.Tag = "0"
End If

If Image2.Tag = "1" Then
    Image2.Picture = Nothing
    Image2.Tag = "0"
End If

If Image3.Tag = "0" Then
    Image3.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrarcantover")
    Image3.Tag = "1"
End If
End Sub
Private Sub image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = General_Load_Skin_Picture_From_Resource_Ex("dejardown")
End Sub

Private Sub image1_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image2.Tag = "1" Then
    Image2.Picture = Nothing
    Image2.Tag = "0"
End If

If Image1.Tag = "0" Then
    Image1.Picture = General_Load_Skin_Picture_From_Resource_Ex("dejarover")
    Image1.Tag = "1"
End If

End Sub
 Private Sub imgMas_Click()
Text1.Text = val(Text1.Text) + 1
End Sub
Private Sub imgMas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub imgMenos_Click()
If val(Text1.Text) > 0 Then _
    Text1.Text = val(Text1.Text) - 1
End Sub
Private Sub imgMenos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = General_Load_Skin_Picture_From_Resource_Ex("dejartododown")
End Sub

Private Sub image2_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Tag = "0" Then
    Image2.Picture = General_Load_Skin_Picture_From_Resource_Ex("dejartodoover")
    Image2.Tag = "1"
End If

If Image1.Tag = "1" Then
    Image1.Picture = Nothing
    Image1.Tag = "0"
End If
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Inventario.SelectedItem = 0 Then Exit Sub
    
    CantidadGlobal = Inventario.Amount(Inventario.SelectedItem)
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
        Unload Me
    Else
        If CurrentUser.UserGLD >= 100000 Then
            Call WriteDrop(Inventario.SelectedItem, 100000)
            Me.Visible = False
        Else
            Call WriteDrop(Inventario.SelectedItem, CurrentUser.UserGLD)
            Unload Me
        End If
    End If
    frmCantidad.Text1.Text = ""
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Inventario.SelectedItem = 0 Then Exit Sub
    If LenB(frmCantidad.Text1.Text) = 0 Then Me.Visible = False

    If LenB(frmCantidad.Text1.Text) > 0 Then
        If Not IsNumeric(frmCantidad.Text1.Text) Then Exit Sub  'Should never happen
        CantidadGlobal = frmCantidad.Text1.Text
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.Text1.Text)
        frmCantidad.Text1.Text = "1"
    End If

    Me.Visible = False
End Sub
Private Sub texT1_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtCant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Text1_Change()
On Error GoTo ErrHandler
    If val(Text1.Text) < 0 Then
        Text1.Text = "1"
    End If
    
    If val(Text1.Text) > MAX_INVENTORY_OBJS Then
        Text1.Text = "10000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub

 

 
