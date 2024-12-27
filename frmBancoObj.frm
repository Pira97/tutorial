VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ControlBox      =   0   'False
   FillColor       =   &H00877365&
   Icon            =   "frmBancoObj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   3960
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   2460
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   3960
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   2520
      Width           =   2490
   End
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1620
      Width           =   480
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Text            =   "1"
      Top             =   6885
      Width           =   510
   End
   Begin VB.Image imgCerrar 
      Height          =   345
      Left            =   6480
      Tag             =   "1"
      Top             =   180
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   1965
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1830
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   195
      Index           =   2
      Left            =   5460
      TabIndex        =   1
      Top             =   1530
      Width           =   675
   End
   Begin VB.Label Label1 
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
      Height          =   435
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1530
      Width           =   2985
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3855
      Tag             =   "1"
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2955
      Tag             =   "1"
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image image1 
      Height          =   450
      Index           =   1
      Left            =   4230
      Tag             =   "1"
      Top             =   6780
      Width           =   2175
   End
   Begin VB.Image image1 
      Height          =   450
      Index           =   0
      Left            =   585
      Tag             =   "1"
      Top             =   6780
      Width           =   2175
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmBancoObj - Link-AO - v1.0.0 R1
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
Public LasActionBuy        As Boolean
Public LastIndex1          As Integer
Public LastIndex2          As Integer

Private m_Number As Integer
Private m_Increment As Integer
Private m_Interval As Integer
Private Sub cantidad_Change()

If val(Cantidad.Text) < 0 Then
    Cantidad.Text = 1
    m_Number = 1
ElseIf val(Cantidad.Text) > MAX_INVENTORY_OBJS Then
    Cantidad.Text = 1
    m_Number = 1
Else
    m_Number = val(Cantidad.Text)
End If

End Sub
Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
Private Sub cantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        cmdMasMenos(Index).Picture = General_Load_Picture_From_Resource_Ex("_50")
        cmdMasMenos(Index).Tag = "1"
        m_Increment = -1
    Case 1
        cmdMasMenos(Index).Picture = General_Load_Picture_From_Resource_Ex("_52")
        cmdMasMenos(Index).Tag = "1"
        m_Increment = 1
End Select

tmrNumber.Interval = 30
tmrNumber.Enabled = True

End Sub
Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case 0
        If cmdMasMenos(Index).Tag = "0" Then
        cmdMasMenos(0).Picture = General_Load_Picture_From_Resource_Ex("_51")
        cmdMasMenos(Index).Tag = "1"
        End If
    
    Case 1
        If cmdMasMenos(Index).Tag = "0" Then
        cmdMasMenos(1).Picture = General_Load_Picture_From_Resource_Ex("_53")
        cmdMasMenos(Index).Tag = "1"
        End If
        
    End Select
End Sub
Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
tmrNumber.Enabled = False
End Sub
Private Sub imgCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Audio.PlayWave(SND_CLICK)
 imgCerrar.Picture = General_Load_Picture_From_Resource_Ex("_48")
 imgCerrar.Tag = "1"
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgCerrar.Tag = "0" Then
 imgCerrar.Picture = General_Load_Picture_From_Resource_Ex("_49")
 imgCerrar.Tag = "1"
End If
End Sub
Private Sub imgCerrar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
    Call WriteBankEnd
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (Button = vbLeftButton) Then
        Call Auto_Drag(Me.hwnd)
    Else
        Call WriteBankEnd
    End If

End Sub
Private Sub Form_Load()
    'Cargamos la interfase
    Me.Picture = General_Load_Picture_From_Resource_Ex("_43")
    
    m_Number = 1
    m_Interval = 30
    Call FormParser.Parse_Form(Me)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = "1" Then
    Image1(0).Picture = Nothing
    Image1(0).Tag = "0"
End If

If Image1(1).Tag = "1" Then
    Image1(1).Picture = Nothing
    Image1(1).Tag = "0"
End If

If cmdMasMenos(0).Tag = "1" Then
    cmdMasMenos(0).Picture = Nothing
    cmdMasMenos(0).Tag = "0"
End If

If cmdMasMenos(1).Tag = "1" Then
    cmdMasMenos(1).Picture = Nothing
    cmdMasMenos(1).Tag = "0"
End If

If imgCerrar.Tag = "1" Then
    imgCerrar.Picture = Nothing
    imgCerrar.Tag = "0"
End If

End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 Call Audio.PlayWave(SND_CLICK)
 Select Case Index
 Case 0
 Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("_56")
 Image1(Index).Tag = "1"
 Case 1
 Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("_54")
 Image1(Index).Tag = "1"
 End Select
 
End Sub
Private Sub image1_mousemove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
 Case 0
    If Image1(Index).Tag = "0" Then
    Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("_57")
    Image1(Index).Tag = "1"
    End If
    
 Case 1
    If Image1(Index).Tag = "0" Then
    Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("_55")
    Image1(Index).Tag = "1"
    End If
    
 End Select
End Sub
Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Form_MouseMove(Button, Shift, X, Y)
    
    If List1(Index).list(List1(Index).ListIndex) = "(" & Locale_GUI_Frase(269) & ")" Or _
       List1(Index).ListIndex < 0 Then Exit Sub
    
    If Not IsNumeric(Cantidad.Text) Or Cantidad.Text = 0 Then Exit Sub
    
    Select Case Index
        Case 0
            frmBancoObj.List1(0).SetFocus
            LastIndex1 = List1(0).ListIndex
            LasActionBuy = True
            Call WriteBankExtractItem(List1(0).ListIndex + 1, Cantidad.Text)
            
       Case 1
            frmBancoObj.List1(1).SetFocus
            LastIndex2 = List1(1).ListIndex
            LasActionBuy = False
            Call WriteBankDeposit(List1(1).ListIndex + 1, Cantidad.Text)
            
    End Select
 
End Sub

Private Sub list1_Click(Index As Integer)

    Select Case Index
        Case 0
            Label1(0).Caption = UserBancoInventory(List1(0).ListIndex + 1).Name
            Label1(2).Caption = UserBancoInventory(List1(0).ListIndex + 1).Amount
            
            
        If Label1(2).Caption <> 0 Then
            Select Case UserBancoInventory(List1(0).ListIndex + 1).ObjType
                Case eObjType.otWeapon, eObjType.otFlechas, eObjType.otNudillos
                    Label1(3).Caption = "Golpe: " & UserBancoInventory(List1(0).ListIndex + 1).MinHit & "/" & UserBancoInventory(List1(0).ListIndex + 1).MaxHit & "."
                    Label1(3).Visible = True
                Case eObjType.otArmadura, eObjType.otCASCO, eObjType.otESCUDO
                    Label1(3).Caption = "Defensa: " & UserBancoInventory(List1(0).ListIndex + 1).MinDef & "/" & UserBancoInventory(List1(0).ListIndex + 1).MaxDef & "."
                    Label1(3).Visible = True
                Case eObjType.otMonturas, eObjType.otBarcos
                    Label1(3).Caption = "Defensa: " & UserBancoInventory(List1(0).ListIndex + 1).MinDef & "/" & UserBancoInventory(List1(0).ListIndex + 1).MaxDef & "."
                    Label1(3).Caption = Label1(3).Caption & vbCrLf & "Golpe: " & UserBancoInventory(List1(0).ListIndex + 1).MinHit & "/" & UserBancoInventory(List1(0).ListIndex + 1).MaxHit & "."
                    Label1(3).Visible = True
                Case Else
                    Label1(3).Caption = General_Locale_Obj(UserBancoInventory(List1(0).ListIndex + 1).OBJIndex, 1)
                    Label1(3).Visible = True
            End Select
            
            If UserBancoInventory(List1(0).ListIndex + 1).Amount <> 0 Then
                Call DrawGrhtoHdc(Picture1.hDC, UserBancoInventory(List1(0).ListIndex + 1).GrhIndex, 0, 0)
            Else
                Picture1.Picture = Nothing
            End If
        End If
            
        Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).ListIndex + 1)
        
    If Label1(2).Caption <> 0 Then
        Select Case Inventario.ObjType(List1(1).ListIndex + 1)
                 Case eObjType.otWeapon, eObjType.otFlechas, eObjType.otNudillos
                    Label1(3).Caption = "Golpe: " & Inventario.MinHit(List1(1).ListIndex + 1) & "/" & Inventario.MaxHit(List1(1).ListIndex + 1) & "."
                    Label1(3).Visible = True
                 Case eObjType.otArmadura, eObjType.otCASCO, eObjType.otESCUDO
                    Label1(3).Caption = "Defensa: " & Inventario.MinDef(List1(1).ListIndex + 1) & "/" & Inventario.MaxDef(List1(1).ListIndex + 1) & "."
                    Label1(3).Visible = True
                 Case eObjType.otMonturas, eObjType.otBarcos
                    Label1(3).Caption = "Defensa: " & Inventario.MinDef(List1(1).ListIndex + 1) & "/" & Inventario.MaxDef(List1(1).ListIndex + 1) & "."
                    Label1(3).Caption = Label1(3).Caption & vbCrLf & "Golpe: " & Inventario.MinHit(List1(1).ListIndex + 1) & "/" & Inventario.MaxHit(List1(1).ListIndex + 1) & "."
                    Label1(3).Visible = True
                 Case Else
                    Label1(3).Caption = General_Locale_Obj(Inventario.OBJIndex(List1(1).ListIndex + 1), 1)
                    Label1(3).Visible = True
                    
                    
                      
        End Select
            
             If Inventario.GrhIndex(List1(1).ListIndex + 1) <> 0 Then
                Call DrawGrhtoHdc(Picture1.hDC, Inventario.GrhIndex(List1(1).ListIndex + 1), 0, 0)
            Else
                Picture1.Picture = Nothing
            End If
    End If
    
    End Select
 
 If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
    Label1(2).Caption = ""
    Label1(3).Visible = False
    Picture1.Visible = False
Else
    Picture1.Visible = True
    Picture1.Refresh
End If
 

End Sub
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
tmrNumber.Enabled = False
End Sub
Private Sub tmrNumber_Timer()
    Const MIN_NUMBER = 1
    Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If

    Cantidad.Text = Format$(m_Number)

    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval
    End If

End Sub
Public Sub Refresh_User_Inventory()
If List1(1).ListIndex <> -1 Then Call list1_Click(1)
'List1(1).Refresh '???
End Sub

Public Sub Refresh_Bank_Inventory()
If List1(0).ListIndex <> -1 Then Call list1_Click(0)
'List1(0).Refresh '???
End Sub
