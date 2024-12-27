VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "CoverAO 1.0"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrearPersonaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   20
      Top             =   1080
      Width           =   5895
   End
   Begin VB.ComboBox lstFamiliar 
      BackColor       =   &H00000000&
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
      Height          =   285
      ItemData        =   "frmCrearPersonaje.frx":15F94E
      Left            =   8880
      List            =   "frmCrearPersonaje.frx":15F950
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox txtFamiliar 
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
      Height          =   255
      Left            =   8760
      MaxLength       =   30
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10920
      ScaleHeight     =   1185
      ScaleWidth      =   840
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox lstRaza 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":15F952
      Left            =   840
      List            =   "frmCrearPersonaje.frx":15F954
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3795
      Width           =   2100
   End
   Begin VB.ComboBox lstGenero 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":15F956
      Left            =   840
      List            =   "frmCrearPersonaje.frx":15F958
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3120
      Width           =   2145
   End
   Begin VB.ComboBox lstProfesion 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":15F95A
      Left            =   840
      List            =   "frmCrearPersonaje.frx":15F95C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2460
      Width           =   2100
   End
   Begin VB.PictureBox HeadView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1725
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   19
      Top             =   4560
      Width           =   375
   End
   Begin VB.ComboBox lstHogar 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":15F95E
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":15F96B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Image imgCabeza 
      Height          =   495
      Index           =   1
      Left            =   2160
      Tag             =   "0"
      Top             =   4560
      Width           =   345
   End
   Begin VB.Label lblFamiInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descropcion del familiar"
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
      Height          =   555
      Left            =   8550
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image imgNoDisp 
      Height          =   3465
      Left            =   5880
      Top             =   9360
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   6000
      TabIndex        =   18
      Top             =   9840
      Width           =   2835
   End
   Begin VB.Image imgClase 
      Height          =   3570
      Left            =   8520
      Top             =   4200
      Width           =   2715
   End
   Begin VB.Image imgCabeza 
      Height          =   510
      Index           =   0
      Left            =   1320
      Tag             =   "0"
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      Top             =   5685
      Width           =   240
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   5685
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   6030
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   6405
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      Top             =   6765
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Index           =   4
      Left            =   2160
      TabIndex        =   14
      Top             =   7110
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   6030
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Left            =   2400
      TabIndex        =   15
      Top             =   6405
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Index           =   3
      Left            =   2400
      TabIndex        =   16
      Top             =   6780
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Index           =   4
      Left            =   2400
      TabIndex        =   17
      Top             =   7110
      Width           =   240
   End
   Begin VB.Image boton 
      Height          =   645
      Index           =   0
      Left            =   9720
      Tag             =   "0"
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Image boton 
      Height          =   645
      Index           =   1
      Left            =   720
      Tag             =   "0"
      Top             =   8160
      Width           =   1575
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Public intHeadInd As Integer
Function CheckData() As Boolean

    CheckData = False


    If Not frmCharList.CheckDataCrearPJ(0, False) Then
        lblInfo.Caption = Locale_Error(10) 'No puedes crear mas personajes, has llegado a tu límite de diez personajes.
        Exit Function
    End If
    
    If Cuenta.UserName = vbNullString Then
        lblInfo.Caption = Locale_GUI_Frase(177) 'Seleccione el nombre del personaje
        Exit Function
    End If
    
    If Not AsciiValidos(Cuenta.UserName) Then
        lblInfo.Caption = Locale_GUI_Frase(251) 'Nombre invalido
        Exit Function
    End If
    
    If Len(Cuenta.UserName) < 2 Then
        lblInfo.Caption = Locale_Error(34) 'Corto
        Exit Function
    End If
        
    If Len(Cuenta.UserName) > 30 Then
        lblInfo.Caption = Locale_GUI_Frase(178) 'Largo
        Exit Function
    End If
    
    If Trim(Cuenta.UserName) = "" Then
        Cuenta.UserName = RTrim$(Cuenta.UserName)
        lblInfo.Caption = Locale_GUI_Frase(251) ' Nombre invalido.
        Exit Function
    End If
    
    If UserRaza <= 0 Or UserRaza > NUMRAZAS Then
        lblInfo.Caption = Locale_Error(75)
        Exit Function
    End If
    
    If UserSexo < eGenero.Hombre Or UserSexo > eGenero.Mujer Then
        lblInfo.Caption = Locale_Error(76)
        Exit Function
    End If
    
    If UserClase <= 0 Or UserClase > NUMCLASES Then
        lblInfo.Caption = Locale_Error(77)
        Exit Function
    End If
    
    If UserHogar <= 0 Or UserHogar > NUMCIUDADES Then
        lblInfo.Caption = Locale_Error(78)
        Exit Function
    End If
    
    If Not ValidarCabeza(UserRaza, UserSexo, frmCrearPersonaje.intHeadInd) Then
        lblInfo.Caption = Locale_Error(79)
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        CurrentUser.UserAtributos(i) = val(lbAtt(i - 1).Caption)
        If CurrentUser.UserAtributos(i) < 6 Or CurrentUser.UserAtributos(i) > 18 Then
            lblInfo.Caption = Locale_GUI_Frase(188)
            Exit Function
        End If
    Next i
    
    CheckData = True

End Function

Private Sub Boton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    Call Audio.PlayWave(SND_CLICK)
    Call imgAccionRestaurar

    Select Case Index
    
        Case 0
        
            Cuenta.UserName = Trim$(txtNombre.Text)
            UserRaza = lstRaza.ListIndex
            UserSexo = lstGenero.ListIndex
            UserClase = lstProfesion.ListIndex
            UserHogar = lstHogar.ListIndex
            
            If Not CheckData() Then Exit Sub
            
            If IntervaloPermiteConectar Then
            
                Call FormParser.Parse_Form(Me, E_WAIT)
                
                Call Protocol.ReConnect(E_MODO.CrearNuevoPj)
                
                If frmMain.Socket1.Connected Then
                    boton(0).Picture = General_Load_Picture_From_Resource_Ex("acccreardes")
                    Me.boton(0).Enabled = False
                    Me.boton(1).Enabled = False
                    lblInfo.Caption = "Espere unos instantes..."
                End If
                
            End If

            
        Case 1
        
            Call FormParser.Parse_Form(frmCharList)
            frmCharList.Show
            Unload Me
        
        End Select
        
End Sub
Private Sub boton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
  
   Case 0 'Crear
      boton(0).Picture = General_Load_Picture_From_Resource_Ex("_31")
      boton(0).Tag = "0"
      
   Case 1 'Volver
      boton(1).Picture = General_Load_Picture_From_Resource_Ex("_32")
      boton(1).Tag = "0"
 End Select

End Sub

Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
   Case 0
        If boton(0).Tag = "1" Then
            boton(0).Picture = General_Load_Picture_From_Resource_Ex("_34")
            boton(0).Tag = "0"
        End If
   Case 1
        If boton(1).Tag = "1" Then
            boton(1).Picture = General_Load_Picture_From_Resource_Ex("_33")
            boton(1).Tag = "0"
        End If
 End Select
 
Call imgAccionRestaurar(Index)

End Sub

Private Sub imgCabezaRestaurar(Optional ByVal NoIndex As Integer = 1000, Optional ByVal Over As Boolean = False)

Dim i As Integer

For i = 0 To 1
    If i <> NoIndex Then
        imgCabeza(i).Picture = Nothing
        imgCabeza(i).Tag = "1"
    ElseIf Over Then
        If i = 0 Then
            imgCabeza(0).Picture = General_Load_Picture_From_Resource_Ex("_35")
            imgCabeza(0).Tag = "0"
        Else
            imgCabeza(1).Picture = General_Load_Picture_From_Resource_Ex("_37")
            imgCabeza(1).Tag = "0"
        End If
    End If
Next i

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


Dim i As Integer

For i = 0 To boton.UBound
    If boton(i).Tag = "0" Then
        boton(i).Picture = Nothing
        boton(i).Tag = "1"
    End If
Next i

For i = 0 To imgCabeza.UBound
    If imgCabeza(i).Tag = "0" Then
        imgCabeza(i).Picture = Nothing
        imgCabeza(i).Tag = "1"
    End If
Next i


End Sub
Private Sub Form_Load()

 
Me.Caption = Form_Caption
'Me.Picture = General_Load_Picture_From_Resource_Ex("_30")

Dim i As Integer

lstProfesion.Clear
lstProfesion.AddItem vbNullString
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstGenero.Clear
lstGenero.AddItem vbNullString
lstGenero.AddItem Locale_GUI_Frase(229)
lstGenero.AddItem Locale_GUI_Frase(230)

lstRaza.Clear
lstRaza.AddItem vbNullString
For i = LBound(ListaRazas()) To UBound(ListaRazas())
    lstRaza.AddItem ListaRazas(i)
Next i

lstProfesion.ListIndex = 0
lstGenero.ListIndex = 0
lstRaza.ListIndex = 0
lstHogar.ListIndex = 0

imgClase.Picture = General_Load_Picture_From_Resource_Ex(LCase(lstProfesion.Text) & vbNullString)
 
Call FormParser.Parse_Form(Me)

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub
Private Sub imgCabeza_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lstRaza.ListIndex <= 0 Then Exit Sub
If lstGenero.ListIndex <= 0 Then Exit Sub

Call imgCabezaRestaurar(Index, True)
 
Select Case Index
    Case 0 'Izq
        intHeadInd = intHeadInd - 1
    Case 1 'Der
        intHeadInd = intHeadInd + 1
End Select

If lstGenero.ListIndex = eGenero.Hombre Then
    If intHeadInd > Head_Range(lstRaza.ListIndex).mEnd Then intHeadInd = Head_Range(lstRaza.ListIndex).mStart
    If intHeadInd < Head_Range(lstRaza.ListIndex).mStart Then intHeadInd = Head_Range(lstRaza.ListIndex).mEnd
Else
    If intHeadInd > Head_Range(lstRaza.ListIndex).fEnd Then intHeadInd = Head_Range(lstRaza.ListIndex).fStart
    If intHeadInd < Head_Range(lstRaza.ListIndex).fStart Then intHeadInd = Head_Range(lstRaza.ListIndex).fEnd
End If
 
Call DrawGrhtoHdc(HeadView.hDC, HeadData(intHeadInd).Head(3).GrhIndex, 5, 5)
HeadView.Refresh

End Sub
 Private Sub imgCabeza_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lstRaza.ListIndex <= 0 Then Exit Sub
If lstGenero.ListIndex <= 0 Then Exit Sub

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0 'Izq
        imgCabeza(0).Picture = General_Load_Picture_From_Resource_Ex("_36")
        imgCabeza(0).Tag = "0"
        
    Case 1 'Der
        imgCabeza(1).Picture = General_Load_Picture_From_Resource_Ex("_38")
        imgCabeza(1).Tag = "0"
End Select

End Sub

Private Sub imgCabeza_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0 'Izq
        If imgCabeza(0).Tag = "1" Then
            imgCabeza(0).Picture = General_Load_Picture_From_Resource_Ex("_35")
            imgCabeza(0).Tag = "0"
        End If
        
    Case 1 'Der
        If imgCabeza(1).Tag = "1" Then
            imgCabeza(1).Picture = General_Load_Picture_From_Resource_Ex("_37")
            imgCabeza(1).Tag = "0"
        End If
End Select

Call imgCabezaRestaurar(Index)

End Sub
 
Private Sub imgAccionRestaurar(Optional ByVal NoIndex As Integer = 1000)

Dim i As Integer

For i = 0 To 1
    If i <> NoIndex Then
        boton(i).Picture = Nothing
        boton(i).Tag = "1"
    End If
Next i

End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call imgAccionRestaurar
Call imgCabezaRestaurar
End Sub
Private Sub lstHogar_Click()

    Select Case lstHogar.Text
    
        Case Locale_GUI_Frase(648) ' Nix (Imperial)
            lblInfo.Caption = Locale_GUI_Frase(190)
    
        Case Locale_GUI_Frase(649) ' Illiandor (Republicano)
            lblInfo.Caption = Locale_GUI_Frase(191)
    End Select
    
End Sub
 

Private Sub lstProfesion_Click()
 
On Error Resume Next
 
imgClase.Picture = General_Load_Picture_From_Resource_Ex(LCase(lstProfesion.Text) & vbNullString)

Select Case lstProfesion.ListIndex
    Case 0 'Nothing
        lbAtt(0).Caption = 6
        lbAtt(1).Caption = 6
        lbAtt(2).Caption = 6
        lbAtt(3).Caption = 6
        lbAtt(4).Caption = 6
        
    Case 2 'Magicas
        lbAtt(0).Caption = 6
        lbAtt(1).Caption = 18
        lbAtt(2).Caption = 18
        lbAtt(3).Caption = 10
        lbAtt(4).Caption = 18

    Case 18, 9, 7, 6, 4, 1 'SemiMagicas
        lbAtt(0).Caption = 14
        lbAtt(1).Caption = 14
        lbAtt(2).Caption = 18
        lbAtt(3).Caption = 6
        lbAtt(4).Caption = 18
    
    Case Else 'No magicas
    
        lbAtt(0).Caption = 18
        lbAtt(1).Caption = 18
        lbAtt(2).Caption = 6
        lbAtt(3).Caption = 10
        lbAtt(4).Caption = 18
        
    End Select

End Sub


 
 
Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
lblInfo.Caption = Locale_GUI_Frase(192)
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
End Sub
Private Sub lstRaza_Click()
        
Dim tmpInt As Integer

If lstRaza.list(lstRaza.ListIndex) = vbNullString Then
    intHeadInd = 0
    frmCrearPersonaje.HeadView.Picture = Nothing

    For tmpInt = 1 To NUMATRIBUTOS
        lbBonificador(tmpInt - 1).Visible = False
    Next tmpInt
    
    Exit Sub
End If

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    tmpInt = BonificadorRaza(i, lstRaza.ListIndex)
    
    lbBonificador(i - 1).Caption = IIf(tmpInt > 0, "+" & CStr(tmpInt), CStr(tmpInt))
    If val(lbBonificador(i - 1)) = 0 Then
        lbBonificador(i - 1).Visible = False
    Else
        lbBonificador(i - 1).Visible = True
    End If
Next i

If LenB(lstGenero.list(lstGenero.ListIndex)) > 0 Then
    
    If lstGenero.ListIndex = eGenero.Hombre Then
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).mStart, Head_Range(lstRaza.ListIndex).mEnd))
    Else
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).fStart, Head_Range(lstRaza.ListIndex).fEnd))
    End If
 
    Call DrawGrhtoHdc(HeadView.hDC, HeadData(intHeadInd).Head(3).GrhIndex, 5, 5)
    HeadView.Refresh
End If
 
End Sub


Private Sub lstGenero_Click()

If LenB(lstGenero.list(lstGenero.ListIndex)) = 0 Then
    intHeadInd = 0
    frmCrearPersonaje.HeadView.Picture = Nothing
    Exit Sub
End If

If LenB(lstRaza.list(lstRaza.ListIndex)) > 0 Then
  
    If lstGenero.ListIndex = eGenero.Hombre Then
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).mStart, Head_Range(lstRaza.ListIndex).mEnd))
    Else
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).fStart, Head_Range(lstRaza.ListIndex).fEnd))
    End If
    
    Call DrawGrhtoHdc(HeadView.hDC, HeadData(intHeadInd).Head(3).GrhIndex, 5, 5)
    HeadView.Refresh
    
End If

End Sub

Public Function BonificadorRaza(ByVal Atributo As Integer, ByVal Raza As Byte) As Integer

Select Case Atributo
    Case Atributos.Fuerza
        If Raza = Humano Then BonificadorRaza = 1
        If Raza = ELFOOSCURO Then BonificadorRaza = 2
        If Raza = enano Then BonificadorRaza = 3
        If Raza = Elfo Then BonificadorRaza = 0
        If Raza = Orco Then BonificadorRaza = 5
        If Raza = gnomo Then BonificadorRaza = -5
    Case Atributos.Agilidad
        If Raza = Humano Then BonificadorRaza = 1
        If Raza = ELFOOSCURO Then BonificadorRaza = 0
        If Raza = enano Then BonificadorRaza = -1
        If Raza = Elfo Then BonificadorRaza = 2
        If Raza = Orco Then BonificadorRaza = -2
        If Raza = gnomo Then BonificadorRaza = 3
    Case Atributos.Inteligencia
        If Raza = Humano Then BonificadorRaza = 1
        If Raza = ELFOOSCURO Then BonificadorRaza = 2
        If Raza = enano Then BonificadorRaza = -5
        If Raza = Elfo Then BonificadorRaza = 3
        If Raza = Orco Then BonificadorRaza = -5
        If Raza = gnomo Then BonificadorRaza = 4
    Case Atributos.Carisma
        If Raza = Humano Then BonificadorRaza = 0
        If Raza = ELFOOSCURO Then BonificadorRaza = -1
        If Raza = enano Then BonificadorRaza = -1
        If Raza = Elfo Then BonificadorRaza = 2
        If Raza = Orco Then BonificadorRaza = -4
        If Raza = gnomo Then BonificadorRaza = 0
    Case Atributos.Constitucion
        If Raza = Humano Then BonificadorRaza = 2
        If Raza = ELFOOSCURO Then BonificadorRaza = 1
        If Raza = enano Then BonificadorRaza = 4
        If Raza = Elfo Then BonificadorRaza = 0
        If Raza = Orco Then BonificadorRaza = 4
        If Raza = gnomo Then BonificadorRaza = -1
End Select

End Function

Private Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserSexo As Byte, ByVal Head As Integer) As Boolean
        
    Select Case UserSexo

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= 1 And Head <= 30)
                    
                Case eRaza.enano
                    ValidarCabeza = (Head >= 301 And Head <= 315)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= 101 And Head <= 121)

                Case eRaza.ELFOOSCURO
                    ValidarCabeza = (Head >= 202 And Head <= 212)
 
                Case eRaza.gnomo
                    ValidarCabeza = (Head >= 401 And Head <= 409)
                    
                Case eRaza.Orco
                    ValidarCabeza = (Head >= 501 And Head <= 514)

            End Select
    
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= 70 And Head <= 80)

                Case eRaza.enano
                    ValidarCabeza = (Head >= 370 And Head <= 373)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= 170 And Head <= 189)

                Case eRaza.ELFOOSCURO
                    ValidarCabeza = (Head >= 270 And Head <= 278)
 
                Case eRaza.gnomo
                    ValidarCabeza = (Head >= 470 And Head <= 481)
                
                Case eRaza.Orco
                    ValidarCabeza = (Head >= 570 And Head <= 573)

            End Select

    End Select
        
End Function

