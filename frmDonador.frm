VERSION 5.00
Begin VB.Form frmDonador 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Icon            =   "frmDonador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$580"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   120
      MouseIcon       =   "frmDonador.frx":000C
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00484444&
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   75
   End
   Begin VB.Image Image14 
      Height          =   615
      Left            =   3405
      Top             =   1410
      Width           =   735
   End
   Begin VB.Image Image13 
      Height          =   615
      Left            =   3120
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image12 
      Height          =   615
      Left            =   2445
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image11 
      Height          =   615
      Left            =   1650
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image10 
      Height          =   615
      Left            =   720
      Top             =   2040
      Width           =   855
   End
   Begin VB.Image Image9 
      Height          =   615
      Left            =   2640
      Top             =   1395
      Width           =   735
   End
   Begin VB.Image Image8 
      Height          =   615
      Left            =   0
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image7 
      Height          =   615
      Left            =   1920
      Top             =   1395
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   1200
      Top             =   1395
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   15
      Top             =   1380
      Width           =   1200
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   3210
      Top             =   750
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   2355
      Top             =   750
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   15
      Top             =   750
      Width           =   2295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$581"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00484444&
      Height          =   195
      Left            =   2070
      TabIndex        =   0
      Top             =   45
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3075
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmDonador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Sub Form_Load()
 Call Audio.PlayWave(SND_INFO)
 
    Call FormParser.Parse_Form(Me)
 
  Make_Transparent_Form Me.hwnd, 180
 

End Sub
 
Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
Private Sub Image3_Click()
Unload Me
End Sub
Private Sub Image4_Click()
Unload Me
End Sub
Private Sub Image5_Click()
Unload Me
End Sub
Private Sub Image6_Click()
Unload Me
End Sub
Private Sub Image7_Click()
Unload Me
End Sub
Private Sub Image8_Click()
Unload Me
End Sub
Private Sub Image9_Click()
Unload Me
End Sub
Private Sub Image10_Click()
Unload Me
End Sub
Private Sub Image11_Click()
Unload Me
End Sub
Private Sub Image12_Click()
Unload Me
End Sub
Private Sub Image13_Click()
Unload Me
End Sub
Private Sub image1_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = ""
End Sub
Private Sub image2_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Armadura de Placas Roja."
End Sub

 
Private Sub image3_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Armadura de Placas Completas."
End Sub
 


Private Sub image4_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Armadura de Asesino."
End Sub


Private Sub image5_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Armadura de Placas Azules."
End Sub

Private Sub image6_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Túnica de Nigromante."
End Sub

Private Sub image7_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Túnica Dorada."
End Sub

Private Sub image8_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Túnica de Clérigo."
End Sub

Private Sub image9_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Armadura Murex."
End Sub

Private Sub image10_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Escudo de Plata +2."
End Sub

Private Sub image11_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Escudo de Torre +1."
End Sub

Private Sub image12_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Yelmo."
End Sub

Private Sub image13_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Espada de Plata."
End Sub

Private Sub Label3_Click()
Call ShellExecute(0, "Open", "http://www.facebook.com/Link-AO", "", App.Path, SW_SHOWNORMAL)
End Sub
