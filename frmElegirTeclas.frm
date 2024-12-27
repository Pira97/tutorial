VERSION 5.00
Begin VB.Form frmElegirTeclas 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$396"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   Icon            =   "frmElegirTeclas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmElegirTeclas.frx":000C
   ScaleHeight     =   5130
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "$1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   1
      Top             =   4530
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4590
      TabIndex        =   0
      Top             =   4530
      Width           =   4335
   End
End
Attribute VB_Name = "frmElegirTeclas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call CustomKeys.LoadDefaultsOpcion2
Dim i As Long

For i = 1 To CustomKeys.Count
    frmCustomKeys.Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
Call CustomKeys.SaveCustomKeys
 If PrimeraVez = 1 Then
     PrimeraVez = 0
    Call WriteVar(App.Path & "\Init\CovAoInit.ini", "LAUNCHER", "PrimeraVez", PrimeraVez)
    End If

Unload Me
End Sub

Private Sub Command2_Click()
Call CustomKeys.LoadDefaults
Dim i As Long

For i = 1 To CustomKeys.Count
    frmCustomKeys.Text1(i) = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
Call CustomKeys.SaveCustomKeys
 If PrimeraVez = 1 Then
     PrimeraVez = 0
    Call WriteVar(App.Path & "\Init\CovAoInit.ini", "LAUNCHER", "PrimeraVez", PrimeraVez)
    End If
Unload Me
End Sub

Private Sub Form_Load()
 Call Audio.PlayWave(SND_INFO)
 Call FormParser.Parse_Form(Me)

End Sub
