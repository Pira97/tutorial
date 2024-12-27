VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 7

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg
    Texto As String
    Segundos As Byte

End Type

Private p_Dialogos() As t_GuildDlg
Private p_Activo     As Boolean

Public Property Let CantidadDialogos(ByVal v As Byte)

    If v > 0 Then
        ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
        p_CANTIDADDIALOGOS = v

    End If

End Property
 
''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Retrieves the number of dialogs to be displayed on screen
    '***************************************************
    CantidadDialogos = p_CANTIDADDIALOGOS

End Property

Public Property Let Activo(ByVal v As Boolean)
    p_Activo = v

End Property

Public Property Get Activo() As Boolean
    Activo = p_Activo

End Property

Private Sub Class_Initialize()
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True

End Sub

Public Sub Draw()
    Dim i As Long
    Dim y As Long

    If Not p_Activo Then Exit Sub

    '500 600 700 800 900
    For i = 1 To p_CANTIDADDIALOGOS

        If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
            If Left$(p_Dialogos(i).Texto, 1) = Chr$(3) Then
            Engine_Text_Render mid$(p_Dialogos(i).Texto, 2), 10, y + (i * 10), green
                
            Else
    
                 Engine_Text_Render p_Dialogos(i).Texto, 10, y + (i * 10), White
            End If

        End If

    Next i

End Sub
