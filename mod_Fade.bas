Attribute VB_Name = "mod_Fade"
 
Option Explicit
 
 Public Function Min(ByVal val1 As Long, ByVal val2 As Long) As Long

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/27/06
    'It's faster than iif and I like it better
    '***************************************************
    If val1 < val2 Then
        Min = val1
    Else
        Min = val2

    End If

End Function
 
Public Sub Copy_RGBAList_WithAlpha(Dest() As Long, Src() As Long, ByVal Alpha As Byte)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
 
    
    Dim i As Long
    
 For i = 0 To 3
         Dest(i) = Src(i)
        Dest(i) = Alpha
      Next i
  
    Exit Sub
 
    
End Sub
 
