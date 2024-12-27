Attribute VB_Name = "modSeguridadClones"
 
'El módulo 'modSeguridadClones' se encarga de limitar la cantidad de personajes
'que puede crear un mismo jugador en un determinado plazo de tiempo.

Option Explicit

Private Const limite_de_personajes_k    As Integer = 10

Private Type jugador_t

    ip_v                                As String
    personajes_creados_v                As Long

End Type

Private jugadores_m()                   As jugador_t

Public Sub seguridad_clones_construir()
    
100    On Error GoTo seguridad_clones_construir_Err
    
102    ReDim jugadores_m(0 To 0)
        
       Exit Sub
       
seguridad_clones_construir_Err:
104     Call RegistrarError(Err.Number, Err.description, "modSeguridadClones.seguridad_clones_construir_Err", Erl)
106     Resume Next
End Sub

Public Sub seguridad_clones_destruir()

100    On Error GoTo seguridad_clones_destruir_Err
    
102    Erase jugadores_m()
       
       Exit Sub
       
seguridad_clones_destruir_Err:
104     Call RegistrarError(Err.Number, Err.description, "modSeguridadClones.seguridad_clones_destruir_Err", Erl)
106     Resume Next
End Sub

Public Function seguridad_clones_validar(ByVal ip_p As String) As Boolean

100  On Error GoTo seguridad_clones_validar_err

102    Dim iterador_v As Long
  
104    ip_p = UCase$(ip_p)
  
106    For iterador_v = LBound(jugadores_m) To UBound(jugadores_m)
  
108        With jugadores_m(iterador_v)
      
110            If .ip_v = ip_p Then
          
112                If .personajes_creados_v >= limite_de_personajes_k Then
              
114                    seguridad_clones_validar = False
116                    Exit Function
                  
118                Else
              
120                    .personajes_creados_v = .personajes_creados_v + 1
                  
122                    seguridad_clones_validar = True
124                    Exit Function
                  
126                End If
          
128            End If
      
        End With
      
    Next
  
130    ReDim Preserve jugadores_m(LBound(jugadores_m) To UBound(jugadores_m) + 1)
  
132    With jugadores_m(UBound(jugadores_m))
  
134        .ip_v = ip_p
136        .personajes_creados_v = 1
  
    End With

138    seguridad_clones_validar = True

       Exit Function
       
seguridad_clones_validar_err:
140     Call RegistrarError(Err.Number, Err.description, "modSeguridadClones.seguridad_clones_validar", Erl)
142     Resume Next
End Function

Public Sub seguridad_clones_limpiar()

100    On Error GoTo seguridad_clones_limpiar_Err


102    Erase jugadores_m()
104    ReDim jugadores_m(0 To 0)

       Exit Sub
       
seguridad_clones_limpiar_Err:
106     Call RegistrarError(Err.Number, Err.description, "modSeguridadClones.seguridad_clones_limpiar", Erl)
108     Resume Next
End Sub

