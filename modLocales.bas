Attribute VB_Name = "modLocales"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Type localeNpc
    Name As String
    status As String
    Desc As String
    Hostil As Byte
    MinHP As Long
    MaxHP As Long
End Type

Private Type tSpellLocale
    strName As String
    strDesc As String
    strHechizeroMsg As String
    strTargetMsg As String
    strOwnMsg As String
    strPalabrasMagicas As String
    strTarget As String
    SkillRequerido As Integer
    ManaRequerido As Integer
    StaRequerido As Integer
End Type

Private Type localeObj
    Name As String
    Desc As String
    tipe As Byte
    GrhIndex As Integer
    MinDef As Integer
    MaxDef As Integer
    MinHit As Integer
    MaxHit As Integer
    CreaLuz As String
    RangoLuz As Byte
    Snd1 As Integer 'Snd Equipar
    Snd2 As Integer 'Snd Golpe
    Snd3 As Integer 'Snd fallas
    Nivel As Byte
End Type

Private Type localeSMG
    Mensaje As String
    fontIndex As Byte
    Extra As Byte
End Type

Private objs() As localeObj
Private npcs() As localeNpc
Private arrLocale_SMG() As localeSMG
Private arrLocale_GUI() As String
Private arrLocale_Soporte() As String
Private arrLocale_FACC() As String
Private arrLocale_Error() As String
Private arrLocale_CMD() As String
Private arrLocale_SPL() As tSpellLocale
Private arrLocale_CONFIRM() As String


Public MapNames(1 To 863) As String
Public MapTable(1 To 30, 1 To 23) As Integer
      
Public Function Locale_Error(ByVal btInd As Byte) As String

On Error GoTo ErrorHandler

If btInd = 0 Or btInd > UBound(arrLocale_Error()) Then
    Locale_Error = "Error con el indice Locale_Error nro " & btInd
    Exit Function
End If

Locale_Error = arrLocale_Error(btInd)

Exit Function

ErrorHandler:

End Function
Public Function Locale_Confirm_Frase(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

If intInd > UBound(arrLocale_CONFIRM()) Then
    Locale_Confirm_Frase = "Error con el indice arrLocale_CONFIRM nro " & intInd
    Exit Function
End If

Locale_Confirm_Frase = arrLocale_CONFIRM(intInd)

Exit Function

ErrorHandler:

End Function
Public Function Locale_Soporte_Frase(ByVal intInd As Integer, Optional ByVal Modo As Integer = 0) As String

On Error GoTo ErrorHandler

If intInd > UBound(arrLocale_Soporte()) Then
    Locale_Soporte_Frase = "Error con el indice Locale_Soporte_Frase nro " & intInd
    Exit Function
End If

If Modo = 1 Then

    Dim i As Integer
    
    For i = 1 To UBound(arrLocale_Soporte)
        frmHlp.txtComandos.Text = frmHlp.txtComandos.Text & Locale_Soporte_Frase(i) & vbCrLf
    Next i
    
Else
    Locale_Soporte_Frase = arrLocale_Soporte(intInd)
End If
Exit Function

ErrorHandler:

End Function

Public Function Locale_GUI_Frase(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

If intInd > UBound(arrLocale_GUI()) Then
    Locale_GUI_Frase = "Error con el indice Locale_GUI_FRASE nro " & intInd
    Exit Function
End If


Locale_GUI_Frase = arrLocale_GUI(intInd)

Exit Function

ErrorHandler:

End Function
Public Function Locale_Facc_Frase(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

If intInd > UBound(arrLocale_FACC()) Then
    Locale_Facc_Frase = "Error con el indice Locale_Facc_Frase nro " & intInd
    Exit Function
End If

Locale_Facc_Frase = arrLocale_FACC(intInd)

Exit Function

ErrorHandler:

End Function

Public Function Locale_Parse_Pregunta(ByVal bytHeader As Integer, Optional ByVal strextra1 As String = vbNullString) As String

    On Error GoTo ErrorHandler

    Dim strLocale As String
    Dim lngPos    As Long
    Dim Indice As Integer

    strLocale = Locale_Confirm_Frase(bytHeader)
    
    If LenB(strextra1) = 0 Then
        Locale_Parse_Pregunta = strLocale
        Exit Function
    End If
    
    lngPos = InStr(1, strLocale, "%N")
    
    If lngPos > 0 Then
        strLocale = Replace$(strLocale, "%N", strextra1)
    End If
    
ErrorHandler:
    Locale_Parse_Pregunta = strLocale

End Function

Public Function Locale_Parse_ServidorMensaje(ByVal bytHeader As Integer, Optional ByVal strextra1 As String = vbNullString) As String
    
    
    On Error GoTo ErrorHandler
 
    
    Dim strLocale As String, Palabras() As String, PalabrasGuardadas() As String
    Dim lngPos    As Long
    Dim Indice As Integer

    strLocale = General_Locale_SMG(bytHeader, 0) 'Leemos el indice de la SMG
    
    If LenB(strextra1) = 0 Then
        Locale_Parse_ServidorMensaje = strLocale
        Exit Function
    End If

    Palabras() = Split(strextra1, "%") 'Dividimos las palabras
    
    ReDim PalabrasGuardadas(0 To UBound(Palabras))
    
    For Indice = 0 To UBound(Palabras)
        PalabrasGuardadas(Indice) = Palabras(Indice) 'Guardamos las palabras
        
        If mid(PalabrasGuardadas(Indice), 1, 1) = "*" Then 'Es un index de skills
            PalabrasGuardadas(Indice) = SkillsNames(CInt(mid(PalabrasGuardadas(Indice), 2)))
            
        ElseIf mid(PalabrasGuardadas(Indice), 1, 1) = "$" Then 'Es un index de NPCs
            PalabrasGuardadas(Indice) = General_Locale_NPCs(CInt(mid(PalabrasGuardadas(Indice), 2)), 0)
                
        ElseIf mid(PalabrasGuardadas(Indice), 1, 1) = "~" Then 'Es nombre de un index ej charindex
            PalabrasGuardadas(Indice) = CStr(charlist(CInt(mid(PalabrasGuardadas(Indice), 2))).Nombre)
             
        ElseIf mid(PalabrasGuardadas(Indice), 1, 1) = "^" Then 'Es un index de Spell
            PalabrasGuardadas(Indice) = General_Locale_Spells(CInt(mid(PalabrasGuardadas(Indice), 2)), 0)
            
        ElseIf mid(PalabrasGuardadas(Indice), 1, 1) = "¬" Then 'Es un index de obj, obtenemos nombre
            PalabrasGuardadas(Indice) = General_Locale_Obj(CInt(mid(PalabrasGuardadas(Indice), 0)), 0)
            
        ElseIf mid(PalabrasGuardadas(Indice), 1, 1) = "=" Then 'Es un index de spell PROPIO
            Locale_Parse_ServidorMensaje = General_Locale_Spells(CInt(mid(PalabrasGuardadas(Indice), 2)), 4)
            Exit Function
        End If
        
    Next Indice
    
    lngPos = InStr(1, strLocale, "#1")
    
    If lngPos > 0 Then
        strLocale = Replace$(strLocale, "#1", PalabrasGuardadas(0))
    End If

    lngPos = InStr(1, strLocale, "#2")
    
    If lngPos > 0 Then
        strLocale = Replace$(strLocale, "#2", PalabrasGuardadas(1))
    End If
 
     lngPos = InStr(1, strLocale, "#3")
    
    If lngPos > 0 Then
        strLocale = Replace$(strLocale, "#3", PalabrasGuardadas(2))
    End If
 
     lngPos = InStr(1, strLocale, "#4")
    
    If lngPos > 0 Then
        strLocale = Replace$(strLocale, "#4", PalabrasGuardadas(3))
    End If
    
ErrorHandler:
    Locale_Parse_ServidorMensaje = strLocale

End Function

Public Function Locale_CMD_Get(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler
 

Locale_CMD_Get = arrLocale_CMD(intInd)

Exit Function

ErrorHandler:

End Function

Public Function Locale_Parse_GUI(ByVal strParse As String) As String

On Error GoTo ErrorHandler

Dim lngPosFirst As Long
Dim strTemp As String

lngPosFirst = InStr(1, strParse, "$")

If lngPosFirst <= 0 Then
    Locale_Parse_GUI = strParse
    Exit Function
End If

If InStr(1, strParse, " ") Then

Else
    Locale_Parse_GUI = arrLocale_GUI(val(mid$(strParse, lngPosFirst + 1)))
End If

Exit Function

ErrorHandler:
    Locale_Parse_GUI = strParse

End Function
 
Public Function LoadLocales() As Boolean

    On Error GoTo ErrorHandler
    LoadLocales = False
    Dim f As Integer
    Dim i As Long
    Dim TmpStr As String
    
    If Extract_File(Scripts, App.Path & "\Recursos", "locale_cmd_es.ind", Resource_Path) Then
    
        ReDim arrLocale_CMD(1 To General_Get_Line_Count(Resource_Path & "locale_cmd_es.ind")) As String
        
        f = FreeFile
        Open Resource_Path & "locale_cmd_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, arrLocale_CMD(i)
        Loop
        Close #f
        Delete_File Resource_Path & "locale_cmd_es.ind"
    Else
        Exit Function
    End If
 
 
       If Extract_File(Scripts, App.Path & "\Recursos", "locale_smg_es.ind", Resource_Path) Then
    
        ReDim arrLocale_SMG(1 To General_Get_Line_Count(Resource_Path & "locale_smg_es.ind")) As localeSMG
        
        f = FreeFile
        Open Resource_Path & "locale_smg_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, TmpStr
            arrLocale_SMG(i).Mensaje = ReadField(1, TmpStr, Asc("|"))
            arrLocale_SMG(i).fontIndex = val(ReadField(2, TmpStr, Asc("|")))
            arrLocale_SMG(i).Extra = val(ReadField(3, TmpStr, Asc("|")))
        Loop
 

        Close #f
        Delete_File Resource_Path & "locale_smg_es.ind"
    Else
        Exit Function
    End If
    
    
       If Extract_File(Scripts, App.Path & "\Recursos", "locale_error_es.ind", Resource_Path) Then
    
        ReDim arrLocale_Error(1 To General_Get_Line_Count(Resource_Path & "locale_error_es.ind")) As String
        
        f = FreeFile
        Open Resource_Path & "locale_error_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, arrLocale_Error(i)
        Loop
        Close #f
        Delete_File Resource_Path & "locale_error_es.ind"
    Else
        Exit Function
    End If
    
 
 
      If Extract_File(Scripts, App.Path & "\Recursos", "locale_facc_es.ind", Resource_Path) Then
    
        ReDim arrLocale_FACC(1 To General_Get_Line_Count(Resource_Path & "locale_facc_es.ind")) As String
        
        f = FreeFile
        Open Resource_Path & "locale_facc_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, arrLocale_FACC(i)
        Loop
        Close #f
        Delete_File Resource_Path & "locale_facc_es.ind"
    Else
        Exit Function
    End If
    
 
 
     If Extract_File(Scripts, App.Path & "\Recursos", "locale_obj_es.ind", Resource_Path) Then
    
        ReDim objs(1 To General_Get_Line_Count(Resource_Path & "locale_obj_es.ind")) As localeObj
        
        f = FreeFile
        Open Resource_Path & "locale_obj_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, TmpStr
            objs(i).Name = ReadField(1, TmpStr, Asc("|"))
            objs(i).Desc = ReadField(2, TmpStr, Asc("|"))
            objs(i).GrhIndex = val(ReadField(3, TmpStr, Asc("|")))
            objs(i).tipe = val(ReadField(4, TmpStr, Asc("|")))
            objs(i).MaxDef = val(ReadField(5, TmpStr, Asc("|")))
            objs(i).MinDef = val(ReadField(6, TmpStr, Asc("|")))
            objs(i).MaxHit = val(ReadField(7, TmpStr, Asc("|")))
            objs(i).MinHit = val(ReadField(8, TmpStr, Asc("|")))
            objs(i).CreaLuz = CStr(ReadField(9, TmpStr, Asc("|")))
            objs(i).RangoLuz = val(ReadField(10, TmpStr, Asc("|")))
            objs(i).Snd1 = val(ReadField(11, TmpStr, Asc("|")))
            objs(i).Snd2 = val(ReadField(12, TmpStr, Asc("|")))
            objs(i).Snd3 = val(ReadField(13, TmpStr, Asc("|")))
            objs(i).Nivel = val(ReadField(14, TmpStr, Asc("|")))
        Loop
        
        Close #f
        Delete_File Resource_Path & "locale_obj_es.ind"
    Else
        Exit Function
    End If
    

    If Extract_File(Scripts, App.Path & "\Recursos", "locale_gui_es.ind", Resource_Path) Then
    
        ReDim arrLocale_GUI(1 To General_Get_Line_Count(Resource_Path & "locale_gui_es.ind")) As String
            
        f = FreeFile
        Open Resource_Path & "locale_gui_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, arrLocale_GUI(i)
        Loop
        
        Close #f
        Delete_File Resource_Path & "locale_gui_es.ind"
    Else
        Exit Function
    End If

 
 
    If Extract_File(Scripts, App.Path & "\Recursos", "locale_ayuda_es.ind", Resource_Path) Then
    
        ReDim arrLocale_Soporte(1 To General_Get_Line_Count(Resource_Path & "locale_ayuda_es.ind")) As String
            
        f = FreeFile
        Open Resource_Path & "locale_ayuda_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, arrLocale_Soporte(i)
        Loop
        
        Close #f
        Delete_File Resource_Path & "locale_ayuda_es.ind"
    Else
        Exit Function
    End If

    If Extract_File(Scripts, App.Path & "\Recursos", "locale_confirm_es.ind", Resource_Path) Then
    
        ReDim arrLocale_CONFIRM(1 To General_Get_Line_Count(Resource_Path & "locale_confirm_es.ind")) As String
            
        f = FreeFile
        Open Resource_Path & "locale_confirm_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, arrLocale_CONFIRM(i)
        Loop
        
        Close #f
        Delete_File Resource_Path & "locale_confirm_es.ind"
    Else
        Exit Function
    End If

 
 
     If Extract_File(Scripts, App.Path & "\Recursos", "locale_npc_es.ind", Resource_Path) Then
    
        ReDim npcs(1 To General_Get_Line_Count(Resource_Path & "locale_npc_es.ind")) As localeNpc
        
        f = FreeFile
        Open Resource_Path & "locale_npc_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, TmpStr
            npcs(i).Name = ReadField(1, TmpStr, Asc("|"))
            npcs(i).status = ReadField(2, TmpStr, Asc("|"))
            npcs(i).Desc = ReadField(3, TmpStr, Asc("|"))
            npcs(i).Hostil = ReadField(4, TmpStr, Asc("|"))
            npcs(i).MinHP = ReadField(5, TmpStr, Asc("|"))
            npcs(i).MaxHP = ReadField(6, TmpStr, Asc("|"))
        Loop
        
        Close #f
        Delete_File Resource_Path & "locale_npc_es.ind"
    Else
        Exit Function
    End If
    
    
    
     If Extract_File(Scripts, App.Path & "\Recursos", "locale_spl_es.ind", Resource_Path) Then
    
        ReDim arrLocale_SPL(1 To General_Get_Line_Count(Resource_Path & "locale_spl_es.ind")) As tSpellLocale
        
        f = FreeFile
        Open Resource_Path & "locale_spl_es.ind" For Input As #f
        
        i = 0
        
        Do While Not EOF(f)
            i = i + 1
        Line Input #f, TmpStr
            
        arrLocale_SPL(i).strName = General_Field_Read(1, TmpStr, "|")
        arrLocale_SPL(i).strDesc = General_Field_Read(2, TmpStr, "|")
        arrLocale_SPL(i).strHechizeroMsg = General_Field_Read(3, TmpStr, "|")
        arrLocale_SPL(i).strTargetMsg = General_Field_Read(4, TmpStr, "|")
        arrLocale_SPL(i).strOwnMsg = General_Field_Read(5, TmpStr, "|")
        arrLocale_SPL(i).strPalabrasMagicas = General_Field_Read(6, TmpStr, "|")
        arrLocale_SPL(i).strTarget = General_Field_Read(7, TmpStr, "|")
        arrLocale_SPL(i).ManaRequerido = General_Field_Read(8, TmpStr, "|")
        arrLocale_SPL(i).StaRequerido = General_Field_Read(9, TmpStr, "|")
        arrLocale_SPL(i).SkillRequerido = General_Field_Read(10, TmpStr, "|")
        
        Loop
        
        Close #f
        
        Delete_File Resource_Path & "locale_spl_es.ind"
    Else
        Exit Function
    End If
    
    
    
     If Extract_File(Scripts, App.Path & "\Recursos", "table.ind", Resource_Path) Then
    
        f = FreeFile
        Open Resource_Path & "table.ind" For Binary As #f
            Get #f, , MapTable
        Close #f
        Delete_File Resource_Path & "table.ind"
    Else
        Exit Function
    End If
    
        
     If Extract_File(Scripts, App.Path & "\Recursos", "mapa.ini", Resource_Path) Then
    
        f = FreeFile
        Open Resource_Path & "mapa.ini" For Input As #f
            For i = 1 To 863
                Line Input #f, MapNames(i)
                MapNames(i) = RTrim$(MapNames(i))
            Next i
        Close #f
        Delete_File Resource_Path & "mapa.ini"
    Else
        Exit Function
    End If
 
    LoadLocales = True
Exit Function

ErrorHandler:
LoadLocales = False
End Function

Public Function General_Locale_SMG(ByVal num As Integer, ByVal Tipo As Integer) As String

    On Error GoTo ErrorHandler


    Select Case Tipo
    
        Case 0 'Mensaje
    
            If num = 0 Or num > UBound(arrLocale_SMG()) Then
                General_Locale_SMG = "Error con el mensaje nro " & num
                Exit Function
            End If
    
            General_Locale_SMG = arrLocale_SMG(num).Mensaje
        
        Case 1 'FontIndex
        
            If num = 0 Or num > UBound(arrLocale_SMG()) Then
                General_Locale_SMG = 12
                Exit Function
            End If
            
            General_Locale_SMG = arrLocale_SMG(num).fontIndex
        
        Case 2 'Extra
    
            If num = 0 Or num > UBound(arrLocale_SMG()) Then
                General_Locale_SMG = 12
                Exit Function
            End If
            
            General_Locale_SMG = arrLocale_SMG(num).Extra
        
    End Select
    
    Exit Function
    
ErrorHandler:
    Call RegistrarError(Err.Number, Err.Description, "modLocales.General_Locale_SMG", Erl)
    Resume Next
    
End Function

Public Function General_Locale_Obj(ByVal num As Integer, ByVal Tipo As Integer) As String

    On Error GoTo ErrorHandler

    Select Case Tipo
    
        Case 0 'Nombre
    
            If num > UBound(objs()) Then
                General_Locale_Obj = "Error con el indice General_Locale_Obj nro " & num
                Exit Function
            ElseIf num = 0 Then
                General_Locale_Obj = ""
                Exit Function
            End If
            
            General_Locale_Obj = objs(num).Name
        
        Case 1 'Desc
        
            If num > UBound(objs()) Then
                General_Locale_Obj = "Error con el indice General_Locale_Obj nro " & num
                Exit Function
            ElseIf num = 0 Then
                General_Locale_Obj = ""
                Exit Function
            End If
            
            General_Locale_Obj = objs(num).Desc
        
        Case 2 'Type
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
            
            General_Locale_Obj = objs(num).tipe
        
        Case 3 'GrhIndex
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).GrhIndex
        
        Case 4 'Max objs
        
            General_Locale_Obj = UBound(objs())
        
        Case 5 'Max Def
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).MaxDef
         
        
        Case 6 'MinDef
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).MinDef
     
        Case 7 'Max Hit
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If

            General_Locale_Obj = objs(num).MaxHit
         
        
        Case 8 'MinHit
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).MinHit
        
        
        Case 9 'Luz
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).CreaLuz
        
        
        Case 10 'RangoLuz
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
            
            General_Locale_Obj = objs(num).RangoLuz
        
        
        Case 11 'SN1
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).Snd1
        
        Case 12 'SN2
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).Snd2
        
        
        Case 13 'SN3
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
            
            General_Locale_Obj = objs(num).Snd3
        
        Case 14 'Nivel para item
        
            If num = 0 Or num > UBound(objs()) Then
                General_Locale_Obj = 0
                Exit Function
            End If
        
            General_Locale_Obj = objs(num).Nivel
            
    End Select
    
    
    Exit Function
    
ErrorHandler:
    Call RegistrarError(Err.Number, Err.Description, "modLocales.General_Locale_Obj", Erl)
    Resume Next
    
End Function
Public Function General_Locale_NPCs(ByVal num As Integer, ByVal Tipo As Integer) As String

    On Error GoTo ErrorHandler

    Select Case Tipo
    
        Case 0 'Nombre
    
            If num = 0 Or num > UBound(npcs()) Then
                General_Locale_NPCs = "Error con el indice General_Locale_NPCs nro " & num
                Exit Function
            End If
            
            General_Locale_NPCs = npcs(num).Name
        
        Case 1 'Desc
    
            If num = 0 Or num > UBound(npcs()) Then
                General_Locale_NPCs = "Error con el indice General_Locale_NPCs nro " & num
                Exit Function
            End If
            
            General_Locale_NPCs = npcs(num).Desc
        
        Case 2 'Status
        
            If num = 0 Or num > UBound(npcs()) Then
                General_Locale_NPCs = 0
                Exit Function
            End If
            
            General_Locale_NPCs = npcs(num).status
        
        Case 3 'Max NPCs
        
            General_Locale_NPCs = UBound(npcs())
        
        Case 4 'Hostil
        
            If num = 0 Or num > UBound(npcs()) Then
                General_Locale_NPCs = 0
                Exit Function
            End If
        
            General_Locale_NPCs = npcs(num).Hostil
        
        Case 5 'MinHP
        
            If num = 0 Or num > UBound(npcs()) Then
                General_Locale_NPCs = 0
                Exit Function
            End If
            
            General_Locale_NPCs = npcs(num).MinHP
        
        Case 6 'MaxHP
        
            If num = 0 Or num > UBound(npcs()) Then
                General_Locale_NPCs = 0
                Exit Function
            End If
        
            General_Locale_NPCs = npcs(num).MaxHP
        
    End Select
    
    Exit Function
    
ErrorHandler:
    Call RegistrarError(Err.Number, Err.Description, "modLocales.General_Locale_NPCs", Erl)
    Resume Next
    
End Function

Public Function General_Locale_Spells(ByVal num As Integer, ByVal Tipo As Integer) As String

    On Error GoTo ErrorHandler
 
    Select Case Tipo
    
        Case 0 'Nombre
                
            If num = 0 Then
                General_Locale_Spells = "(" & Locale_GUI_Frase(269) & ")"
                Exit Function
            End If
            
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = "Error con el indice General_Locale_Spells nro " & num
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).strName
        
        Case 1 'Desc
        
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = "Error con el indice General_Locale_Spells nro " & num
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).strDesc
        
        Case 2 'Has curado a ..."
                
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = "Error con el indice General_Locale_Spells nro " & num
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).strHechizeroMsg
        
        Case 3 '""" Te ha curado el envenenamiento
        
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = "Error con el indice General_Locale_Spells nro " & num
                Exit Function
            End If
            
            General_Locale_Spells = arrLocale_SPL(num).strTargetMsg
        
        Case 4 'Te has curado
                    
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = "Error con el indice General_Locale_Spells nro " & num
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).strOwnMsg
        
        
        Case 5 'Palabras magicas
                    
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = "Error con el indice General_Locale_Spells nro " & num
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).strPalabrasMagicas
        
        Case 6 'Target
                
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = 0
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).strTarget
        
        Case 7 'Max hechizos
                
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = 0
                Exit Function
            End If
        
            General_Locale_Spells = UBound(arrLocale_SPL())
    
        Case 8 'Mana req
                
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = 0
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).ManaRequerido
         
        Case 9 'Sta req
                
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = 0
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).StaRequerido
        
        Case 10 'Skill req
                
            If num > UBound(arrLocale_SPL()) Then
                General_Locale_Spells = 0
                Exit Function
            End If
        
            General_Locale_Spells = arrLocale_SPL(num).SkillRequerido
            
    End Select
    
    Exit Function
    
ErrorHandler:
    Call RegistrarError(Err.Number, Err.Description, "modLocales.General_Locale_Spells", Erl)
    Resume Next
    
End Function


