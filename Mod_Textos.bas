Attribute VB_Name = "Mod_Textos"
Option Explicit


Private Type D3DXIMAGE_INFO_A

    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long

End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Sub Text_Font_Initialize()

Dim a As Integer

font_types(1).Font_size = 9
font_types(1).Ascii_code(48) = 21452
font_types(1).Ascii_code(49) = 21453
font_types(1).Ascii_code(50) = 21454
font_types(1).Ascii_code(51) = 21455
font_types(1).Ascii_code(52) = 21456
font_types(1).Ascii_code(53) = 21457
font_types(1).Ascii_code(54) = 21458
font_types(1).Ascii_code(55) = 21459
font_types(1).Ascii_code(56) = 21460
font_types(1).Ascii_code(57) = 21461
For a = 0 To 25
font_types(1).Ascii_code(a + 97) = 21400 + a
Next a

For a = 0 To 25
font_types(1).Ascii_code(a + 65) = 21426 + a
Next a
font_types(1).Ascii_code(33) = 21462
font_types(1).Ascii_code(161) = 21463
font_types(1).Ascii_code(34) = 21464
font_types(1).Ascii_code(36) = 21465
font_types(1).Ascii_code(191) = 21466
font_types(1).Ascii_code(35) = 21467
font_types(1).Ascii_code(36) = 21468
font_types(1).Ascii_code(37) = 21469
font_types(1).Ascii_code(38) = 21470
font_types(1).Ascii_code(47) = 21471
font_types(1).Ascii_code(92) = 21472
font_types(1).Ascii_code(40) = 21473
font_types(1).Ascii_code(41) = 21474
font_types(1).Ascii_code(61) = 21475
font_types(1).Ascii_code(39) = 21476
font_types(1).Ascii_code(123) = 21477
font_types(1).Ascii_code(125) = 21478
font_types(1).Ascii_code(95) = 21479
font_types(1).Ascii_code(45) = 21480
font_types(1).Ascii_code(63) = 21465
font_types(1).Ascii_code(64) = 21481
font_types(1).Ascii_code(94) = 21482
font_types(1).Ascii_code(91) = 21483
font_types(1).Ascii_code(93) = 21484
font_types(1).Ascii_code(60) = 21485
font_types(1).Ascii_code(62) = 21486
font_types(1).Ascii_code(42) = 21487
font_types(1).Ascii_code(43) = 21488
font_types(1).Ascii_code(46) = 21489
font_types(1).Ascii_code(44) = 21490
font_types(1).Ascii_code(58) = 21491
font_types(1).Ascii_code(59) = 21492
font_types(1).Ascii_code(124) = 21493
font_types(1).Ascii_code(252) = 21800
font_types(1).Ascii_code(220) = 21801
font_types(1).Ascii_code(225) = 21802
font_types(1).Ascii_code(233) = 21803
font_types(1).Ascii_code(237) = 21804
font_types(1).Ascii_code(243) = 21805
font_types(1).Ascii_code(250) = 21806
font_types(1).Ascii_code(253) = 21807
font_types(1).Ascii_code(193) = 21808
font_types(1).Ascii_code(201) = 21809
font_types(1).Ascii_code(205) = 21810
font_types(1).Ascii_code(211) = 21811
font_types(1).Ascii_code(218) = 21812
font_types(1).Ascii_code(221) = 21813
font_types(1).Ascii_code(224) = 21814
font_types(1).Ascii_code(232) = 21815
font_types(1).Ascii_code(236) = 21816
font_types(1).Ascii_code(242) = 21817
font_types(1).Ascii_code(249) = 21818
font_types(1).Ascii_code(192) = 21819
font_types(1).Ascii_code(200) = 21820
font_types(1).Ascii_code(204) = 21821
font_types(1).Ascii_code(210) = 21822
font_types(1).Ascii_code(217) = 21823
font_types(1).Ascii_code(241) = 21824
font_types(1).Ascii_code(209) = 21825
font_types(1).Ascii_code(196) = 25238
font_types(1).Ascii_code(194) = 25239
font_types(1).Ascii_code(203) = 25240
font_types(1).Ascii_code(207) = 25241
font_types(1).Ascii_code(214) = 25242
font_types(1).Ascii_code(212) = 25243

font_types(3).Font_size = 9
font_types(3).Ascii_code(97) = 21936
font_types(3).Ascii_code(108) = 21937
font_types(3).Ascii_code(115) = 21938
font_types(3).Ascii_code(70) = 21939
font_types(3).Ascii_code(48) = 21940
font_types(3).Ascii_code(49) = 21941
font_types(3).Ascii_code(50) = 21942
font_types(3).Ascii_code(51) = 21943
font_types(3).Ascii_code(52) = 21944
font_types(3).Ascii_code(53) = 21945
font_types(3).Ascii_code(54) = 21946
font_types(3).Ascii_code(55) = 21947
font_types(3).Ascii_code(56) = 21948
font_types(3).Ascii_code(57) = 21949
font_types(3).Ascii_code(33) = 21950
font_types(3).Ascii_code(161) = 21951
font_types(3).Ascii_code(42) = 21952

End Sub
Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Function Text_Width(Texto As String, ByVal font_index As Integer) As Integer
On Error Resume Next
Dim a As Integer, b As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = font_types(1).Ascii_code(b)
        If (b <> 32) And (b <> 5) And (b <> 129) And (b <> 9) And (b <> 4) And (b <> 255) And (b <> 2) And graf.GrhIndex <> 0 Then
            Text_Width = Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
        Else
            If b = 16 Then
                Text_Width = Text_Width + 16
            Else
                Text_Width = Text_Width + 4
            End If
        End If
    Next a
End Function
Public Sub Char_Dialog_Remove(ByVal charindex As Long)
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 7/10/10
'Delete the dialog chat of the charIndex
'***************************************************
'mermas
With charlist(charindex)
    'Destroit the array string dialog chat
    Erase .dialog
    
    'Set FALSE dialog
    .dl = False
    
    'Set default color , White
    .dialogColor(0) = -1
    .dialogColor(1) = -1
    .dialogColor(2) = -1
    .dialogColor(3) = -1
End With
End Sub
Public Sub Char_Dialog_Remove_All()
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 7/10/10
'Delete the all dialog chat
'***************************************************

'Simple
Dim i As Long
For i = 1 To LastChar
    Char_Dialog_Remove i
Next i
End Sub

Public Function FormatChat(ByRef chat As String) As String()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Formats a dialog into different text lines.
'**************************************************************
    Dim word As String
    Dim curPos As Long
    Dim Length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    Length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(FieldCount(chat, 32)) As String
    
    'Start formating
    Do While acumLength < Length
        word = ReadField(curPos, chat, 32)
        
        wordLength = Len(word)
        
        ' Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            chatLines(curLine) = word
            
            lineLength = wordLength
            acumLength = wordLength
        Else
            ' Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > 18 Then
                'Put it in the next line
                curLine = curLine + 1
                chatLines(curLine) = word
                
                lineLength = wordLength
            Else
                'Add it to this line
                chatLines(curLine) = chatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop
    
    ' If it's only one line, center text
    If curLine = 0 And Length < 18 Then
        chatLines(curLine) = String((18 - Length) \ 2 + 1, " ") & chatLines(curLine)
        chatLines(curLine) = RTrim$(LTrim$(chatLines(curLine)))
    End If
    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines
End Function


Sub Engine_Text_Render(Texto As String, X As Integer, Y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

If font_index = 5 Then
    If NickModerno Then
    Call Engine_Render_Text(SpriteBatch, cfonts, Texto, X, Y, text_color)
    Exit Sub
    End If
End If

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer
Dim graf As Grh

Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD
temp_array(0) = text_color(0)
temp_array(1) = text_color(1)
temp_array(2) = text_color(2)
temp_array(3) = text_color(3)

If (Len(Texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = font_types(font_index).Ascii_code(b)
        If b <> 32 Then
            If graf.GrhIndex <> 0 Then
                'mega sombra O-matica
                graf.GrhIndex = font_types(font_index).Ascii_code(b) + 100
                Grh_Render graf, (X + d) + 1, Y + 1, temp_array, False, False, False
                graf.GrhIndex = font_types(font_index).Ascii_code(b)
                Grh_Render graf, (X + d), Y, temp_array, False, False, False
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
            End If
        Else
            d = d + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = font_types(font_index).Ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.GrhIndex > 12 Then
                'mega sombra O-matica
                graf.GrhIndex = font_types(font_index).Ascii_code(b) + 100
                Grh_Render graf, (X + d) + 1, Y + 1 + f * 14, temp_array, False, False, False
                graf.GrhIndex = font_types(font_index).Ascii_code(b)
                Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
            End If
       End If
       e = e + 1
    Next a
End If

End Sub



Function Engine_Text_Width(Texto As String, Optional multi As Boolean = False) As Integer
Dim a As Integer, b As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

If multi = False Then
    For a = 1 To Len(Texto)
        b = Asc(mid$(Texto, a, 1))
        graf.GrhIndex = font_types(1).Ascii_code(b)
        If (b <> 32) And (b <> 5) And (b <> 129) And (b <> 9) And (b <> 4) And (b <> 255) And (b <> 2) And (b <> 151) And (b <> 152) Then
            Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
        Else
            Engine_Text_Width = Engine_Text_Width + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(Texto)
        b = Asc(mid$(Texto, a, 1))
        graf.GrhIndex = font_types(1).Ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.GrhIndex > 12 Then
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                If d > Engine_Text_Width Then Engine_Text_Width = d
            End If
        End If
        e = e + 1
    Next a
End If
End Function

Function Engine_Text_Height(Texto As String, Optional multi As Boolean = False) As Integer
Dim a As Integer, b As Integer, c  As Integer, d  As Integer, e As Integer, f As Integer
  
If multi = False Then
    Engine_Text_Height = 0
Else
    e = 0
    f = 0
    
    For a = 1 To Len(Texto)
        b = Asc(mid$(Texto, a, 1))

        If b = 32 Or b = 13 Then
            If e >= 20 Then
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then
                    d = d + 4
                End If
            End If
        End If
        e = e + 1
    Next a
  
Engine_Text_Height = f * 14
  
End If

End Function
Public Sub Char_Dialog_Create(ByVal charindex As Long, ByRef chat As String, ByVal color As Long, Optional dialogIndex As Byte = 1)
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 6/10/10
'Enter the dialog chat string and color on charindex
'***************************************************
    If charindex <= 0 Then Exit Sub

    With charlist(charindex)
        'Set the string .Dialog with format for aline
        .dialog = FormatChat(chat)
        
        'Set the color of dialog chat
        .dialogColor(0) = color
        .dialogColor(1) = color
        .dialogColor(2) = color
        .dialogColor(3) = color
        
        'Set TRUE dialog
        .dl = True

        .dialogLife = 5000 + (100 * Len(chat))
        .dialogStart = FrameTime
        
        .dialogHeight = 12

        .dialogIndex = dialogIndex
    End With
End Sub


  Public Sub RemoveDialogsNPCArea()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes all dialogs from all chars.
'**************************************************************
Dim PosX As Byte, PosY As Byte
For PosX = charlist(CurrentUser.UserCharIndex).Pos.X - HalfWindowTileWidth To charlist(CurrentUser.UserCharIndex).Pos.X + HalfWindowTileWidth
    For PosY = charlist(CurrentUser.UserCharIndex).Pos.Y - HalfWindowTileHeight To charlist(CurrentUser.UserCharIndex).Pos.Y + HalfWindowTileHeight
        If MapData(PosX, PosY).charindex > 0 Then
        'If MapData(PosX, PosY).CharIndex > 0 Then _
            'If Len(charlist(MapData(PosX, PosY).CharIndex).Nombre) <= 1 Then _
            Call Mod_Textos.Char_Dialog_Remove(MapData(PosX, PosY).CharIndex)
            If charlist(MapData(PosX, PosY).charindex).EsNPC = True Then Call Mod_Textos.Char_Dialog_Remove(MapData(PosX, PosY).charindex)
            End If
    Next PosY
Next PosX
End Sub
 


Sub Engine_Init_FontTextures()
    '*****************************************************************
    'Init the custom font textures
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
    '*****************************************************************
    
    
5    On Error GoTo Engine_Init_FontTextures_Err:
    
1    If Not Extract_File(Scripts, App.Path & "\Recursos", "font2.png", Resource_Path, False) Then
2        Err.Description = "¡No se puede cargar el archivo de recurso!"
3        GoTo Engine_Init_FontTextures_Err
4    End If

7    Dim i       As Long
8    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
9    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
        
        'Set the texture
6        Set cfonts.Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, _
                                                                   App.Path & "\Recursos\font2.png", _
                                                                   D3DX_DEFAULT, _
                                                                   D3DX_DEFAULT, _
                                                                   0, _
                                                                   0, _
                                                                   D3DFMT_UNKNOWN, _
                                                                   D3DPOOL_MANAGED, _
                                                                   D3DX_FILTER_POINT, _
                                                                   D3DX_FILTER_POINT, _
                                                                   &HFF000000, _
                                                                   ByVal 0, _
                                                                   ByVal 0)
        
        'Store the size of the texture
14        cfonts.TextureSize.X = TexInfo.Width
15        cfonts.TextureSize.Y = TexInfo.Height

    
    
16    Delete_File Resource_Path & "font2.png"
    
10    Exit Sub
    
13 Engine_Init_FontTextures_Err:
11    Call RegistrarError(Err.number, Err.Description, "mod_Textos.Engine_Init_FontTextures", Erl)
12    If General_File_Exists(Resource_Path & "font2.png", vbNormal) Then Delete_File Resource_Path & "font2.png"
    
    
End Sub
Sub Engine_Init_FontSettings()
    '*****************************************************************
    'Init the custom font settings
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
    '*****************************************************************
    
    On Error GoTo Engine_Init_FontSettings_err
    
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "font2.dat", Resource_Path, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo Engine_Init_FontSettings_err
    End If
     

    Dim FileNum  As Byte
    Dim LoopChar As Long
    Dim Row      As Single
    Dim u        As Single
    Dim v        As Single
    Dim i As Long
    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
   
        Open App.Path & "\Recursos\Font2.dat" For Binary As #FileNum
            Get #FileNum, , cfonts.HeaderInfo
        Close #FileNum
        
        'Calculate some common values
        cfonts.CharHeight = cfonts.HeaderInfo.CellHeight - 4
        cfonts.RowPitch = cfonts.HeaderInfo.BitmapWidth \ cfonts.HeaderInfo.CellWidth
        cfonts.ColFactor = cfonts.HeaderInfo.CellWidth / cfonts.HeaderInfo.BitmapWidth
        cfonts.RowFactor = cfonts.HeaderInfo.CellHeight / cfonts.HeaderInfo.BitmapHeight
        
        'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
        For LoopChar = 0 To 255
            
            'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
            Row = (LoopChar - cfonts.HeaderInfo.BaseCharOffset) \ cfonts.RowPitch
            u = ((LoopChar - cfonts.HeaderInfo.BaseCharOffset) - (Row * cfonts.RowPitch)) * cfonts.ColFactor
            v = Row * cfonts.RowFactor
    
            'Set the verticies
            With cfonts.HeaderInfo.CharVA(LoopChar)
                .X = 0
                .Y = 0
                .w = cfonts.HeaderInfo.CellWidth
                .h = cfonts.HeaderInfo.CellHeight
                .Tx1 = u
                .Ty1 = v
                .Tx2 = u + cfonts.ColFactor
                .Ty2 = v + cfonts.RowFactor
            End With
            
        Next LoopChar
        
        Delete_File Resource_Path & "font2.dat"
        
        Exit Sub
        
Engine_Init_FontSettings_err:
    Call RegistrarError(Err.number, Err.Description, "mod_Textos.Engine_Init_FontSettings_err", Erl)
    If General_File_Exists(Resource_Path & "font2.dat", vbNormal) Then Delete_File Resource_Path & "font2.dat"
    
    
End Sub


Public Sub Engine_Render_Text(ByRef Batch As clsBatch, _
                                ByRef UseFont As CustomFont, _
                                ByVal Text As String, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByRef color() As Long)
                                
    '*****************************************************************
    'Render text with a custom font
    '*****************************************************************
    
    Dim TempVA As CharVA
    Dim tempStr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim J As Long
    Dim yOffset As Single
    
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
 
    
    'Get the text into arrays (split by vbCrLf)
    tempStr = Split(Text, vbCrLf)

    X = X - CInt(Engine_GetTextWidth(cfonts, Text) * 0.5)
    
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempStr)
        If Len(tempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempStr(i), vbFromUnicode)
        
            'Loop through the characters
            For J = 1 To Len(tempStr(i))

                Call CopyMemory(TempVA, UseFont.HeaderInfo.CharVA(ascii(J - 1)), 24) 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + yOffset
                
                'Set the colors
       
                
                Call Batch.SetAlpha(False)
                    'Set the texture
               Call Batch.SetTexture(UseFont.Texture)
    
                Call Batch.Draw(TempVA.X, TempVA.Y, TempVA.w, TempVA.h, color, TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2)

                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(J - 1))
                
            Next J
            
        End If
    Next i

End Sub

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer
Dim Len_text As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    Len_text = Len(Text)
    
    'Loop through the text
    For i = 1 To Len_text
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function


