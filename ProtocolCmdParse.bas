Attribute VB_Name = "ProtocolCmdParse"
'Argentum Online
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)ArgumentosRaw

'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'

Option Explicit
Private Const MAX_DESC  As Integer = 20
Public Enum eNumber_Types

    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger

End Enum
 
''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, _
                          Optional ByVal red As Integer = 204, _
                          Optional ByVal green As Integer = 193, _
                          Optional ByVal blue As Integer = 115, _
                          Optional ByVal bold As Boolean = False, _
                          Optional ByVal italic As Boolean = False)
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/03/07
    '
    '***************************************************
    Call AddtoRichTextBox(Message, red, green, blue, bold, italic)

End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, _
                            ByVal Tipo As eNumber_Types) As Boolean

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/06/07
    '
    '***************************************************
    Dim Minimo As Long

    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then Exit Function
    
    Select Case Tipo

        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6

    End Select
    
    If val(Numero) >= Minimo And val(Numero) <= Maximo Then ValidNumber = True

End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal IP As String) As Boolean

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/06/07
    '
    '***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(IP, ".")
    
    If UBound(tmpArr) <> 3 Then Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then Exit Function
    
    validipv4str = True

End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Public Function str2ipv4l(ByVal IP As String) As Byte()

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/26/07
    'Last Modified By: Rapsodius
    'Specify Return Type as Array of Bytes
    'Otherwise, the default is a Variant or Array of Variants, that slows down
    'the function
    '***************************************************
    Dim tmpArr() As String

    Dim bArr(3)  As Byte
    
    tmpArr = Split(IP, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr

End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()

    '***************************************************
    'Author: Lucas Tavolaro Ortuz (Tavo)
    'Useful for AEMAIL BUG FIX
    'Last Modification: 07/26/07
    'Last Modified By: Rapsodius
    'Specify Return Type as Array of Strings
    'Otherwise, the default is a Variant or Array of Variants, that slows down
    'the function
    '***************************************************
    Dim tmpArr(0 To 1) As String

    Dim Pos            As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString

    End If
    
    AEMAILSplit = tmpArr

End Function

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

