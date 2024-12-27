Attribute VB_Name = "Statistics"
'**************************************************************
' modStatistics.bas - Takes statistics on the game for later study.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private Type trainningData

    startTick As Long
    trainningTime As Long

End Type

Private Type fragLvlRace

    matrix(1 To 50, 1 To 5) As Long

End Type

Private Type fragLvlLvl

    matrix(1 To 50, 1 To 50) As Long

End Type

Private trainningInfo()                       As trainningData

Private fragLvlRaceData(1 To 7)               As fragLvlRace
Private fragLvlLvlData(1 To 7)                As fragLvlLvl
Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long

'Currency just in case.... chats are way TOO often...
Private keyOcurrencies(255)                   As Currency

Public Sub Initialize()
    ReDim trainningInfo(1 To MaxUsers) As trainningData

End Sub

Public Sub ParseChat(ByRef S As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i   As Long
    Dim key As Integer
    
    For i = 1 To Len(S)
        key = Asc(mid$(S, i, 1))
        
        keyOcurrencies(key) = keyOcurrencies(key) + 1
    Next i
    
    'Add a NULL-terminated to consider that possibility too....
    keyOcurrencies(0) = keyOcurrencies(0) + 1

End Sub
