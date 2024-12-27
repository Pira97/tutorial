VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private InventoryOffset                As Long
Private InvSelectedItem                As Long             'Currently selected item
Private TempItem                       As Integer
Private TempSlot                       As Integer

Private MouseX                         As Integer
Private MouseY                         As Integer
Private DragItem                       As Byte
Private DragActivado As Boolean

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1
Public Sub SelectGold()
 InvSelectedItem = FLAGORO
End Sub
Public Property Get SelectedItem() As Long
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer

    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    MaxDef = Inventory(slot).MaxDef

End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer

    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    MinDef = Inventory(slot).MinDef

End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single

    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    Valor = Inventory(slot).Valor
    
End Property

Public Property Get Amount(ByVal slot As Byte) As Long

    If slot = FLAGORO Then
        Amount = CurrentUser.UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        Amount = Inventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer

    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    OBJIndex = Inventory(slot).OBJIndex

End Property

Public Property Get ObjType(ByVal slot As Byte) As Integer
 
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    ObjType = Inventory(slot).ObjType

End Property

Public Property Get ItemName(ByVal slot As Byte) As String
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    ItemName = Inventory(slot).Name

End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean

    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    Equipped = Inventory(slot).Equipped

End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, ByVal eValor As Single, ByVal Puede As Byte, Optional ByVal Refresh As Boolean = False)
    
    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    
    With Inventory(slot)
    
    If Refresh = True Then
    
        .Amount = 0
        .OBJIndex = 0
        .Equipped = 0
        .Valor = 0
        .PuedeUsar = 0
        
        .MaxDef = 0
        .MinDef = 0
        .GrhIndex = 0
        .MaxHit = 0
        .MinHit = 0
        .Name = ""
        .ObjType = 0
    
    Else
    
        .Amount = eAmount
        .OBJIndex = eOBJIndex
        .Equipped = eEquipped
        .Valor = eValor
        .PuedeUsar = Puede
        
        .MaxDef = CInt(General_Locale_Obj(eOBJIndex, 5))
        .MinDef = CInt(General_Locale_Obj(eOBJIndex, 6))
        .GrhIndex = CInt(General_Locale_Obj(eOBJIndex, 3))
        .MaxHit = CInt(General_Locale_Obj(eOBJIndex, 7))
        .MinHit = CInt(General_Locale_Obj(eOBJIndex, 8))
        .Name = General_Locale_Obj(eOBJIndex, 0)
        .ObjType = CInt(General_Locale_Obj(eOBJIndex, 2))
    
    End If
    
    End With
    
    If CurrentUser.Logged = True Then
    If frmMain.picInv.Visible Then RenderInv = True
    End If
    
    Exit Sub

End Sub


Public Sub SetItemUser(ByVal slot As Byte, ByVal Accion As Integer, ByVal Valor As Integer)

    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
    
        Select Case Accion
        
        Case 1 ' Equipa
            .Equipped = 1
            
        Case 2 'Des Equipa
            .Equipped = 0
            
        Case 3 'Cantidad
        
            .Amount = Valor
            If .Amount = 0 Then .GrhIndex = 0
        
        End Select
    End With
    
    If CurrentUser.Logged = True Then
        If frmMain.picInv.Visible = True Then RenderInv = True
    End If
    
    Exit Sub

End Sub


Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
    
    Dim TempItem As Long
    
    TempItem = (X \ 32) + 1 + (Y \ 32) * 5
    TempSlot = TempItem
    
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS And TempItem >= 1 Then
        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex > 0 Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    RenderInv = True
     
    
End Function

Public Sub DrawInventory()

    Dim i As Integer
    Dim X As Single, Y As Single
    Dim color(3) As Long
    
    For i = 1 To UBound(Inventory)
    
     If Inventory(i).GrhIndex Then
        
        If i <> DragItem Then
    
            X = ((i - 1) Mod 5) * 32
            Y = ((i - 1) \ 5) * 32
    
            If Inventory(i).PuedeUsar Then
                Engine_Long_To_RGB_List color, &HFFFFFFFF
            Else
                Engine_Long_To_RGB_List color, -1763311616
            End If
            
            DrawGrhIndextoSurface Inventory(i).GrhIndex, X - 1, Y - 1, color(0)

            Engine_Long_To_RGB_List color, &HFFFFFFFF
  
            Call Engine_Text_Render(str(Inventory(i).Amount), X - 5, Y + 19, color)
  
            Engine_Long_To_RGB_List color, &HFFFF0000

            X = ((i - 1) Mod 5)
            Y = ((i - 1) \ 5)

            If Inventory(i).Equipped Then Call Engine_Text_Render("+", X * 32 + 22, Y * 32 - 2, color, 1)
            
        End If
      
     End If
     
    Next i
    
    If InvSelectedItem > 0 Then
        Engine_Long_To_RGB_List color, &HFFFFFFFF
        DrawGrhIndextoSurface 2, ((InvSelectedItem - 1) Mod 5) * 32, ((InvSelectedItem - 1) \ 5) * 32, color(0)
    End If
    
     
    If DragItem > 0 And DragActivado Then

    If Inventory(DragItem).PuedeUsar Then
        Engine_Long_To_RGB_List color, &HFFFFFFFF
    Else
        Engine_Long_To_RGB_List color, -1763311616
    End If
    
    RenderInv = True
    
    DrawGrhIndextoSurface Inventory(DragItem).GrhIndex, MouseX - 16, MouseY - 16, color(0)

    Else
    
    RenderInv = False
    
    End If
    
End Sub

Public Sub Initialize(ByRef InvPic As PictureBox)
    
    Set InventoryWindow = InvPic
    
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub
    
Private Sub InventoryWindow_MouseDown(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, Y As Single)


    
    If frmComerciar.Visible = True Or frmBancoObj.Visible Then
    InvSelectedItem = 0
    Exit Sub
    End If
    
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
    
    If Button = vbRightButton Then
   
    'Get the currently clickced item
    DragItem = (X \ 32) + 1 + (Y \ 32) * 5
    
    If DragItem <= MAX_INVENTORY_SLOTS And DragItem >= 1 Then
        If Inventory(DragItem).GrhIndex > 0 Then
            DragActivado = True
            RenderInv = True
            InvSelectedItem = 0
            End If
        End If
    End If
    

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Implements the mouse move event of the inventory picture box
    'Displays a ToolTip of the item under the mouse
    '***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim tStr As String
    
    temp_x = X \ 32 + 1
    temp_y = Y \ 32
    
    TempItem = (temp_x + temp_y * 5)
    
    TempSlot = TempItem
    
    If TempItem = 0 Then Exit Sub
    
    If TempItem > 0 And TempItem <= MAX_INVENTORY_SLOTS And TempItem <> CurrentUser.LastItem Then
        CurrentUser.LastItem = TempItem
        
        Select Case Inventory(TempItem).ObjType
            Case eObjType.otWeapon, eObjType.otFlechas, eObjType.otNudillos
                tStr = Inventory(TempItem).Name & vbCr & Locale_GUI_Frase(175) & ": " & Inventory(TempItem).MinHit & "/" & Inventory(TempItem).MaxHit
            Case eObjType.otArmadura, eObjType.otCASCO, eObjType.otESCUDO
                tStr = Inventory(TempItem).Name & vbCr & Locale_GUI_Frase(176) & ": " & Inventory(TempItem).MinDef & "/" & Inventory(TempItem).MaxDef
            Case eObjType.otMonturas, eObjType.otBarcos
                tStr = Inventory(TempItem).Name & vbCr & Locale_GUI_Frase(175) & ": " & Inventory(TempItem).MinHit & "/" & Inventory(TempItem).MaxHit & vbCrLf & Locale_GUI_Frase(176) & ": " & Inventory(TempItem).MinDef & "/" & Inventory(TempItem).MaxDef
            Case Else
                tStr = Inventory(TempItem).Name
        End Select

    frmMain.lblInvInfo.FontSize = IIf(Len(tStr) > 62, 7, 8)
    frmMain.lblInvInfo.Caption = tStr
    frmMain.picInv.ToolTipText = vbNullString
    End If
 
    MouseX = CInt(X): MouseY = CInt(Y)
 
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If frmComerciar.Visible = True Or frmBancoObj.Visible Then
    InvSelectedItem = 0
    Exit Sub
    End If
 
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
        
    
    'Store previously selected item
    Dim prevSelItem As Long
   
    prevSelItem = InvSelectedItem
    
    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(MouseX), CInt(MouseY))
     
    If Button = vbRightButton Then
        If DragActivado = True Then
            If TempItem <= MAX_INVENTORY_SLOTS And TempItem >= 1 Then
            WriteSwapObjects TempItem, DragItem
            DragActivado = False
            DragItem = 0
            End If
        End If
    End If
    
End Sub

