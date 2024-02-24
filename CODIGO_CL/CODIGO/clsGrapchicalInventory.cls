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
 '*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private ShowInfo As Boolean
Private ShowEquipped As Boolean
Private Inventory_Pictures            As Integer ' Cada inventario comprende un rectangulo de AxB donde conforma X cantidad de TILES. (Comercios se usa 28, Inventario se usa el clásico)

Private InventoryIndex             As Integer

Private InventoryOffset            As Long             'Number of lines we scrolled down from topmost

Private InvSelectedItem            As Long             'Currently selected item

Private ownTilePixelWidth          As Integer

Private ownTilePixelHeight         As Integer

Private startdX                    As Integer

Private startdY                    As Integer

Private ShowValue                  As Boolean

Private ValueFormated              As Boolean

Private ImgContainer               As Boolean

Private slotDragged                As Integer

Private dragAndDropOn              As Boolean

Private PictureModernOn            As Boolean

Private Inventory()                     As Inventory    'User's inventory

Private MouseDownSelectedItem      As Long

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Public uMoveItem                   As Boolean 'USANDO DRAG

Public sMoveItem                   As Boolean 'USABA DRAG
    
Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)

Private lFrameTimer                As Long

Private ViewInfo                   As Boolean

Private FontSize As Byte

Private Sub Class_Initialize()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    '
    '***************************************************
End Sub

Private Sub Class_Terminate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    '
    '***************************************************
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 14/01/2010
    'Scrolls the graphical inventory up / down
    '14/01/2010: ZaMa - Now redraws inventory only if needed
    '***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, Inventory_Pictures, -Inventory_Pictures)
    End If

    DrawInventory
End Sub

Public Function CanScroll(ByVal Up As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 14/01/2010
    'Determines wether inventory can be scrolled up / down
    '***************************************************
    If Up Then
        CanScroll = (InventoryOffset + Inventory_Pictures) <= UBound(Inventory) + (InventoryOffset / Inventory_Pictures)
    Else
        CanScroll = InventoryOffset > 0
    End If
    
End Function

Public Sub SelectGold()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Sets the gold as the currently selected item
    '***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
End Sub

Public Sub SelectEldhir()

    '***************************************************
    'Author: WAICON
    'Last Modify Date:
    'Sets the Dsp as the currently selected item
    '***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the dsps money
    InvSelectedItem = FLAGELDHIR
End Sub

Public Sub DeselectItem()

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 30/11/2009
    'Deselect the currently selected item
    '***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
End Sub

Public Sub ChangeSlotItemAmount(ByVal Slot As Integer, ByVal Amount As Long)
    '***************************************************
    'Author: ZaMa
    'Created Date: 30/11/2009
    'Last Modify Date: 09/12/2009 (Budi)
    'Change the amount of the required item, and update the slot.
    '***************************************************
    
    Inventory(Slot).Amount = Amount

End Sub

Public Property Get SelectedItem() As Long
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item index
    '***************************************************
    SelectedItem = InvSelectedItem
    
End Property

Public Property Get MaxHit(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the max hit of the selected item
    '***************************************************
    MaxHit = Inventory(Slot).MaxHit
End Property

Public Property Get MinHit(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the min hit of the selected item
    '***************************************************
    MinHit = Inventory(Slot).MinHit
End Property

Public Property Get MaxHitMag(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the max hit of the selected item
    '***************************************************
    MaxHitMag = Inventory(Slot).MaxHitMag
End Property

Public Property Get MinHitMag(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the min hit of the selected item
    '***************************************************
    MinHitMag = Inventory(Slot).MinHitMag
End Property

Public Property Get MaxDef(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the defense of the selected item
    '***************************************************
    MaxDef = Inventory(Slot).MaxDef
End Property

Public Property Get MinDef(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Budi
    'Last Modify Date: 02/1209
    'Retrieves the defense of the selected item
    '***************************************************
    MinDef = Inventory(Slot).MinDef
End Property

Public Property Get MaxDefMag(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the defense of the selected item
    '***************************************************
    MaxDefMag = Inventory(Slot).MaxDefMag
End Property

Public Property Get MinDefMag(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Budi
    'Last Modify Date: 02/1209
    'Retrieves the defense of the selected item
    '***************************************************
    MinDefMag = Inventory(Slot).MinDefMag
End Property

Public Property Get GrhIndex(ByVal Slot As Integer) As Long
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the grh index of the selected item
    '***************************************************
    GrhIndex = Inventory(Slot).GrhIndex
End Property

Public Property Get Valor(ByVal Slot As Integer) As Single
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the value of the selected item
    '***************************************************
    Valor = Inventory(Slot).Valor
End Property

Public Property Get ExistSkin(ByVal Slot As Integer) As Single
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the value of the selected item
    '***************************************************
    ExistSkin = Inventory(Slot).ExistSkin
End Property

Public Property Get ValorAzul(ByVal Slot As Integer) As Single
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the value of the selected item
    '***************************************************
    ValorAzul = Inventory(Slot).ValorAzul
End Property

Public Property Get CanUse(ByVal Slot As Integer) As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the value of the selected item
    '***************************************************
    CanUse = Inventory(Slot).CanUse
End Property

Public Property Get Amount(ByVal Slot As Integer) As Long

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's amount
    '***************************************************
    If Slot = FLAGORO Then
        Amount = UserGLD
    ElseIf Slot >= LBound(Inventory) And Slot <= UBound(Inventory) Then
        Amount = Inventory(Slot).Amount
    End If

End Property

Public Property Get ObjIndex(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's object index
    '***************************************************
    ObjIndex = Inventory(Slot).ObjIndex
End Property

Public Property Get ObjType(ByVal Slot As Integer) As Integer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's object type
    '***************************************************
    ObjType = Inventory(Slot).ObjType
End Property

Public Property Get ItemName(ByVal Slot As Integer) As String
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves the selected item's name
    '***************************************************
    ItemName = Inventory(Slot).Name
End Property

Public Property Get Equipped(ByVal Slot As Integer) As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Retrieves True if the item at the given pos is eqiupped
    '***************************************************
    Equipped = Inventory(Slot).Equipped
End Property

Public Property Get MaxObjs() As Integer
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 09/16/09
    'Retrieves the capacity of the Inventory
    '***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 09/16/09
    'Set the capacity of the Inventary
    '***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 11/12
    'Indicates if the amount of the items in the inventory is shown
    '***************************************************
    ShowValue = bShow
End Sub

Public Sub ClearAllSlots()
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Clear All Slots.
    '***************************************************

    Dim Slot As Long

    For Slot = 1 To Me.MaxObjs
        Call SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, False, 0, 0, 0, 0)
    Next Slot

End Sub

Public Sub SetItem(ByVal Slot As Integer, _
                   ByVal eOBJIndex As Integer, _
                   ByVal eAmount As Long, _
                   ByVal eEquipped As Byte, _
                   ByVal eGrhIndex As Long, _
                   ByVal eOBJType As Integer, _
                   ByVal eMaxHit As Integer, _
                   ByVal eMinHit As Integer, _
                   ByVal MaxDef As Integer, _
                   ByVal MinDef As Integer, _
                   ByVal eValor As Single, _
                   ByVal eName As String, _
                   ByVal eValorAzul As Single, _
                   ByVal eCanUse As Boolean, _
                   ByVal eMinHitMag As Integer, _
                   ByVal eMaxHitMag As Integer, _
                   ByVal eMinDefMag As Integer, _
                   ByVal eMaxDefMag As Integer, _
                   Optional ByVal eBronce As Byte = 0, _
                   Optional ByVal ePlata As Byte = 0, _
                   Optional ByVal eOro As Byte = 0, _
                   Optional ByVal ePremium As Byte = 0, _
                   Optional ByVal ExistSkin As Integer = 0)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/04/06
    'Sets all data for a given inventory slot
    '***************************************************
    If Slot < 1 Or Slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(Slot)
        .Amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .ObjIndex = eOBJIndex
        .ObjType = eOBJType
        .Valor = eValor
        .ValorAzul = eValorAzul
        .CanUse = eCanUse
        .MinHitMag = eMinHitMag
        .MaxHitMag = eMaxHitMag
        .MinDefMag = eMinDefMag
        .MaxDefMag = eMaxDefMag
        .Bronce = eBronce
        .Plata = ePlata
        .Oro = eOro
        .Premium = ePremium
        .ExistSkin = ExistSkin
    End With
    
  '  Call Invalidate(InventoryWindow.hWnd)
    
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Selects the item clicked if it's valid and return's it's index
    '***************************************************
    Dim TempItem As Long

    Dim temp_x   As Long

    Dim temp_y   As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    Dim Last As Long
    
    Last = UBound(Inventory) - InventoryOffset
    
    'If Last > 0 Then
        TempItem = temp_x + (temp_y) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 + (UBound(Inventory) - Last)
    'Else
       ' TempItem = temp_x + (temp_y + Last) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
   ' End If
    
    
    
    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then

        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If

End Function

Public Sub DrawInventory()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Renders the inventory to the given PictureBox
    '***************************************************
    Dim LoopC As Long

    Dim Mult  As Byte

    Dim Div   As Byte

    Dim Porc  As Single
    
    Call wGL_Graphic.Use_Device(g_Captions(InventoryIndex))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, InventoryWindow.ScaleWidth, InventoryWindow.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    Dim lastTileRow As Integer

    Dim NotBig      As Boolean

    lastTileRow = InventoryWindow.ScaleWidth / ownTilePixelWidth

    Mult = 1
    Div = 1
    Porc = 0.2
    
    #If ModoBig > 0 Then
        
        If InventoryIndex <> eCaption.MainPicture Then
            Mult = 2
            ' Div = 2
            NotBig = True
            
            Call Draw_Texture_Graphic_Gui(133, 0, 0, To_Depth(1), 245, 140, 0, 0, 245, 140, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
        Else
            Div = 1
            NotBig = False

        End If
    
        If InventoryIndex = eCaption.MainPicture Then
            
            ' Call Draw_Texture_Graphic_Gui(51, 0, 0, To_Depth(1), 160 * 2, 192 * 2, 0, 0, 160 * 2, 192 * 2, -1, 0, eTechnique.t_Default)

        End If
        
    #Else
          If InventoryIndex <> eCaption.MainPicture Then
            Call Draw_Texture_Graphic_Gui(133, 0, 0, To_Depth(1), 245, 140, 0, 0, 245, 140, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
        End If
            
        NotBig = True
        

       ' Call Draw_Texture_Graphic_Gui(51, 0, 0, To_Depth(1), 160, 192, 0, 0, 160, 192, -1, 0, eTechnique.t_Default)
    #End If

    Dim Last  As Long

    Dim Value As Long
    
    Last = UBound(Inventory) - InventoryOffset
    
    For LoopC = InventoryOffset + 1 To UBound(Inventory)

        With Inventory(LoopC)
        
            Dim DestX As Long, DestY As Long
            
            If Last > 0 Then
                DestX = (((LoopC - InventoryOffset - 1) Mod lastTileRow) * ownTilePixelWidth)
                DestY = ((LoopC - InventoryOffset - 1) \ lastTileRow) * ownTilePixelHeight
            Else
                DestX = (((LoopC - 1) Mod lastTileRow) * ownTilePixelWidth)
                DestY = ((LoopC - 1) \ lastTileRow) * ownTilePixelHeight

            End If
            
            If PictureModernOn Then
                If ClientSetup.bResolution = 0 Then
                    Call Draw_Texture_Graphic_Gui(3, DestX, DestY, To_Depth(2), ownTilePixelWidth / Div, ownTilePixelWidth / Div, 0, 0, ownTilePixelWidth / Div, ownTilePixelWidth / Div, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
                End If
            End If
            
            
            If .GrhIndex Then
                
               ' If ObjData(.ObjIndex).Color <> 0 Then
                   ' Call Draw_Texture(86747, DestX, DestY, To_Depth(3), ownTilePixelWidth / Div, ownTilePixelWidth / Div, ObjData(.ObjIndex).Color, 0, eTechnique.t_Alpha)
               ' End If
               
               If .ExistSkin > 0 Then
                      If Skins_CheckingItems(.ObjIndex) Then
                            Call Draw_Texture(6579, DestX, DestY, To_Depth(7), ownTilePixelWidth, ownTilePixelWidth, ARGB(0, 255, 0, 255), 0, eTechnique.t_Alpha)
                    
                    Else
                        Call Draw_Texture(6579, DestX, DestY, To_Depth(6), ownTilePixelWidth, ownTilePixelWidth, ARGB(255, 0, 0, 255), 0, eTechnique.t_Alpha)
                End If
                    
                    
                  
                
               End If
               
                
              
                
                If ((LoopC) = SelectedItem) Then
                    Call Draw_Texture(6579, DestX, DestY, To_Depth(7), ownTilePixelWidth, ownTilePixelWidth, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
                Else

                    'Call Draw_Texture(24146, DestX, DestY, -1, GrhData(.GrhIndex).pixelWidth, GrhData(.GrhIndex).pixelHeight, -1, 0, Techniques(eTechnique.t_Alpha))
                End If
                
                If .CanUse Then
                    Call Draw_Texture(.GrhIndex, DestX + 3, DestY + 2, To_Depth(5), (GrhData(.GrhIndex).pixelWidth / Mult), (GrhData(.GrhIndex).pixelHeight / Mult), ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
                Else
                    'Call Draw_Texture(30652, DestX, DestY, -1, 32, 32, ARGB(200, 0, 0, 255), 0, Techniques(eTechnique.t_Alpha))
                    Call Draw_Texture(.GrhIndex, DestX + 3, DestY + 2, To_Depth(6), (GrhData(.GrhIndex).pixelWidth / Mult), (GrhData(.GrhIndex).pixelHeight / Mult), ARGB(255, 20, 20, 255), 0, eTechnique.t_Alpha)

                End If
                
                If .CanUse Then
                        
                Else

                    If PictureModernOn Then Call Draw_Texture(30652, DestX, DestY, To_Depth(1), ownTilePixelWidth, ownTilePixelWidth, ARGB(200, 0, 0, 255), 0, eTechnique.t_Alpha)

                End If
                
                If ShowValue Then
                    If .Amount > 1 Then
                        
                        Call Draw_Text(f_Tahoma, 14, DestX, DestY, To_Depth(7), 0#, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP, CStr(.Amount), False, NotBig)

                    End If

                End If
                
                'If equipped we render "E"
                If ShowEquipped Then
                    If .Equipped Then
                        If InventoryIndex = eCaption.MainPicture Then
                            #If ModoBig > 0 Then
                                Call Draw_Text(f_Tahoma, 14, DestX + 5, DestY + 40, To_Depth(6), 0#, &HFFFFFF00, FONT_ALIGNMENT_TOP, "E", False)
                            #Else
                                Call Draw_Text(f_Tahoma, 14, DestX + 5, DestY + 5, To_Depth(6), 0#, &HFFFFFF00, FONT_ALIGNMENT_TOP, "E", False)
                            #End If
                        Else
                            Call Draw_Texture(6579, DestX, DestY, To_Depth(6), ownTilePixelWidth, ownTilePixelWidth, ARGB(0, 255, 0, 255), 0, eTechnique.t_Alpha)
                        End If
                    End If

                End If
                  
            Else

                'If IsPremium Then Call Draw_Texture(24146, DestX, DestY, -1, 32, 32, -1, 0, Techniques(eTechnique.t_Alpha))
            End If

        End With

    Next LoopC
    
    Call wGL_Graphic_Renderer.Flush
    
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
    Dim temp_x   As Integer

    Dim temp_y   As Integer

    Dim TempItem As Integer
    
    Dim Slot     As Integer
    
    Dim Temp     As String
    
    Dim ObjIndex As Integer
    
    Dim Last As Long
    
    'Exit if it got outside the control's areanv
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then
        If dragAndDropOn Then
            If uMoveItem = True Then sMoveItem = True

        End If

        Exit Sub

    End If
    
     Last = UBound(Inventory) - InventoryOffset
    
    If Button = vbRightButton Then
        uMoveItem = True
        
    Else
        temp_x = X \ ownTilePixelWidth
        temp_y = Y \ ownTilePixelHeight
        
        TempItem = temp_x + (temp_y) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 + (UBound(Inventory) - Last)
    
        If TempItem <= UBound(Inventory) Then

            ObjIndex = Inventory(TempItem).ObjIndex
                
            If ObjIndex > 0 Then
                InventoryWindow.ToolTipText = ObjData(ObjIndex).Name
            End If
            
            If ShowInfo Then
                If ObjIndex > 0 Then
                    If InventoryIndex <> eCaption.MainPicture Then
                        Call ShowInfoItem(ObjIndex)
                    End If
                    
                Else
                    If MirandoObjetos Then
                       FrmObject_Info.Close_Form
                    End If
                    
                End If
            End If
            
        Else
            If MirandoObjetos Then
                FrmObject_Info.Close_Form
            End If
        End If

    End If

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)

    Dim temp_x   As Integer

    Dim temp_y   As Integer

    Dim TempItem As Integer
    
    Dim Last As Long
    
    Last = UBound(Inventory) - InventoryOffset
    
    If dragAndDropOn Then

        'Exit if it got outside the control's area
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
        
        If Button = vbRightButton Then
            temp_x = X \ ownTilePixelWidth
            temp_y = Y \ ownTilePixelHeight
        
            TempItem = temp_x + (temp_y) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 + (UBound(Inventory) - Last)
            
            slotDragged = TempItem
        End If
    End If

End Sub

Public Sub Initialize(ByRef InvPic As PictureBox, _
                      ByVal MaxObjs_Render As Integer, _
                      ByVal MaxObjs As Integer, _
                      ByVal CaptionIndex As Integer, _
                      Optional ByVal TileWidth As Integer = 35, _
                      Optional ByVal TileHeight As Integer = 35, _
                      Optional ByVal startX As Integer = 0, _
                      Optional ByVal startY As Integer = 0, _
                      Optional ByVal bImgContainer As Boolean = True, _
                      Optional ByVal bShowText As Boolean = True, _
                      Optional FormatValue As Boolean = False, _
                      Optional ByVal dragAndDrop As Boolean = False, _
                      Optional ByVal eViewInfo As Boolean = False, _
                      Optional ByVal PictureModern As Boolean = False, _
                      Optional ByVal bShowEquipped As Boolean = False, _
                      Optional ByVal eShowInfo As Boolean = False)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 03/12/09
    'Sets the reference to the picture box on which to render the inventory
    '03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
    '***************************************************
    Set InventoryWindow = InvPic
    ValueFormated = FormatValue
    
    'Make sure auto-redraw is set to false
    InventoryWindow.AutoRedraw = False
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    InventoryIndex = CaptionIndex
    
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ShowEquipped = bShowEquipped
    ImgContainer = bImgContainer
    ShowValue = bShowText
   
    dragAndDropOn = dragAndDrop
    
    PictureModernOn = True 'PictureModern
    
    ViewInfo = eViewInfo
    If ImgContainer Then InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

    Inventory_Pictures = MaxObjs_Render
    
    ShowInfo = eShowInfo
   ' Call Invalidate(InventoryWindow.hWnd)
    
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 27/07/04
    'Implements the mouse up event of the inventory picture box
    'Check outs which item was clicked
    '***************************************************
    'Store previously selected item
    Dim prevSelItem As Long

    Dim TempItem    As Long

    Dim temp_x      As Long

    Dim temp_y      As Long
    
    Dim Last As Long
    
    Dim ObjIndex As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
    
    prevSelItem = InvSelectedItem
    Last = UBound(Inventory) - InventoryOffset
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
            
            
    InvSelectedItem = temp_x + (temp_y) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 + (UBound(Inventory) - Last)
            
    If Button = vbRightButton Then
        If dragAndDropOn Then
           
            If InvSelectedItem <> slotDragged Then
                
                If frmBancoObj.visible Then
                    Call Protocol.WriteMoveItem(slotDragged, InvSelectedItem, eMoveType.Bank)
                Else
                    Call Protocol.WriteMoveItem(slotDragged, InvSelectedItem, eMoveType.Inventory)
                End If
                
                uMoveItem = False
                sMoveItem = False
                FrmMain.PicInv.MouseIcon = Nothing
            End If
            

        End If
        
        
        ObjIndex = Inventory(InvSelectedItem).ObjIndex
        
        If ObjIndex > 0 Then
            'If (Shift = 1) Then
              '  Call ShowInfoItem(ObjIndex)
           ' End If
        Else
             'If MirandoObjetos Then
                      '  FrmObject_Info.Close_Form
                   ' End If
        End If
    Else

        'Get the currently clickced item
        If Not sMoveItem Then InvSelectedItem = ClickItem(CInt(X), CInt(Y))
        
        
      '  Call Invalidate(InventoryWindow.hWnd)
        
        If uMoveItem And InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 And Not sMoveItem And Not Comerciando Then
        
            Call Protocol.WriteMoveItem(InvSelectedItem, MouseDownSelectedItem, eMoveType.Inventory)
            uMoveItem = False
        Else
            uMoveItem = False
            
        End If
         
        'Update needed inventory slots
        If prevSelItem <> InvSelectedItem Then
            If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventory

            If InvSelectedItem Then Call DrawInventory
        End If
    End If

End Sub

Private Sub InventoryWindow_Paint()
    Call DrawInventory
End Sub
