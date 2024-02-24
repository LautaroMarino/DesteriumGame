VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmComerciar 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmComerciar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tUpdate 
      Interval        =   50
      Left            =   840
      Top             =   240
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   360
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2190
      TabIndex        =   2
      Text            =   "1"
      Top             =   4350
      Width           =   810
   End
   Begin VB.PictureBox picInvNpc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   810
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   2175
      Width           =   3675
   End
   Begin VB.PictureBox picInvUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   825
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   4770
      Width           =   3675
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "frmComerciar.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblPriceDSP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   3270
      TabIndex        =   6
      Top             =   1830
      Width           =   1575
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   4
      Top             =   1830
      Width           =   1575
   End
   Begin VB.Image imgValueGLD 
      Height          =   360
      Left            =   360
      Picture         =   "frmComerciar.frx":10BE
      Top             =   1755
      Width           =   2010
   End
   Begin VB.Image imgValueDSP 
      Height          =   360
      Left            =   2880
      Picture         =   "frmComerciar.frx":3A59
      Top             =   1755
      Width           =   2010
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   6945
      Width           =   1545
   End
   Begin VB.Label lblGld 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   6945
      Width           =   1545
   End
   Begin VB.Image imgMas 
      Height          =   210
      Left            =   3075
      Top             =   4410
      Width           =   255
   End
   Begin VB.Image imgMenos 
      Height          =   210
      Left            =   1875
      Top             =   4410
      Width           =   255
   End
   Begin VB.Image imgQuest 
      Height          =   435
      Left            =   210
      MouseIcon       =   "frmComerciar.frx":6463
      MousePointer    =   99  'Custom
      Top             =   315
      Width           =   1695
   End
   Begin VB.Image imgVender 
      Height          =   375
      Left            =   3360
      MouseIcon       =   "frmComerciar.frx":65B5
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Image imgComprar 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmComerciar.frx":6707
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   4320
      Width           =   1155
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
' Copyright (C) 2002 Márquez Pablo Ignacio
' Copyright (C) 2002 Otto Perez
' Copyright (C) 2002 Aaron Perkins
' Copyright (C) 2002 Matías Fernando Pequeño
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
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Type tChar
    Body As Integer
    Head As Integer
    Helm As Integer
    Weapon As Integer
    Shield As Integer
        
    Heading As E_Heading
End Type

Private Char As tChar


Private Enum eSelectedPrice
    eGLD = 0
    eDSP = 1
End Enum

Private SelectedPrice As eSelectedPrice

Private clsFormulario As clsFormMovementManager

Public LastIndex1     As Integer

Public LastIndex2     As Integer

Public LasActionBuy   As Boolean

Private ClickNpcInv   As Boolean

Private lIndex        As Byte


' Consola Transparente
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&


' Botones Gráficos
Private cBotonComprar     As clsGraphicalButton
Private cBotonVender        As clsGraphicalButton
Private cBotonMenos        As clsGraphicalButton
Private cBotonMas      As clsGraphicalButton
Public LastButtonPressed   As clsGraphicalButton

Public MouseX As Long
Public MouseY As Long

Const WM_NCMOUSEMOVE = &HA0
Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Enum eSolapa
    Npc = 1
    User = 2
End Enum

Public SolapaView As eSolapa

Private Sub cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Click()
    InvComUser.DeselectItem
    InvComNpc.DeselectItem
End Sub


Private Sub Form_Load()
    cantidad.Text = "1"
    
    Dim I As Long
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
          
    Call LoadButtons
    
    ' Ventanas de comercio con NPCS
    g_Captions(eCaption.Comercio_Npc) = wGL_Graphic.Create_Device_From_Display(frmComerciar.picInvNpc.hWnd, frmComerciar.picInvNpc.ScaleWidth, frmComerciar.picInvNpc.ScaleHeight)
    g_Captions(eCaption.Comercio_User) = wGL_Graphic.Create_Device_From_Display(frmComerciar.picInvUser.hWnd, frmComerciar.picInvUser.ScaleWidth, frmComerciar.picInvUser.ScaleHeight)

    Me.Picture = LoadPicture(App.path & "\resource\interface\commerce\commerce.jpg")
    
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = DirInterface

    Set cBotonComprar = New clsGraphicalButton
    Set cBotonVender = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
   
    Set LastButtonPressed = New clsGraphicalButton
    'Call cBotonComprar.Initialize(imgComprar, vbNullString, GrhPath & "commerce\BotonComprarActivo.jpg", vbNullString, Me)
'    Call cBotonVender.Initialize(imgVender, vbNullString, GrhPath & "commerce\BotonVenderActivo.jpg", vbNullString, Me)
  '  Call cBotonMenos.Initialize(imgMenos, vbNullString, GrhPath & "commerce\BotonMenosActivo.jpg", vbNullString, Me)
   ' Call cBotonMas.Initialize(imgMas, vbNullString, GrhPath & "commerce\BotonMasActivo.jpg", vbNullString, Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
     
     MouseX = X
     MouseY = Y
     
    
    If MirandoObjetos Then
        FrmObject_Info.Close_Form
    End If
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    MirandoComerciar = False
    Call WriteCommerceEnd
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.Comercio_Npc))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.Comercio_User))
    
    If MirandoObjetos Then
       Call FrmObject_Info.Close_Form
    End If
   
    FrmMain.SetFocus
End Sub

Public Function UpdatePrice() As Long

    Dim ItemSlot As Byte
    
    If ClickNpcInv Then
        
        ItemSlot = InvComNpc.SelectedItem
    
        If ItemSlot = 0 Then Exit Function
        
        UpdatePrice = CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text))
        
    Else
        ItemSlot = InvComUser.SelectedItem
        
        If ItemSlot = 0 Then Exit Function
        
        UpdatePrice = CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text))

    End If
    
    
    

End Function

Private Sub cantidad_Change()

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
          
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If

    lblPrice.Caption = PonerPuntos(UpdatePrice)
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If

End Sub



''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, _
                                   ByVal ObjAmount As Long) As Long

    '*************************************************
    'Author: Marco Vanotti (MarKoxX)
    'Last modified: 19/08/2008
    'Last modify by: Franco Zeoli (Noich)
    '*************************************************
    On Error GoTo error

    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * ObjAmount)
          
    Exit Function

error:
    MsgBox err.Description, vbExclamation, "Error: " & err.Number
End Function

Private Sub imgComprar_Click()
    
    If Timer.Enabled Then Exit Sub

    ' Debe tener seleccionado un item para comprarlo.
    If InvComNpc.SelectedItem = 0 Then Exit Sub
          
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
          
    Call Audio.PlayInterface(SND_CLICK)
          
    LasActionBuy = True
    
    Dim ObjIndex As Integer
    
    ObjIndex = NPCInventory(InvComNpc.SelectedItem).ObjIndex
    
    Dim A As Long
    
    If ObjData(ObjIndex).Upgrade.RequiredCant > 0 Then
        For A = 1 To ObjData(ObjIndex).Upgrade.RequiredCant
            If TieneObjetos(ObjData(ObjIndex).Upgrade.Required(A).ObjIndex) < ObjData(ObjIndex).Upgrade.Required(A).Amount Then
                Call ShowConsoleMsg("¡No tienes " & ObjData(ObjData(ObjIndex).Upgrade.Required(A).ObjIndex).Name & " (x" & ObjData(ObjIndex).Upgrade.Required(A).Amount & ")")
                Exit Sub
            End If
        Next A
    End If
    
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        'If MsgBox("¿Estas seguro que deseas comprar este objeto?", vbYesNo) = vbYes Then
            Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text), SelectedPrice)
        'End If
    Else
        Call AddtoRichTextBox(FrmMain.RecTxt, "Se necesita más oro.", 2, 51, 223, 1, 1)

        Exit Sub

    End If
          
    'InvComUser.DrawInventory
    'InvComNpc.DrawInventory
    Timer.Enabled = True
End Sub

Private Sub imgCross_Click()


End Sub

Private Function CheckAdding(ByVal Value As Long) As Long
    If Value <= 100 Then
         CheckAdding = Value + 1
    ElseIf Value <= 1000 Then
        CheckAdding = Value + 100
    ElseIf Value <= 10000 Then
        CheckAdding = Value + 1000
    Else
        CheckAdding = MAX_INVENTORY_OBJS
    End If
End Function
Private Sub imgMas_Click()
          
    cantidad.Text = CheckAdding(cantidad.Text)
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    
End Sub
Private Function CheckAdding_Menos(ByVal Value As Long) As Long
    If Value <= 100 Then
         CheckAdding_Menos = Value - 1
    ElseIf Value <= 1000 Then
        CheckAdding_Menos = Value - 100
    ElseIf Value <= 10000 Then
        CheckAdding_Menos = Value - 1000
    Else
        CheckAdding_Menos = 1
    End If
End Function
Private Sub imgMenos_Click()
          
    cantidad.Text = CheckAdding_Menos(cantidad.Text)
    
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
End Sub
Private Sub imgQuest_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If QuestLast = 0 Then
        Call MsgBox("El comerciante no puede fabricar objetos.")
    Else
        InitQuest = True
        Call Invalidate(frmCriatura_Quest.hWnd)
        Call frmCriatura_Quest.Show(, FrmMain)
        Me.Hide
        
        'imgCross_Click
    End If
    
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub imgValueDSP_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ValuePrize_Selected(eDSP)
End Sub

Private Sub imgValueGLD_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ValuePrize_Selected(eGLD)
End Sub
Private Sub lblPrice_Click()
    imgValueGLD_Click
End Sub

Private Sub lblPriceDSP_Click()
    imgValueDSP_Click
End Sub

Private Sub Detect_Prize()
    If InvComNpc.SelectedItem > 0 Then
        If InvComNpc.Valor(InvComNpc.SelectedItem) > 0 Then
            ValuePrize_Selected eGLD
            Exit Sub
        End If
        
        If InvComNpc.ValorAzul(InvComNpc.SelectedItem) > 0 Then
            ValuePrize_Selected eDSP
            Exit Sub
        End If
    ElseIf InvComUser.SelectedItem > 0 Then
        If InvComUser.Valor(InvComUser.SelectedItem) > 0 Then
            ValuePrize_Selected eGLD
            Exit Sub
        End If
        
        If InvComUser.ValorAzul(InvComUser.SelectedItem) > 0 Then
            ValuePrize_Selected eDSP
            Exit Sub
        End If
    End If
End Sub
Private Sub ValuePrize_Selected(ByRef Tipo As eSelectedPrice)
    
    Select Case Tipo
    
        Case eSelectedPrice.eGLD

            If Val(lblPrice.Caption) = 0 Then
                Call ShowConsoleMsg("¡El Objeto no es comprado por Monedas de Oro!")
                Exit Sub

            End If
            
            imgValueGLD.Picture = LoadPicture(DirInterface & "\shop\gld_hover.jpg")
            imgValueDSP.Picture = LoadPicture(DirInterface & "\shop\dspvalue.jpg")

        Case eSelectedPrice.eDSP

            If Val(lblPriceDSP.Caption) = 0 Then
                Call ShowConsoleMsg("¡El Objeto no es comprado por DSP!")
                Exit Sub

            End If

            imgValueDSP.Picture = LoadPicture(DirInterface & "\shop\dspvalue_hover.jpg")
            imgValueGLD.Picture = LoadPicture(DirInterface & "\shop\gld.jpg")

    End Select
    
    SelectedPrice = Tipo

End Sub

Private Sub imgVender_Click()

    If Timer.Enabled Then Exit Sub
    
    ' Debe tener seleccionado un item para comprarlo.
    If InvComUser.SelectedItem = 0 Then Exit Sub

    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
          
    Call Audio.PlayInterface(SND_CLICK)
          
    LasActionBuy = False

    Call WriteCommerceSell(InvComUser.SelectedItem, Val(cantidad.Text), SelectedPrice)
    'InvComUser.DrawInventory
    'InvComNpc.DrawInventory
    Timer.Enabled = True

End Sub



Private Sub picInvNpc_Click()

    Dim ItemSlot As Byte
          
    ItemSlot = InvComNpc.SelectedItem
    Call Audio.PlayInterface(SND_CLICK)

    If ItemSlot = 0 Then Exit Sub
          
    ClickNpcInv = True
    InvComUser.DeselectItem
    
    lblPrice = Format$(CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)), "##,##")
    lblPriceDSP = Format$(CalculateSellPrice(NPCInventory(ItemSlot).ValorAzul, Val(cantidad.Text)), "##,##")
    
    Call Detect_Prize
   Call Char_PreSelected(NPCInventory(ItemSlot).ObjIndex)
    
   ' picInvView.visible = True
    
End Sub

Private Sub Char_PreSelected(ByVal ObjIndex As Integer)
    
    Char.Body = CharList(UserCharIndex).iBody
    Char.Head = CharList(UserCharIndex).iHead
    Char.Heading = E_Heading.SOUTH
    
    If ObjIndex = 0 Then Exit Sub
    
    With ObjData(ObjIndex)
    
        Select Case .ObjType
    
            Case eOBJType.otarmadura
                Char.Body = .Anim
                If (UserRaza = Gnomo Or UserRaza = Enano) And .AnimBajos > 0 Then Char.Body = .AnimBajos
                    
            Case eOBJType.otWeapon
                Char.Weapon = .Anim
                
            Case eOBJType.otescudo
                Char.Shield = .Anim
                
            Case eOBJType.otcasco
                Char.Helm = .Anim
        End Select
    
    End With
 
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyLeft Then
        Char.Heading = E_Heading.WEST
    ElseIf KeyCode = vbKeyDown Then
        Char.Heading = E_Heading.SOUTH
    ElseIf KeyCode = vbKeyRight Then
        Char.Heading = E_Heading.EAST
    ElseIf KeyCode = vbKeyUp Then
        Char.Heading = E_Heading.NORTH
    End If
    
End Sub

Private Sub picInvNpc_DblClick()
    Dim ItemSlot As Byte
    ItemSlot = InvComNpc.SelectedItem

    If ItemSlot = 0 Then Exit Sub
    SelectedObjIndex = NPCInventory(ItemSlot).ObjIndex
    SelectedObjIndex_Update
End Sub

Private Sub picInvNpc_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub picInvUser_Click()

    Dim ItemSlot As Byte
          
    ItemSlot = InvComUser.SelectedItem
          
    If ItemSlot = 0 Then Exit Sub
          
    ClickNpcInv = False
    InvComNpc.DeselectItem
          
    lblPrice = Format$(CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text)), "##,##")
    lblPriceDSP = Format$(CalculateBuyPrice(Inventario.ValorAzul(ItemSlot), Val(cantidad.Text)), "##,##")
    
    
    Detect_Prize
End Sub

Private Sub picInvUser_DblClick()
    Dim ItemSlot As Byte
    ItemSlot = InvComNpc.SelectedItem

    If ItemSlot = 0 Then Exit Sub
    SelectedObjIndex = Inventario.ObjIndex(ItemSlot)
    SelectedObjIndex_Update
End Sub

Private Sub picInvUser_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub picInvView_Click()
   ' Me.picInvView.visible = False
End Sub

Private Sub picInvView_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
End Sub

' Drag And Drop

Private Function BuscarI(gh As Integer) As Integer

    Dim I As Integer
       
    For I = 1 To frmComerciar.ImageList1.ListImages.Count

        If frmComerciar.ImageList1.ListImages(I).Key = "g" & CStr(gh) Then
            BuscarI = I

            Exit For

        End If

    Next I
       
End Function

Private Sub PicInvUser_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    LastButtonPressed.ToggleToNormal
    
    Dim Position  As Integer

    Dim I         As Long

    Dim file_path As String

    Dim data()    As Byte

    Dim bmpInfo   As BITMAPINFO

    Dim handle    As Integer

    Dim bmpData   As StdPicture

    Dim Last_I    As Long

    MouseX = X
    MouseY = Y
    SolapaView = User
    
    If (Button = vbRightButton) Then
        If InvComUser.SelectedItem = 0 Then Exit Sub
        
        If InvComUser.GrhIndex(InvComUser.SelectedItem) > 0 Then

            Last_I = InvComUser.SelectedItem

            If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                          
                Position = BuscarI(3057)
                  
                If Position = 0 Then
                    I = GrhData(InvComUser.GrhIndex(InvComUser.SelectedItem)).FileNum
                    Call Get_Image(DirGraficos & GRH_RESOURCE_FILE_DEFAULT, CStr(3057), data, False)
                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1)
                    ImageList1.ListImages.Add , "g3057", Picture:=bmpData
                    Position = ImageList1.ListImages.Count
                    Set bmpData = Nothing
                End If
                  
                '  InvComUsu.uMoveItem = True
                  
                Set picInvUser.MouseIcon = ImageList1.ListImages(Position).ExtractIcon
                picInvUser.MousePointer = vbCustom

                Exit Sub

            End If
        End If
    End If
    
    
End Sub

Private Sub PicInvNpc_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    LastButtonPressed.ToggleToNormal
    
    Dim Position  As Integer

    Dim I         As Long

    Dim file_path As String

    Dim data()    As Byte

    Dim bmpInfo   As BITMAPINFO

    Dim handle    As Integer

    Dim bmpData   As StdPicture

    Dim Last_I    As Long
    
   
     MouseX = X
     MouseY = Y
     SolapaView = Npc
     
    If (Button = vbRightButton) Then
        If InvComNpc.SelectedItem = 0 Then Exit Sub
        If InvComNpc.GrhIndex(InvComNpc.SelectedItem) > 0 Then

            Last_I = InvComNpc.SelectedItem

            If Last_I > 0 And Last_I <= MAX_NPC_INVENTORY_SLOTS Then
                          
                Position = BuscarI(3057)
                  
                If Position = 0 Then
                    I = GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum
                    Call Get_Image(DirGraficos & GRH_RESOURCE_FILE_DEFAULT, CStr(3057), data, False)
                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                    ImageList1.ListImages.Add , "g3057", Picture:=bmpData
                    Position = ImageList1.ListImages.Count
                    Set bmpData = Nothing
                End If
                  
                '  InvComNpc.uMoveItem = True
                  
                Set picInvNpc.MouseIcon = ImageList1.ListImages(Position).ExtractIcon
                picInvNpc.MousePointer = vbCustom

                Exit Sub

            End If
        End If
        
        
    Else

    End If

End Sub

Private Sub picInvNpc_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
       
    If (Button = vbRightButton) Then
        If Not (X > 0 And X < picInvNpc.ScaleWidth And Y > 0 And Y < picInvNpc.ScaleHeight) Then
            Call imgComprar_Click
        End If
    End If
    
    picInvNpc.MousePointer = vbDefault
    
    InvComNpc.sMoveItem = False
    InvComNpc.uMoveItem = False
End Sub

Private Sub picInvUser_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
       
    If (Button = vbRightButton) Then
        If Not (X > 0 And X < picInvUser.ScaleWidth And Y > 0 And Y < picInvUser.ScaleHeight) Then
            Call imgVender_Click
        End If
    End If
    
    picInvUser.MousePointer = vbDefault
    
    InvComUser.sMoveItem = False
    InvComUser.uMoveItem = False
       
End Sub


Private Sub tUpdate_Timer()
    If Not Me.visible Then Exit Sub
    InvComNpc.DrawInventory
    InvComUser.DrawInventory
    
End Sub

