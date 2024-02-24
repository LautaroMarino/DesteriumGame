VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "frmBancoObj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   0  'User
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBancoInv 
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
      Left            =   840
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3675
   End
   Begin VB.Timer tUpdate 
      Interval        =   200
      Left            =   3360
      Top             =   240
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3840
      Top             =   240
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2220
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "1"
      Top             =   4380
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.PictureBox PicInv 
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
      Left            =   840
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4800
      Width           =   3675
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "frmBancoObj.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Label GldUser 
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   6960
      Width           =   1545
   End
   Begin VB.Label lblDsp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9.999.999"
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
      Height          =   285
      Left            =   3180
      TabIndex        =   5
      Top             =   1845
      Width           =   1635
   End
   Begin VB.Label lblGld 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9.999.999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1845
      Width           =   1635
   End
   Begin VB.Image imgDsp 
      Height          =   285
      Left            =   2745
      OLEDropMode     =   1  'Manual
      Top             =   1815
      Width           =   2175
   End
   Begin VB.Image imgGld 
      Height          =   285
      Left            =   435
      OLEDropMode     =   1  'Manual
      Top             =   1815
      Width           =   2175
   End
   Begin VB.Image imgScroll 
      Height          =   375
      Index           =   3
      Left            =   4560
      Top             =   5760
      Width           =   375
   End
   Begin VB.Image imgScroll 
      Height          =   375
      Index           =   2
      Left            =   4560
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgScroll 
      Height          =   375
      Index           =   1
      Left            =   4560
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgScroll 
      Height          =   375
      Index           =   0
      Left            =   4560
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   5040
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   3480
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   0
      Left            =   600
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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




'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private clsFormulario As clsFormMovementManager

Private Last_I        As Long

Dim Button            As Integer

Public Attack         As Boolean

Public tX             As Byte

Public tY             As Byte

Public MouseX         As Long

Public MouseY         As Long

Public MouseBoton     As Long

Public MouseShift     As Long

Private clicX         As Long

Private clicY         As Long

Private ClickNpcInv   As Boolean

Public LasActionBuy   As Boolean

Public LastIndex1     As Integer

Public LastIndex2     As Integer

Public NoPuedeMover   As Boolean

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&

Public Enum eSolapaBanco
    Npc = 1
    User = 2
End Enum

Public Enum eModo
        M_GLD = 1
        M_DSP = 2
End Enum

Public ModoBank As eModo
Public SolapaView As eSolapaBanco

Private Sub cantidad_Change()

    If Val(cantidad.Text) <= 1 Then
        cantidad.Text = 1
    End If
    
    If Not IsNumeric(cantidad.Text) Then
        cantidad.Text = 1
    End If
        
    If ModoBank = M_GLD Or ModoBank = M_DSP Then
        If Val(cantidad.Text) > 10000000 Then
            cantidad.Text = 10000000
            cantidad.SelStart = Len(cantidad.Text)
        End If
    Else
        If Val(cantidad.Text) > 10000 Then
            cantidad.Text = 10000
            cantidad.SelStart = Len(cantidad.Text)
        End If
    End If
    
    Dim ItemSlot As Byte

End Sub

Private Sub cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If

End Sub


Private Sub Form_Click()
    If Not InvComNpc Is Nothing Then
    InvComNpc.DeselectItem
    End If
    
    If Not InvComUsu Is Nothing Then
    InvComUsu.DeselectItem
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmMain.SetFocus
        Unload Me
    End If
End Sub

Private Sub Form_Load()


    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    ' Ventanas de BOVEDA
    g_Captions(eCaption.Boveda_Npc) = wGL_Graphic.Create_Device_From_Display(frmBancoObj.PicBancoInv.hWnd, frmBancoObj.PicBancoInv.ScaleWidth, frmBancoObj.PicBancoInv.ScaleHeight)
    g_Captions(eCaption.Boveda_User) = wGL_Graphic.Create_Device_From_Display(frmBancoObj.PicInv.hWnd, frmBancoObj.PicInv.ScaleWidth, frmBancoObj.PicInv.ScaleHeight)
    
    Me.Picture = LoadPicture(App.path & "\resource\interface\bank\bank_item.jpg")
    
    lblGld.Caption = IIf(UserBankGold > 0, Format$(UserBankGold, "##,##"), "0")
    lblDSP.Caption = IIf(UserBankEldhir > 0, Format$(UserBankEldhir, "##,##"), "0")
    
    GldUser.Caption = PonerPuntos(UserGLD)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If MirandoObjetos Then
        FrmObject_Info.Close_Form
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmMain.SetFocus
    
    Call WriteBankEnd
    MirandoBanco = False
    NoPuedeMover = False

    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.Boveda_Npc))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.Boveda_User))

    If MirandoObjetos Then
        Call FrmObject_Info.Close_Form
    End If

End Sub

Public Sub Bank_Money(ByVal Modo As eModo, ByVal ModoExtract As Boolean)

    Select Case Modo
        Case eModo.M_GLD
            Call WriteBankGold(Val(cantidad.Text), 0, ModoExtract)
        Case eModo.M_DSP
            Call WriteBankGold(Val(cantidad.Text), 1, ModoExtract)
    End Select
    
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
          
    If ModoBank > 0 Then

        Select Case Index

            Case 0
                Call Bank_Money(ModoBank, True)

            Case 1
                Call Bank_Money(ModoBank, False)

        End Select
    
        Exit Sub

    End If
    
    If Timer.Enabled Then Exit Sub
     
    If InvBanco(Index).SelectedItem = 0 Then Exit Sub
          
    If Not IsNumeric(cantidad.Text) Then Exit Sub
          
    Select Case Index

        Case 0
            LastIndex1 = InvBanco(0).SelectedItem
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text, SelectedBank)

        Case 1
            LastIndex2 = InvBanco(1).SelectedItem
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text, SelectedBank)

    End Select
    
    Timer.Enabled = True

End Sub

Private Sub SelectedModo(ByVal Modo As eModo)
    ModoBank = Modo
    
    Select Case ModoBank
    
        Case eModo.M_DSP
            imgDsp.Picture = LoadPicture(App.path & "\resource\interface\bank\dsp_clic.jpg")
            imgGld.Picture = Nothing
        Case eModo.M_GLD
            imgGld.Picture = LoadPicture(App.path & "\resource\interface\bank\gld_clic.jpg")
            imgDsp.Picture = Nothing
    End Select
    
End Sub
Private Sub imgDsp_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    InvBanco(0).DeselectItem
    InvBanco(1).DeselectItem
    
    SelectedModo (M_DSP)
End Sub

Private Sub imgGld_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    InvBanco(0).DeselectItem
    InvBanco(1).DeselectItem
    
    SelectedModo (M_GLD)
    
End Sub

Private Sub imgScroll_Click(Index As Integer)
    
    Select Case Index
    
        Case 0 ' Scroll volver >> Banco
            InvBanco(0).ScrollInventory (False)
        Case 1 ' Scroll Avanzar >> Banco
            InvBanco(0).ScrollInventory (True)
        Case 2 ' Scroll Volver >> User
            InvBanco(1).ScrollInventory (False)
        Case 3 ' Scroll Avanzar >> User
            InvBanco(1).ScrollInventory (True)
    End Select
    
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub lblDsp_Click()
    imgDsp_Click
End Sub

Private Sub lblGld_Click()
    imgGld_Click
End Sub

Private Sub PicBancoInv_Click()
    If ModoBank > 0 Then
        ModoBank = 0
        imgDsp.Picture = Nothing
        imgGld.Picture = Nothing
        
        If Val(cantidad.Text) > 10000 Then
            cantidad.Text = 10000
            cantidad.SelStart = Len(cantidad.Text)
        End If
    End If
    
End Sub

Private Sub PicBancoInv_DblClick()
    Dim ItemSlot As Byte
    ItemSlot = InvBanco(0).SelectedItem

    If ItemSlot = 0 Then Exit Sub
    SelectedObjIndex = InvBanco(0).ObjIndex(InvBanco(0).SelectedItem)
    SelectedObjIndex_Update
End Sub

Private Sub PicBancoInv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub PicInv_DblClick()
    Dim ItemSlot As Byte
    ItemSlot = InvBanco(1).SelectedItem

    If ItemSlot = 0 Then Exit Sub
    SelectedObjIndex = InvBanco(1).ObjIndex(InvBanco(1).SelectedItem)
    SelectedObjIndex_Update
End Sub


''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, _
                                    ByVal ObjAmount As Long) As Long

    '*************************************************
    'Author: Marco Vanotti (MarKoxX)
    'Last modified: 19/08/2008
    'Last modify by: Franco Zeoli (Noich)
    '*************************************************
    On Error GoTo error

    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * ObjAmount + 0.5
          
    Exit Function

error:
    MsgBox err.Description, vbExclamation, "Error: " & err.Number
End Function

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

Private Sub PicInv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
End Sub


' Drag And Drop
Private Function Search_GhID(gh As Integer) As Integer

    Dim I As Integer

    For I = 1 To frmBancoObj.ImageList1.ListImages.Count

        If frmBancoObj.ImageList1.ListImages(I).Key = "g" & CStr(gh) Then
            Search_GhID = I

            Exit For

        End If

    Next I

End Function
Private Sub PicInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Dim Position  As Integer

    Dim I         As Long

    Dim file_path As String

    Dim data()    As Byte

    Dim bmpInfo   As BITMAPINFO

    Dim handle    As Integer

    Dim bmpData   As StdPicture

    MouseX = X
    MouseY = Y
    
    SolapaView = User
    
    If InvBanco(1).SelectedItem = 0 Then Exit Sub
    
    If (Button = vbRightButton) Then
        If InvBanco(1).GrhIndex(InvBanco(1).SelectedItem) > 0 Then
            Last_I = InvBanco(1).SelectedItem

            If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                 
                Position = Search_GhID(3057)
                  
                If Position = 0 Then
                    I = GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum
                    Call Get_Image(DirGraficos & GRH_RESOURCE_FILE_DEFAULT, CStr(3057), data, False)
                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                    frmBancoObj.ImageList1.ListImages.Add , "g3057", Picture:=bmpData
                    Position = frmBancoObj.ImageList1.ListImages.Count
                    Set bmpData = Nothing
                End If
                 
                Set PicInv.MouseIcon = frmBancoObj.ImageList1.ListImages(Position).ExtractIcon
                frmBancoObj.PicInv.MousePointer = vbCustom
       
                Exit Sub

            End If
        End If
        
        
    'Else
       ' If MirandoObjetos Then
          '  FrmObject_Info.Top = Me.Top + Y
          ' FrmObject_Info.Left = Me.Left + X
        'End If
    End If

End Sub
Private Sub PicBancoInv_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    Dim Position  As Integer

    Dim I         As Long

    Dim file_path As String

    Dim data()    As Byte

    Dim bmpInfo   As BITMAPINFO

    Dim handle    As Integer

    Dim bmpData   As StdPicture
    
    MouseX = X
    MouseY = Y
    
    SolapaView = Npc
    
    If InvBanco(0).SelectedItem = 0 Then Exit Sub
    If (Button = vbRightButton) Then
        
        
        If InvBanco(0).GrhIndex(InvBanco(0).SelectedItem) > 0 Then
            Last_I = InvBanco(0).SelectedItem

            If Last_I > 0 And Last_I <= MAX_BANCOINVENTORY_SLOTS Then
                 
                Position = Search_GhID(3057)
                  
                If Position = 0 Then
                    I = GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum
                    Call Get_Image(DirGraficos & GRH_RESOURCE_FILE_DEFAULT, CStr(3057), data, False)
                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1)
                    frmBancoObj.ImageList1.ListImages.Add , "g3057", Picture:=bmpData
                    Position = frmBancoObj.ImageList1.ListImages.Count
                    Set bmpData = Nothing
                End If
                 
                Set PicBancoInv.MouseIcon = frmBancoObj.ImageList1.ListImages(Position).ExtractIcon
                frmBancoObj.PicBancoInv.MousePointer = vbCustom
       
                Exit Sub

            End If
        End If
       
    End If

End Sub
Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If InvBanco(1).SelectedItem = 0 Then Exit Sub
    
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (X > 0 And X < PicInv.ScaleWidth And Y > 0 And Y < PicInv.ScaleHeight) Then
        Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text, SelectedBank)
          
        InvBanco(1).sMoveItem = False
        InvBanco(1).uMoveItem = False
          
    End If
    PicInv.MousePointer = vbDefault

End Sub

Private Sub PicBancoInv_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If InvBanco(0).SelectedItem = 0 Then Exit Sub
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (X > 0 And X < PicBancoInv.ScaleWidth And Y > 0 And Y < PicBancoInv.ScaleHeight) Then
        Call WriteBankExtractItem(InvBanco(0).SelectedItem, Val(cantidad.Text), SelectedBank)
    End If

    PicBancoInv.MousePointer = vbDefault
    InvBanco(0).sMoveItem = False
    InvBanco(0).uMoveItem = False
       
End Sub

Private Sub tUpdate_Timer()
    If Not MirandoBanco Then Exit Sub
    Call InvBanco(0).DrawInventory
    Call InvBanco(1).DrawInventory
End Sub
