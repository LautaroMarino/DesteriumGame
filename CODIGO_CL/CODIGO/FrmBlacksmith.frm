VERSION 5.00
Begin VB.Form FrmBlacksmith 
   BorderStyle     =   0  'None
   Caption         =   "Yunque"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Yunque"
   Picture         =   "FrmBlacksmith.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   742
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicPremium 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5880
      Picture         =   "FrmBlacksmith.frx":3E48E
      ScaleHeight     =   375
      ScaleWidth      =   3930
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   3930
      Begin VB.Label lblPremium 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Comprar a 5.000 Fragmentos Premium"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   105
         Width           =   3495
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
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
      Height          =   5280
      Left            =   240
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   0
      Top             =   525
      Width           =   4320
   End
   Begin VB.ComboBox cmbShop 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      ItemData        =   "FrmBlacksmith.frx":43DA2
      Left            =   120
      List            =   "FrmBlacksmith.frx":43DA4
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   195
      Width           =   4335
   End
   Begin VB.Timer tHeading 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
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
      Height          =   5610
      Left            =   4560
      ScaleHeight     =   374
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   423
      TabIndex        =   1
      Top             =   210
      Width           =   6345
   End
   Begin VB.Image imgUnload 
      Height          =   585
      Left            =   10800
      Top             =   240
      Width           =   345
   End
End
Attribute VB_Name = "FrmBlacksmith"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private picRecuadroItem       As Picture
Private picRecuadroRecursos   As Picture
Private cPicConstruir(0 To 3) As clsGraphicalButton


Private IsPremium As Boolean
Private Selected As eItemsConstruibles_Subtipo

Private Sub cmbShop_Click()
    Dim A As Long
    Dim Slot As Long
    Dim ObjType As eItemsConstruibles_Subtipo
    Dim Index As Integer
    
    Index = cmbShop.ListIndex
    SlotObjIndex = 0
    ColourBody = -1

     'Inicializo los inventarios
    Call BlacksmithInv.Initialize(PicInv, 99, eCaption.InvBlacksmith, , , , , False, , , , True, True)
    
    ObjType = cmbShop.ListIndex
    
    If ObjType = eFundicion Then
        Call SeparateFundition
    Else
        Call SeparateObjtype(ObjType)
        
    End If
    
    Selected = ObjType
    
    MirandoHerreria = True
End Sub

Private Sub Form_Load()

    g_Captions(eCaption.InvBlacksmith) = wGL_Graphic.Create_Device_From_Display(PicInv.hWnd, PicInv.ScaleWidth, PicInv.ScaleHeight)
    g_Captions(eCaption.InvBlacksmithInfo) = wGL_Graphic.Create_Device_From_Display(FrmBlacksmith.PicInfo.hWnd, FrmBlacksmith.PicInfo.ScaleWidth, FrmBlacksmith.PicInfo.ScaleHeight)
    
    Set BlacksmithInv = New clsGrapchicalInventory
     
    SlotObjIndex = 0
    
    cmbShop.AddItem "Items que usa el personaje '" & UCase$(UserName) & "'"
    cmbShop.AddItem "Armadura"
    cmbShop.AddItem "Casco"
    cmbShop.AddItem "Escudo"
    cmbShop.AddItem "Armas"
    cmbShop.AddItem "Municiones"
    cmbShop.AddItem "Embarcaciones"
    cmbShop.AddItem "Objetos mágicos"
    cmbShop.AddItem "Instrumentos"
    cmbShop.AddItem "Fundición de Objetos"
    cmbShop.ListIndex = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set BlacksmithInv = Nothing
    
    MirandoHerreria = False
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.InvBlacksmith))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.InvBlacksmithInfo))
End Sub

Private Sub SeparateObjtype(ByRef ObjType As eItemsConstruibles_Subtipo)
     
     On Error GoTo ErrHandler
     
    Dim A As Long
    Dim Slot As Long
    
    ReDim ObjBlacksmith_Copy(0) As tItemsConstruibles
    
    For A = 1 To ObjBlacksmith_Amount
        If ObjType <> 0 Then
            If ObjBlacksmith(A).SubType = ObjType Then
                Slot = Slot + 1
                    
                ReDim Preserve ObjBlacksmith_Copy(Slot) As tItemsConstruibles
                ObjBlacksmith_Copy(Slot) = ObjBlacksmith(A)
            End If
        Else
            If ObjBlacksmith(A).CanUse Then
                Slot = Slot + 1
                    
                ReDim Preserve ObjBlacksmith_Copy(Slot) As tItemsConstruibles
                ObjBlacksmith_Copy(Slot) = ObjBlacksmith(A)
            End If
        End If
        
    Next A
    
    
    For A = 1 To ObjBlacksmith_Amount
        If A <= Slot Then
            Call BlacksmithInv.SetItem(A, ObjBlacksmith_Copy(A).ObjIndex, ObjBlacksmith_Copy(A).Amount, 0, ObjBlacksmith_Copy(A).GrhIndex, 0, 0, 0, 0, 0, 0, ObjBlacksmith_Copy(A).Name, 0, ObjBlacksmith_Copy(A).CanUse, 0, 0, 0, 0, ObjBlacksmith_Copy(A).Bronce, ObjBlacksmith_Copy(A).Plata, ObjBlacksmith_Copy(A).Oro, ObjBlacksmith_Copy(A).Premium)
        Else
            Call BlacksmithInv.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)
        End If
    Next A
    
    BlacksmithInv.DrawInventory
    Exit Sub
    
ErrHandler:
End Sub

Private Sub SeparateFundition()
     
     On Error GoTo ErrHandler
     
    Dim A As Long
    Dim Slot As Long
    
    ReDim ObjBlacksmith_Copy(0) As tItemsConstruibles
    
    For A = 1 To 30
        
        If Inventario.Amount(A) > 0 Then
            Slot = Slot + 1
                    
            ReDim Preserve ObjBlacksmith_Copy(Slot) As tItemsConstruibles
            
            ObjBlacksmith_Copy(Slot).Name = Inventario.ItemName(A)
            ObjBlacksmith_Copy(Slot).ObjIndex = Inventario.ObjIndex(A)
            ObjBlacksmith_Copy(Slot).Amount = Inventario.Amount(A)
            ObjBlacksmith_Copy(Slot).GrhIndex = Inventario.GrhIndex(A)
            ObjBlacksmith_Copy(Slot).SubType = eFundicion
            ObjBlacksmith_Copy(Slot).ObjType = Inventario.ObjType(A)
            
            Call BlacksmithInv.SetItem(Slot, Inventario.ObjIndex(A), Inventario.Amount(A), 0, Inventario.GrhIndex(A), 0, 0, 0, 0, 0, 0, Inventario.ItemName(A), 0, Inventario.CanUse(A), 0, 0, 0, 0, 0, 0, 0, 0)
        Else
            Call BlacksmithInv.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)
        End If
    Next A
    
    BlacksmithInv.DrawInventory
    Exit Sub
    
ErrHandler:
End Sub


Private Sub imgSelected_Click(Index As Integer)

End Sub

Private Sub imgUnload_Click()
    Unload Me
End Sub



Private Sub lblPremium_Click()
    If MsgBox("¿Desea construir el objeto '" & ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).Name & "'? ¡Se descontarán " & ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).RequiredPremium & " Fragmentos! ", vbYesNo) = vbYes Then
        Call WriteCraftBlacksmith(ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).ObjIndex, 1, 1)
    End If
End Sub

Private Sub picInv_Click()
    
    If BlacksmithInv.SelectedItem = 0 Then Exit Sub
    SlotObjIndex = BlacksmithInv.SelectedItem
    
    If Selected = eFundicion Then
        If ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).ObjType = otBarcos Then Exit Sub
        
        Dim Slot As Integer, A As Long, Temp As Long
        
        Slot = Blacksmith_SearchSlot(BlacksmithInv.ObjIndex(BlacksmithInv.SelectedItem))
        
        With ObjBlacksmith_Copy(BlacksmithInv.SelectedItem)
            .RequiredCant = ObjBlacksmith(Slot).RequiredCant
            
            If ObjBlacksmith(Slot).RequiredCant > 0 Then
                ReDim .Required(1 To ObjBlacksmith(Slot).RequiredCant) As tItemsConstruibles_Required
                
                For A = 1 To ObjBlacksmith(Slot).RequiredCant
                    Temp = Int(ObjBlacksmith(Slot).Required(A).Amount * 0.3)
                    
                    .Required(A).Amount = Temp
                    .Required(A).GrhIndex = ObjBlacksmith(Slot).Required(A).GrhIndex
                    .Required(A).Name = ObjBlacksmith(Slot).Required(A).Name
                    .Required(A).ObjIndex = ObjBlacksmith(Slot).Required(A).ObjIndex
                Next A
            End If
        End With
        
    Else
        Call WriteObj_RequiredInfo(BlacksmithInv.ObjIndex(BlacksmithInv.SelectedItem))
    End If
End Sub

Private Function Blacksmith_SearchSlot(ByVal ObjIndex As Integer) As Integer
    Dim A As Long
    
    For A = 1 To ObjBlacksmith_Amount
        If ObjBlacksmith(A).ObjIndex = ObjIndex Then
            Blacksmith_SearchSlot = A
            Exit Function
        End If
    Next A
End Function

Private Sub picInv_DblClick()
    
    If BlacksmithInv.SelectedItem = 0 Then
        Call MsgBox("Selecciona un objeto que desees adquirir.")
        Exit Sub
    End If
            
    If Selected = eFundicion Then
        If MsgBox("¿Desea fundir el objeto '" & ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).Name & "'? ¡Una vez fundido el objeto no hay vuelta atrás!", vbYesNo) = vbYes Then
            Call WriteCraftBlacksmith(ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).ObjIndex, 1, 2)
            Unload Me
        End If
        
        
    Else
        If MsgBox("¿Desea construir el objeto '" & ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).Name & "'?", vbYesNo) = vbYes Then
            Call WriteCraftBlacksmith(ObjBlacksmith_Copy(BlacksmithInv.SelectedItem).ObjIndex, 1, 0)
        End If
    End If
    
    

    
End Sub

Private Sub PicPremium_Click()
    lblPremium_Click
End Sub

Private Sub tHeading_Timer()
    
    If SlotObjIndex = 0 Then Exit Sub
    ObjBlacksmith_Copy(SlotObjIndex).Heading = ObjBlacksmith_Copy(SlotObjIndex).Heading + 1
    
    If ObjBlacksmith_Copy(SlotObjIndex).Heading > 4 Then
        ObjBlacksmith_Copy(SlotObjIndex).Heading = 1
    End If
End Sub



