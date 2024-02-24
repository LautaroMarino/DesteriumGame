VERSION 5.00
Begin VB.Form frmMercader_List 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Lista de Personajes a la Venta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMercader_List.frx":0000
   LinkTopic       =   "Lista de Personajes a la Venta"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHechiz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   6840
      MousePointer    =   99  'Custom
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   267
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.PictureBox Picture1 
      Height          =   225
      Index           =   1
      Left            =   11025
      ScaleHeight     =   165
      ScaleWidth      =   270
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3885
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Height          =   225
      Index           =   0
      Left            =   11025
      ScaleHeight     =   165
      ScaleWidth      =   270
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3675
      Width           =   330
   End
   Begin VB.TextBox txtMercader 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   1395
      Width           =   1500
   End
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   7320
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Timer tUpdate 
      Interval        =   100
      Left            =   3480
      Top             =   1200
   End
   Begin VB.Image imgMercaderRemove 
      Height          =   375
      Left            =   9000
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image imgMercaderOffer 
      Height          =   375
      Left            =   9000
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image imgRequired 
      Height          =   480
      Index           =   3
      Left            =   5760
      MouseIcon       =   "FrmMercader_List.frx":000C
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgOffer 
      Height          =   1215
      Left            =   10800
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   8
      Left            =   795
      Top             =   4380
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   7
      Left            =   795
      Top             =   4020
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   6
      Left            =   795
      Top             =   3630
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   5
      Left            =   795
      Top             =   3240
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   4
      Left            =   795
      Top             =   2880
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   3
      Left            =   795
      Top             =   2505
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   2
      Left            =   795
      Top             =   2130
      Width           =   10005
   End
   Begin VB.Image imgPublicationLarge 
      Height          =   360
      Index           =   1
      Left            =   795
      Top             =   1740
      Width           =   10005
   End
   Begin VB.Image imgRequired 
      Height          =   480
      Index           =   2
      Left            =   5250
      MouseIcon       =   "FrmMercader_List.frx":015E
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   480
   End
   Begin VB.Image imgRequired 
      Height          =   480
      Index           =   1
      Left            =   4725
      MouseIcon       =   "FrmMercader_List.frx":02B0
      MousePointer    =   99  'Custom
      Top             =   8025
      Width           =   480
   End
   Begin VB.Image imgRequired 
      Height          =   480
      Index           =   0
      Left            =   4185
      MouseIcon       =   "FrmMercader_List.frx":0402
      MousePointer    =   99  'Custom
      Top             =   8025
      Width           =   480
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   10
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":0554
      MousePointer    =   99  'Custom
      Top             =   7185
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   9
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":06A6
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   8
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":07F8
      MousePointer    =   99  'Custom
      Top             =   6810
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   7
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":094A
      MousePointer    =   99  'Custom
      Top             =   6630
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   6
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":0A9C
      Top             =   6450
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   5
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":0BEE
      MousePointer    =   99  'Custom
      Top             =   6270
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   4
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":0D40
      MousePointer    =   99  'Custom
      Top             =   6090
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   3
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":0E92
      MousePointer    =   99  'Custom
      Top             =   5895
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   2
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":0FE4
      MousePointer    =   99  'Custom
      Top             =   5715
      Width           =   2850
   End
   Begin VB.Image imgPJ 
      Height          =   195
      Index           =   1
      Left            =   735
      MouseIcon       =   "FrmMercader_List.frx":1136
      MousePointer    =   99  'Custom
      Top             =   5535
      Width           =   2850
   End
   Begin VB.Image imgPublication 
      Height          =   540
      Index           =   0
      Left            =   4680
      Top             =   630
      Width           =   2775
   End
   Begin VB.Image imgUnload 
      Height          =   420
      Left            =   11460
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "frmMercader_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MouseData As clsMouse
Attribute MouseData.VB_VarHelpID = -1
Private Sub Form_Load()
    g_Captions(eCaption.cMercaderList) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    g_Captions(eCaption.cMercaderInv) = wGL_Graphic.Create_Device_From_Display(Me.picInv.hWnd, Me.picInv.ScaleWidth, Me.picInv.ScaleHeight)

  '  Set MouseData = New clsMouse
  '  MouseData.Hook Me

    MercaderID = 0
    MercaderID_Selected = 1
    MercaderPJ = 0
    'imgPublicationLarge_Click (0)
    
    Call RenderMercaderList
    
    
    Set hlstMercader = New clsGraphicalList
    
    Call hlstMercader.Initialize(Me.picHechiz, RGB(200, 190, 190))
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cMercaderList))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cMercaderInv))
    
    'MouseData.Hook
End Sub

Private Sub imgMercaderOffer_Click()
    If Account.SaleSlot = 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    MercaderOff = 3
    picInv.visible = False
    picHechiz.visible = False
    txtMercader.visible = False
    Prepare_And_Connect E_MODO.e_LoginMercaderOff
End Sub

Private Sub imgMercaderRemove_Click()
    If Account.SaleSlot = 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    MercaderOff = 0
    Prepare_And_Connect E_MODO.e_LoginMercaderOff
End Sub

Private Sub imgOffer_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Mercader_ModoOferta Then
        If MsgBox("¿Estás seguro que deseas aceptar la oferta? ¡Una vez aceptada no hay vuelta atrás!", vbYesNo) = vbYes Then
            If MercaderID > 0 Then
                Call WriteMercader_Required(5, MercaderID, 0)
            Else
                Call MsgBox("¡Selecciona la Oferta que deseas aceptar!")
                Exit Sub
            End If
        End If
    Else
        If MercaderPJ > 0 Then
            MercaderSelected = ePanelOffer
            
        Else
            MercaderSelected = ePanelPublication
        End If
        
        
        #If FullScreen = 0 Then
            frmMercader.Show vbModeless, FrmMain
        #Else
            Call MsgBox("¡Lo sentimos! Solo esta disponible desde la resolucion de 800x600")
        #End If
    End If
    
    Unload Me
End Sub

Private Sub imgPublication_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    MercaderID = 0
    MercaderID_Selected = 1
    MercaderPJ = 0
    MercaderOff = 0
    Call WriteMercader_Required(1, 1, 255)
    picInv.visible = False
End Sub

Private Sub MouseData_MouseWheel(ByVal wKeys As Long, ByVal zDelta As Long, ByVal xPos As Long, ByVal yPos As Long)
    ' Cuando se mueve la rueda del ratón
    'Call MsgBox("Se ha movido la rueda del ratón pos arriba +" & wKeys & ", " & zDelta)
    
    If zDelta <= 240 Then
       ' lblUP_Click (0)
    Else
        'lblUP_Click (1)
    End If
End Sub


Private Sub imgPJ_Click(Index As Integer)
    If MercaderID <= 0 Then Exit Sub
    If Index > MercaderList_Copy(MercaderID).Char Then Exit Sub
    
    
    MercaderPJ = Index
    Call Audio.PlayInterface(SND_CLICK)
    
    Call WriteMercader_Required(IIf(Mercader_ModoOferta, 4, 2), MercaderList_Copy(MercaderID).ID, MercaderPJ)
    imgRequired_Click 0

End Sub
Public Sub imgPublicationLarge_Click(Index As Integer)

    If Index = 0 Then MercaderID = 0
    
    If Not Mercader_ModoOferta Then
        If MercaderList(Index + MercaderID_Selected - 1).Char = 0 Then Exit Sub
    End If
    
    Call Audio.PlayInterface(SND_CLICK)
    
    MercaderID = MercaderList(Index + MercaderID_Selected - 1).ID
    
    picInv.visible = False
    
    If MercaderID > 0 Then
        
        MercaderPJ = 1
        imgPJ_Click (1)
        
    Else
        MercaderPJ = 0
    End If
    
    
End Sub

Public Sub SelectedButtonChar()
    
    imgRequired_Click (1)
End Sub
Private Sub imgRequired_Click(Index As Integer)
    
    If MercaderID <= 0 Then Exit Sub
    If MercaderPJ <= 0 Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)
    
    MercaderPJ_Selected = Index + 1
    
    Dim A As Long
    Dim ObjIndex As Integer
    Dim ObjAmount As Integer
    
    Select Case Index
    
        Case 0 ' Inventario
            picInv.visible = True
             picHechiz.visible = False
            Call Inventory_Draw
        Case 1 ' Banco
            picInv.visible = True
             picHechiz.visible = False
            Call Bank_Draw
        Case 2 ' Hechizos
            picInv.visible = False
            picHechiz.visible = True
            Call Spells_Draw
        Case 3 ' Skills
            picInv.visible = False
            picHechiz.visible = True
            Skills_Draw
    End Select
    
    InvMercader.DrawInventory
End Sub

Private Sub Spells_Draw()
    Dim A As Long
    
    hlstMercader.Clear
    
    For A = 1 To MAXHECHI
        If Len(MercaderList_Copy(MercaderID).Chars(MercaderPJ).Spells(A)) > 0 Then
            hlstMercader.AddItem MercaderList_Copy(MercaderID).Chars(MercaderPJ).Spells(A)
        End If
    Next A
    
End Sub

Private Sub Skills_Draw()
    Dim A As Long
    
    hlstMercader.Clear
    
    For A = 1 To NUMSKILLS
        If MercaderList_Copy(MercaderID).Chars(MercaderPJ).Skills(A) > 0 Then
            hlstMercader.AddItem SkillsNames(A) & ": " & MercaderList_Copy(MercaderID).Chars(MercaderPJ).Skills(A)
        End If
    Next A
    
End Sub
Public Sub Inventory_Draw()

    Dim C As Long
    
    If InvMercader Is Nothing Then
        Set InvMercader = New clsGrapchicalInventory
        Call InvMercader.Initialize(frmMercader_List.picInv, 40, 40, cMercaderInv, , , , , , , , , , True)
    
    End If
    
    For C = 1 To MAX_BANCOINVENTORY_SLOTS

        If C <= MAX_INVENTORY_SLOTS Then

            With MercaderList_Copy(MercaderID).Chars(MercaderPJ)

                If .Object(C).ObjIndex > 0 Then
                    Call InvMercader.SetItem(C, .Object(C).ObjIndex, .Object(C).Amount, 0, ObjData(.Object(C).ObjIndex).GrhIndex, 0, 0, 0, 0, 0, 0, ObjData(.Object(C).ObjIndex).Name, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)
                Else
                    Call InvMercader.SetItem(C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)

                End If

            End With

        Else
            Call InvMercader.SetItem(C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)

        End If

    Next C

    InvMercader.DrawInventory
    
End Sub

Private Sub Bank_Draw()
    Dim C As Long
     
    For C = 1 To MAX_BANCOINVENTORY_SLOTS
        With MercaderList_Copy(MercaderID).Chars(MercaderPJ)
            If .Bank(C).ObjIndex > 0 Then
                Call InvMercader.SetItem(C, .Bank(C).ObjIndex, .Bank(C).Amount, 0, ObjData(.Bank(C).ObjIndex).GrhIndex, 0, 0, 0, 0, 0, 0, ObjData(.Bank(C).ObjIndex).Name, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)
            Else
                Call InvMercader.SetItem(C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0, 0, 0, 0, 0)
            End If
        End With
    Next C
    
    InvMercader.DrawInventory
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    MercaderSelected = ePanelInitial
    Mercader_ModoOferta = False
    If FrmConnect_Account.visible Then
        
        frmMercader.Show , FrmConnect_Account
    Else
        frmMercader.Show , FrmMain
    End If
    
    Unload Me
End Sub

Private Sub Picture1_Click(Index As Integer)

    Dim A        As Long

    Dim TempSlot As Long
    
    Select Case Index
    
        Case 0 ' Arriba

            If MercaderID_Selected <= 1 Then Exit Sub
            MercaderID_Selected = MercaderID_Selected - 1

        Case 1 ' Abajo

            If MercaderID_Selected < MERCADER_MAX_LIST Then
                MercaderID_Selected = MercaderID_Selected + 1

            End If

    End Select
    
    If Not MercaderList_Copy(MercaderID_Selected).Loaded Then
        Call Protocol.WriteMercader_Required(1, MercaderID_Selected, MercaderID_Selected)
        MercaderLoaded = True

    End If
    
     If MercaderList(MercaderID_Selected).Char = 0 Then
        MercaderPJ = 0
        picInv.visible = False
    End If
End Sub

Private Sub tUpdate_Timer()
    Call RenderMercaderList
End Sub

Private Sub MouseData_Activate()
    ' Cuando se activa
    'Label2(2).Caption = "Activate"
End Sub

Private Sub MouseData_Deactivate()
    ' Cuando se desactiva
    '(2).Caption = "Deactivate"
End Sub

Private Sub MouseData_DisplayChange(ByVal BitsPerPixel As Long, ByVal cxScreen As Long, ByVal cyScreen As Long)
    ' DisplayChange
End Sub

Private Sub MouseData_FontChange()
    ' FontChange
End Sub

Private Sub MouseData_LowMemory()
    ' LowMemory
End Sub

Private Sub MouseData_MenuSelected(ByVal mnuItem As Long, mnuFlags As eWSCMF, ByVal hMenu As Long)
    ' MenuSelected
    'List1.AddItem "Has seleccionado el menú: " & mnuItem
End Sub

Private Sub MouseData_MouseEnter()
    ' Cuando entra el ratón en el formulario
   ' Label1 = "MouseEnter"
End Sub

Private Sub MouseData_MouseEnterOn(unControl As Object)
    ' Cuando entra el ratón en uno de los controles
   ' Label1.Caption = "MouseEnterOn: " & unControl.Name
End Sub

Private Sub MouseData_MouseLeave()
    ' Cuando sale el ratón
   ' Label1 = "MouseLeave"
End Sub

Private Sub MouseData_MouseLeaveOn(unControl As Object)
    ' Cuando sale el ratón de un control
   ' Label1.Caption = "MouseLeaveOn: " & unControl.Name
End Sub

Private Sub MouseData_Move(ByVal wLeft As Long, ByVal wTop As Long)
    ' Cuando se mueve el formulario
    'Label2(0).Caption = "Posición del formulario: Left= " & wLeft & ", Top= " & wTop
End Sub

Private Sub MouseData_SetCursor(unControl As Object, ByVal HitTest As eWSCHitTest, ByVal MouseMsg As Long)
    'Label2(2).Caption = "SetCursor: " & unControl.Name & ", HitTest:" & HitTest & ", MouseMsg: " & MouseMsg
End Sub

Private Sub MouseData_WindowPosChanged(ByVal wLeft As Long, ByVal wTop As Long, ByVal wWidth As Long, ByVal wHeight As Long)
    ' Cuando cambia la posición de la ventana
   ' Label2(1).Caption = "WndPosChanged: L:" & wLeft & ", T:" & wTop & ", W:" & wWidth & ", H:" & wHeight
End Sub


Private Sub txtMercader_Change()
    
    MercaderID_Selected = 1
    Dim A As Long
    
    If Len(txtMercader.Text) <= 0 Then
        For A = 1 To MERCADER_MAX_LIST
            MercaderList(A) = MercaderList_Copy(A)
        Next A
        
        'Call Mercader_OrdenClass
    Else
    
        Call Filter_MAO(txtMercader.Text)
    End If
End Sub

Private Sub Filter_MAO(ByRef sCompare As String)

    Dim lIndex As Long, b As Long, C As Long
    Dim MaoNull As tMercader
    Dim Slot As Long
    Dim A As Long
    
      For A = 1 To MERCADER_MAX_LIST
            MercaderList(A) = MaoNull
        Next A
        
    If UBound(MercaderList_Copy) <> 0 Then
        For lIndex = 1 To UBound(MercaderList_Copy)
            For b = 1 To ACCOUNT_MAX_CHARS
                If StrComp(UCase$(MercaderList_Copy(lIndex).Chars(b).Name), UCase$(sCompare)) = 0 Then
                    Slot = Slot + 1
                    MercaderList(Slot) = MercaderList_Copy(lIndex)
                End If
            Next b
        Next lIndex

    End If


End Sub

' Lista Gráfica de Hechizos
Private Sub picHechiz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picHechiz.ScaleHeight / hlstMercader.Pixel_Alto) * hlstMercader.Pixel_Alto - 1 Then Y = Int(picHechiz.ScaleHeight / hlstMercader.Pixel_Alto) * hlstMercader.Pixel_Alto - 1
If X < picHechiz.ScaleWidth - 10 Then
    hlstMercader.ListIndex = Int(Y / hlstMercader.Pixel_Alto) + hlstMercader.Scroll
    hlstMercader.DownBarrita = 0

Else
    hlstMercader.DownBarrita = Y - hlstMercader.Scroll * (picHechiz.ScaleHeight - hlstMercader.BarraHeight) / (hlstMercader.ListCount - hlstMercader.VisibleCount)
End If
End Sub

Private Sub picHechiz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picHechiz.ScaleHeight / hlstMercader.Pixel_Alto) * hlstMercader.Pixel_Alto - 1 Then yy = Int(picHechiz.ScaleHeight / hlstMercader.Pixel_Alto) * hlstMercader.Pixel_Alto - 1
    If hlstMercader.DownBarrita > 0 Then
        hlstMercader.Scroll = (Y - hlstMercader.DownBarrita) * (hlstMercader.ListCount - hlstMercader.VisibleCount) / (picHechiz.ScaleHeight - hlstMercader.BarraHeight)
    Else
        hlstMercader.ListIndex = Int(yy / hlstMercader.Pixel_Alto) + hlstMercader.Scroll

        'If ScrollArrastrar = 0 Then
            'If (Y < yy) Then hlstMercader.Scroll = hlstMercader.Scroll - 1
           ' If (Y > yy) Then hlstMercader.Scroll = hlstMercader.Scroll + 1
        'End If
    End If
ElseIf Button = 0 Then
    hlstMercader.ShowBarrita = X > picHechiz.ScaleWidth - hlstMercader.BarraWidth * 2
End If
End Sub

Private Sub picHechiz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
hlstMercader.DownBarrita = 0
End Sub

