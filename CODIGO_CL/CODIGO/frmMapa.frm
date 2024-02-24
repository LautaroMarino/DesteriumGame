VERSION 5.00
Begin VB.Form FrmMapa 
   BorderStyle     =   0  'None
   Caption         =   "Mapa"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicNpcs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   3075
      MousePointer    =   99  'Custom
      Picture         =   "frmMapa.frx":24FBF
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   1965
      Begin VB.ComboBox cmbMaps 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   765
         Width           =   1725
      End
      Begin VB.TextBox txtSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   3
         Top             =   1275
         Width           =   1095
      End
      Begin VB.PictureBox PicList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         DrawStyle       =   3  'Dash-Dot
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2700
         Index           =   0
         Left            =   120
         MousePointer    =   99  'Custom
         ScaleHeight     =   180
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1740
      End
      Begin VB.Image chkOrdenOro 
         Height          =   225
         Left            =   150
         Top             =   5160
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image chkOrdenExp 
         Height          =   225
         Left            =   150
         Top             =   4905
         Width           =   210
      End
      Begin VB.Image chkOrdenName 
         Height          =   225
         Left            =   150
         Top             =   4650
         Width           =   210
      End
      Begin VB.Image chkStats 
         Height          =   225
         Left            =   150
         Top             =   4395
         Width           =   210
      End
      Begin VB.Image chkMisiones 
         Height          =   225
         Left            =   120
         Top             =   120
         Width           =   210
      End
   End
   Begin VB.Timer TimerDraw 
      Interval        =   100
      Left            =   600
      Top             =   360
   End
   Begin VB.PictureBox PicMapa 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   5970
      Left            =   375
      ScaleHeight     =   398
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   1020
      Width           =   2595
   End
   Begin VB.PictureBox PicObjs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   3060
      MousePointer    =   99  'Custom
      Picture         =   "frmMapa.frx":2A7B7
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   1965
      Begin VB.PictureBox PicList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         DrawStyle       =   3  'Dash-Dot
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   4140
         Index           =   1
         Left            =   120
         MousePointer    =   99  'Custom
         ScaleHeight     =   276
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1740
      End
      Begin VB.TextBox txtSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   5
         Top             =   885
         Width           =   1095
      End
      Begin VB.Image chkCofres 
         Height          =   225
         Left            =   120
         Top             =   600
         Width           =   210
      End
      Begin VB.Image chkDrops 
         Height          =   225
         Left            =   120
         Top             =   360
         Width           =   210
      End
      Begin VB.Image chkStatsObj 
         Height          =   225
         Left            =   120
         Top             =   120
         Width           =   210
      End
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "frmMapa.frx":2D240
      Top             =   0
      Width           =   330
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   21
      Left            =   45
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   20
      Left            =   945
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   19
      Left            =   1770
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   16
      Left            =   0
      Top             =   885
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   17
      Left            =   900
      Top             =   885
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   18
      Left            =   1725
      Top             =   885
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   7
      Left            =   0
      Top             =   1770
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   6
      Left            =   900
      Top             =   1770
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   5
      Left            =   1725
      Top             =   1770
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   8
      Left            =   0
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   1
      Left            =   900
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   4
      Left            =   1725
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   9
      Left            =   0
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   2
      Left            =   900
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   3
      Left            =   1725
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   10
      Left            =   0
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   11
      Left            =   900
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   12
      Left            =   1725
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   15
      Left            =   0
      Top             =   5160
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   14
      Left            =   900
      Top             =   5160
      Width           =   855
   End
   Begin VB.Image Map 
      Height          =   855
      Index           =   13
      Left            =   1725
      Top             =   5160
      Width           =   855
   End
   Begin VB.Image chkPanel 
      Height          =   225
      Left            =   3075
      Top             =   6690
      Width           =   210
   End
   Begin VB.Image chkNumber 
      Height          =   225
      Left            =   465
      Top             =   7035
      Width           =   210
   End
   Begin VB.Image chkCuadrilla 
      Height          =   225
      Left            =   1815
      Top             =   7035
      Width           =   210
   End
   Begin VB.Image ButtonObj 
      Height          =   405
      Left            =   4080
      Picture         =   "frmMapa.frx":2E2F2
      Top             =   1020
      Width           =   975
   End
   Begin VB.Image ButtonNpc 
      Height          =   405
      Left            =   3060
      Picture         =   "frmMapa.frx":2F745
      Top             =   1020
      Width           =   975
   End
End
Attribute VB_Name = "FrmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const LAST_RENDER_MAP As Byte = 21

Private clsFormulario        As clsFormMovementManager
Private ListMapa(1)             As clsGraphicalList
Private picCheckBox          As Picture
Private picCheckBoxNulo      As Picture


Public VerNumeros As Boolean
Public VerCuadrilla As Boolean
Public VerPanel As Boolean
Public VerMisiones As Boolean
Public VerStats As Boolean

Public MouseX As Integer
Public MouseY As Integer

Private ListCopy() As Integer    ' Copia de Npcs-Cofres-Objs
Private ListMap(LAST_RENDER_MAP) As Integer    ' Lista de Mapas donde la Criatura-Cofre-Objeto está
Private ListSelected As Integer ' Index seleccionado de Npcs-Cofre-Objs

Private LastList As Integer


Public NpcIndexSelected As Integer

Private OrdenView As Byte

Private Sub ButtonObj_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Set ButtonNpc.Picture = Nothing
    Set ButtonObj.Picture = LoadPicture(App.path & "\resource\interface\menucompacto\ButtonObjs_Selected.jpg")
    
    SelectedPanel (1)
End Sub

Private Sub ButtonNpc_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Set ButtonNpc.Picture = LoadPicture(App.path & "\resource\interface\menucompacto\ButtonNpcs_Selected.jpg")
    Set ButtonObj.Picture = Nothing
    
    SelectedPanel (0)
End Sub


Private Sub SelectedPanel(ByVal Index As Byte)

    Select Case Index
        Case 0 ' Npcs
            PicNpcs.visible = True
            PicObjs.visible = False
            
         
        Case 1 ' Objs
            PicNpcs.visible = False
            PicObjs.visible = True
    End Select
End Sub


Private Sub chkMisiones_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    VerMisiones = Not VerMisiones
    Set chkMisiones.Picture = IIf(VerMisiones, picCheckBox, picCheckBoxNulo)
End Sub

Private Sub chkOrdenExp_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call Orden_Exp(1)
    
    OrdenView = 1
    Set chkOrdenName.Picture = picCheckBoxNulo
    Set chkOrdenOro.Picture = picCheckBoxNulo
    Set chkOrdenExp.Picture = picCheckBox
End Sub

Private Sub chkOrdenName_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call Orden_Exp(3)
    
    OrdenView = 3
    Set chkOrdenName.Picture = picCheckBox
    Set chkOrdenOro.Picture = picCheckBoxNulo
    Set chkOrdenExp.Picture = picCheckBoxNulo
     
End Sub

Private Sub chkOrdenOro_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call Orden_Exp(2)
    
    OrdenView = 2
    Set chkOrdenName.Picture = picCheckBoxNulo
    Set chkOrdenOro.Picture = picCheckBox
    Set chkOrdenExp.Picture = picCheckBoxNulo
End Sub

Private Sub chkStats_Click()
     Call Audio.PlayInterface(SND_CLICK)
    
    VerStats = Not VerStats
    Set chkStats.Picture = IIf(VerStats, picCheckBox, picCheckBoxNulo)
    
    
    UpdateStatsNpc
End Sub

Private Sub UpdateStatsNpc()
    If VerStats And NpcIndexSelected Then
        If MirandoNpc Then
            FrmMapaNpc.UpdateNpc (NpcIndexSelected)
        Else
            FrmMapaNpc.Show , FrmMain
            FrmMapaNpc.UpdateNpc (NpcIndexSelected)
        End If
    End If
End Sub

Private Sub cmbMaps_Click()
    
    If cmbMaps.ListIndex = -1 Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Call ListarNpcs_Map(cmbMaps.ListIndex)
    
End Sub

Private Sub Form_Load()

    g_Captions(eCaption.eMapa) = wGL_Graphic.Create_Device_From_Display(PicMapa.hWnd, PicMapa.ScaleWidth, PicMapa.ScaleHeight)
    
    Me.Picture = LoadPicture(App.path & "\resource\interface\menucompacto\VentanaMapa.jpg")
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    Set ListMapa(0) = New clsGraphicalList
    Set ListMapa(1) = New clsGraphicalList
    
    Call ListMapa(0).Initialize(PicList(0), RGB(200, 190, 190))
    Call ListMapa(1).Initialize(PicList(1), RGB(200, 190, 190))

    Dim GrhPath As String
    
    GrhPath = DirInterface
                        
    Set picCheckBox = LoadPicture(DirInterface & "options\CheckBoxOpciones.jpg")
    Set picCheckBoxNulo = LoadPicture(DirInterface & "options\CheckBoxOpcionesNulo.jpg")
    
    ListSelected = 0
    
    Call ButtonNpc_Click
    
    ' # Cargamos el defecto
    Call ListarNpcs
    
    ' # Cargamos el Chk Nombre
    Call chkOrdenName_Click
    
End Sub







' # Botones


Private Sub chkCuadrilla_Click()
    Call Audio.PlayInterface(SND_CLICK)
    VerCuadrilla = Not VerCuadrilla
    Set chkCuadrilla.Picture = IIf(VerCuadrilla, picCheckBox, picCheckBoxNulo)
    
End Sub


Private Sub chkNumber_Click()
    Call Audio.PlayInterface(SND_CLICK)
    VerNumeros = Not VerNumeros
    Set chkNumber.Picture = IIf(VerNumeros, picCheckBox, picCheckBoxNulo)
End Sub

Private Sub chkPanel_Click()

    Call Audio.PlayInterface(SND_CLICK)
End Sub



' # Funciones Internas de proceso

' # Listamos todos los Npcs
Public Sub ListarNpcs()
    
    Dim A As Long

    ListMapa(0).Clear
    cmbMaps.Clear
    
    Dim NpcIndex As Integer
    
    ReDim ListCopy(1 To NpcsGlobal_Last) As Integer
    
    For A = 1 To NpcsGlobal_Last
        NpcIndex = NpcsGlobal(A)
        ListCopy(A) = NpcIndex
        
        ListMapa(0).AddItem NpcList(NpcIndex).Name
        
    Next A
    
    
    cmbMaps.AddItem "TODOS"
    
    For A = 1 To LAST_RENDER_MAP
         cmbMaps.AddItem MiniMap(A).Name
    Next A
   
   Call Check_Orden
   Call Render
End Sub

' # Listar Npcs del Mapa seleccionado
Public Sub ListarNpcs_Map(ByVal Map As Integer)
    
    If Map = 0 Then
        ListarNpcs
        Exit Sub
        
    End If
    
    
    ListMapa(0).Clear
    
    Dim A As Long
    Dim NpcIndex As Integer
    ReDim ListCopy(1 To NpcsGlobal_Last) As Integer
    
    
    With MiniMap(Map)
        For A = 1 To .NpcsNum
            NpcIndex = .Npcs(A).NpcIndex
            ListCopy(A) = NpcIndex
            
            ListMapa(0).AddItem NpcList(NpcIndex).Name
        Next A
        
    End With
    
    Call Check_Orden
    Call PicList_Click(0)
    Call Render
End Sub

' # Ordena los Npcs listados según (EXP)
Public Sub Orden_Exp(ByVal Tipo As Byte)

    Dim A    As Long, b As Long
    Dim Temp As Integer
    
    Dim cantidad As Integer
    cantidad = ListMapa(0).ListCount
    
    ListMapa(0).Clear
    
    Dim Value(1) As Long
    
    For A = 1 To cantidad - 1
        For b = 1 To cantidad - A
                Select Case Tipo
                    Case 1 ' Exp
                        Value(0) = NpcList(ListCopy(b)).GiveExp
                        Value(1) = NpcList(ListCopy(b + 1)).GiveExp
                        
                    Case 2 ' Oro
                        Value(0) = NpcList(ListCopy(b)).GiveGld
                        Value(1) = NpcList(ListCopy(b + 1)).GiveGld
                        
                    Case 3 ' Nombre
                        ' Usamos StrComp para comparar las cadenas (Nombre)
                        Value(0) = StrComp(NpcList(ListCopy(b)).Name, NpcList(ListCopy(b + 1)).Name, vbTextCompare)
                        If Value(0) > 0 Then
                            ' Si el valor devuelto es mayor que 0, intercambiamos las posiciones
                            Temp = ListCopy(b)
                            ListCopy(b) = ListCopy(b + 1)
                            ListCopy(b + 1) = Temp
                        End If
                        
                End Select
                
                If Tipo <> 3 Then
                    If Value(0) < Value(1) Then
                        Temp = ListCopy(b)
                        ListCopy(b) = ListCopy(b + 1)
                        ListCopy(b + 1) = Temp
                        
                    End If
                End If
        Next b
    Next A
    
    For A = 1 To UBound(ListCopy)
        If ListCopy(A) > 0 Then
            ListMapa(0).AddItem NpcList(ListCopy(A)).Name
        End If
    Next A
End Sub

' # Chequeo de Orden (Experiencia, Oro, Nombre)
Private Sub Check_Orden()
    Call Orden_Exp(OrdenView)
End Sub


' # Reinicia el mapita de info
Private Sub ResetListMapa()
    Dim A As Long
    
    For A = 1 To LAST_RENDER_MAP
        ListMap(A) = 0
    Next A

End Sub

Private Sub Image1_Click()

End Sub

Private Sub PicList_Click(Index As Integer)
    If ListMapa(Index).ListIndex = -1 Then Exit Sub
    
    
    
    
    ShowConsoleMsg "Npc: " & NpcList(ListCopy(ListMapa(0).ListIndex + 1)).Name

    ' Actualizo la lista de mapas en los que está la criatura
    ResetListMapa
    
    Dim NpcIndex As Integer, A As Long, b As Long
    
    
    NpcIndexSelected = ListCopy(ListMapa(0).ListIndex + 1)
    
    If MirandoNpc Then
        FrmMapaNpc.UpdateNpc (NpcIndexSelected)
    End If
    
    For A = 1 To LAST_RENDER_MAP
        If MiniMap(A).NpcsNum > 0 Then
            For b = 1 To MiniMap(A).NpcsNum
                NpcIndex = MiniMap(A).Npcs(b).NpcIndex
                
                If NpcIndex = NpcIndexSelected Then
                    ListMap(A) = A
                End If
            Next b
        End If
    Next
    
    UpdateStatsNpc
    Render
End Sub

' # FIN Funciones Internas de proceso

Private Sub txtSearch_Change(Index As Integer)
    Dim A As Long
    
    If Len(txtSearch(Index).Text) <= 0 Then
        ListarNpcs
    Else
        Call FiltrarNpcs(txtSearch(Index).Text)
    End If
End Sub

' # Filtra los npcs por BUSCADOR
Public Sub FiltrarNpcs(ByRef sCompare As String)

    Dim lIndex As Long, b As Long
    Dim NpcIndex As Integer
    
    ListMapa(0).Clear
    
    ReDim ListCopy(0) As Integer
    
    For lIndex = 1 To NpcsGlobal_Last
        NpcIndex = NpcsGlobal(lIndex)
        
        If InStr(1, UCase$(NpcList(NpcIndex).Name), UCase$(sCompare)) Then
            ListMapa(0).AddItem NpcList(NpcIndex).Name
            
            ReDim Preserve ListCopy(LBound(ListCopy) To UBound(ListCopy) + 1)
            
            ListCopy(UBound(ListCopy)) = NpcIndex
        End If
    Next lIndex
    
End Sub


' # Renderizado del mapa
Private Sub Render()
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.eMapa))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, PicMapa.ScaleWidth, PicMapa.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    
    Call Draw_Texture_Graphic_Gui(144, 0, 0, To_Depth(1), 173, 398, 0, 0, 173, 398, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    
    
    ' Personaje
    
    Dim X As Long
    Dim Y As Long
    Dim GrhIndex As Long
    Dim NpcIndex As Integer
    Dim A As Long
    
   ' If VerPersonaje Then
        
        X = MapConfig(UserMap).RenderX
        Y = MapConfig(UserMap).RenderY
        
        With CharList(UserCharIndex)
            If .iBody > 0 Then
                GrhIndex = BodyData(.iBody).Walk(.Heading).GrhIndex
                Call Draw_Grh(BodyData(.iBody).Walk(.Heading), X, Y, To_Depth(6), 1, 1, 0, , , eTechnique.t_Alpha, _
                       GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight)
            End If
            
            If .iHead > 0 Then
                GrhIndex = HeadData(.iHead).Head(.Heading).GrhIndex
                Call Draw_Grh(HeadData(.iHead).Head(.Heading), X + BodyData(.iBody).HeadOffset.X, Y + BodyData(.iBody).HeadOffset.Y, To_Depth(6), 1, 1, , , , , _
                GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight)
            End If
        End With
    ' End If
    
    
    Dim Color As Long
    
    
    
    For A = 1 To LAST_RENDER_MAP
        X = MapConfig(A).RenderX
        Y = MapConfig(A).RenderY
                    
        If ListMap(A) > 0 Then
            If ListMapa(0).ListIndex <> -1 Then
            NpcIndex = ListCopy(ListMapa(0).ListIndex + 1)
                    
            
            With NpcList(NpcIndex)
                If .Body > 0 Then
                    GrhIndex = BodyData(.Body).Walk(E_Heading.SOUTH).GrhIndex
                    
                    
                    Dim Porc As Single
                    
                    If GrhData(GrhIndex).pixelWidth >= 80 Or GrhData(GrhIndex).pixelHeight >= 60 Then
                        Porc = 0.8
                    Else
                        Porc = 1
                    End If
                    
                    Call Draw_Grh(BodyData(.Body).Walk(E_Heading.SOUTH), X, Y, To_Depth(6), 1, 1, 0, , , eTechnique.t_Alpha, _
                                GrhData(GrhIndex).pixelWidth * Porc, GrhData(GrhIndex).pixelHeight * Porc)
                End If
                            
                If .Head > 0 Then
                    GrhIndex = HeadData(.Head).Head(E_Heading.SOUTH).GrhIndex
                    Call Draw_Grh(HeadData(.Head).Head(E_Heading.SOUTH), X + BodyData(.Body).HeadOffset.X, Y + BodyData(.Body).HeadOffset.Y, To_Depth(6), 1, 1, , , , , _
                        GrhData(GrhIndex).pixelWidth * Porc, GrhData(GrhIndex).pixelHeight * Porc)
                End If
                            
            End With
                    
            End If
            
            Color = ARGB(255, 147, 0, 255)
        Else
            Color = ARGB(255, 255, 255, 100)
        End If
            
        If VerNumeros Then
            If VerCuadrilla Then
                Draw_Text f_Verdana, 17, X - 20, Y - 20, To_Depth(9), 0#, Color, FONT_ALIGNMENT_CENTER, A, True, True
            Else
                Draw_Text f_Morpheus, 50, X, Y, To_Depth(9), 0#, Color, FONT_ALIGNMENT_CENTER, A, True, True
            End If
            
            '
            
        End If
        
    Next A
                
    ' Cuadrilla
    If VerCuadrilla Then Call Draw_Texture_Graphic_Gui(143, 0, 0, To_Depth(2), 173, 398, 0, 0, 173, 398, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
     
    
  '  Draw_Text f_Tahoma, 14, 10, 10, To_Depth(9), 0#, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_LEFT, "X: " & MouseX, True, True
  '  Draw_Text f_Tahoma, 14, 10, 30, To_Depth(9), 0#, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_LEFT, "Y: " & MouseY, True, True
    
    Call wGL_Graphic_Renderer.Flush
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub

Private Sub Map_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
End Sub

Private Sub Map_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub PicMapa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub TimerDraw_Timer()
    Render
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMapa))
End Sub


' # Listas graficas

Private Sub picList_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(PicList(Index).ScaleHeight / ListMapa(Index).Pixel_Alto) * ListMapa(Index).Pixel_Alto - 1 Then Y = Int(PicList(Index).ScaleHeight / ListMapa(Index).Pixel_Alto) * ListMapa(Index).Pixel_Alto - 1

If X < PicList(Index).ScaleWidth - 10 Then
    ListMapa(Index).ListIndex = Int(Y / ListMapa(Index).Pixel_Alto) + ListMapa(Index).Scroll
    ListMapa(Index).DownBarrita = 0

Else
    ListMapa(Index).DownBarrita = Y - ListMapa(Index).Scroll * (PicList(Index).ScaleHeight - ListMapa(Index).BarraHeight) / (ListMapa(Index).ListCount - ListMapa(Index).VisibleCount)
End If
End Sub

Private Sub picList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then


    Static TimeArrastre As Long
    
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(PicList(Index).ScaleHeight / ListMapa(Index).Pixel_Alto) * ListMapa(Index).Pixel_Alto - 1 Then yy = Int(PicList(Index).ScaleHeight / ListMapa(Index).Pixel_Alto) * ListMapa(Index).Pixel_Alto - 1
    If ListMapa(Index).DownBarrita > 0 Then
        ListMapa(Index).Scroll = (Y - ListMapa(Index).DownBarrita) * (ListMapa(Index).ListCount - ListMapa(Index).VisibleCount) / (PicList(Index).ScaleHeight - ListMapa(Index).BarraHeight)
    Else
        ListMapa(Index).ListIndex = Int(yy / ListMapa(Index).Pixel_Alto) + ListMapa(Index).Scroll
        
        If ListMapa(Index).ListIndex <> LastList Then
            LastList = ListMapa(Index).ListIndex
            PicList_Click (Index)
        End If
        
        If ScrollArrastrar = 0 Then
            If (GetSystemTime - TimeArrastre) >= 150 Then
                TimeArrastre = GetSystemTime
                If (Y < yy) Then ListMapa(Index).Scroll = ListMapa(Index).Scroll - 1
                If (Y > yy) Then ListMapa(Index).Scroll = ListMapa(Index).Scroll + 1
            End If
        End If
    End If
ElseIf Button = 0 Then
    ListMapa(Index).ShowBarrita = X > PicList(Index).ScaleWidth - ListMapa(Index).BarraWidth * 2
End If
End Sub

Private Sub picList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ListMapa(Index).DownBarrita = 0
End Sub


