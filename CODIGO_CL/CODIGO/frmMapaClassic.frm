VERSION 5.00
Begin VB.Form frmMapaClassic 
   BorderStyle     =   0  'None
   Caption         =   "Mundo Desterium"
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
   Icon            =   "frmMapaClassic.frx":0000
   LinkTopic       =   "Mundo Desterium"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCofres 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1860
      Left            =   7080
      MousePointer    =   99  'Custom
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5760
      Width           =   4815
   End
   Begin VB.PictureBox picCofreItem 
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
      Height          =   1050
      Left            =   7125
      MousePointer    =   4  'Icon
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   315
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7680
      Width           =   4725
   End
   Begin VB.PictureBox picInv 
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
      Height          =   1050
      Left            =   3900
      MousePointer    =   4  'Icon
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.PictureBox picNpcs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2100
      Left            =   9585
      MousePointer    =   99  'Custom
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2295
   End
   Begin VB.PictureBox picMaps 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2100
      Left            =   7080
      MousePointer    =   99  'Custom
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2295
   End
   Begin VB.Timer tUpdate 
      Interval        =   10
      Left            =   4440
      Top             =   840
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   92
      Left            =   2400
      Top             =   7800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   91
      Left            =   2400
      Top             =   7200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   86
      Left            =   2880
      Top             =   6600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   85
      Left            =   3480
      Top             =   6600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   84
      Left            =   4080
      Top             =   6600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   83
      Left            =   4080
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   87
      Left            =   3480
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   82
      Left            =   4680
      Top             =   7200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   81
      Left            =   4680
      Top             =   6600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   80
      Left            =   4680
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   94
      Left            =   3600
      Top             =   8400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   93
      Left            =   3000
      Top             =   8400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   90
      Left            =   3000
      Top             =   7800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   89
      Left            =   3000
      Top             =   7200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   88
      Left            =   3480
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   74
      Left            =   4680
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   75
      Left            =   4080
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   76
      Left            =   3480
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   77
      Left            =   2880
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   78
      Left            =   2400
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   79
      Left            =   1800
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   99
      Left            =   6360
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   98
      Left            =   5760
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   97
      Left            =   5160
      Top             =   120
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   69
      Left            =   5880
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   68
      Left            =   5280
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   60
      Left            =   5280
      Top             =   3600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   67
      Left            =   5880
      Top             =   3600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   66
      Left            =   5880
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   65
      Left            =   5880
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   64
      Left            =   5280
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   63
      Left            =   4680
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   62
      Left            =   4560
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   61
      Left            =   4560
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   59
      Left            =   4560
      Top             =   3600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   58
      Left            =   4560
      Top             =   3000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   57
      Left            =   4560
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   56
      Left            =   3960
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   55
      Left            =   3480
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   47
      Left            =   2880
      Top             =   720
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   46
      Left            =   2880
      Top             =   1200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   53
      Left            =   2880
      Top             =   1800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   54
      Left            =   2880
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   52
      Left            =   2400
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   38
      Left            =   4680
      Top             =   7800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   37
      Left            =   5280
      Top             =   7800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   36
      Left            =   5280
      Top             =   8400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   35
      Left            =   4680
      Top             =   8400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   34
      Left            =   1200
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   32
      Left            =   4200
      Top             =   8400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   31
      Left            =   1800
      Top             =   600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   30
      Left            =   1800
      Top             =   1200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   29
      Left            =   1800
      Top             =   1800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   28
      Left            =   1800
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   27
      Left            =   1200
      Top             =   1800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   26
      Left            =   1200
      Top             =   2400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   25
      Left            =   1200
      Top             =   3000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   24
      Left            =   1200
      Top             =   3600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   23
      Left            =   1200
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   22
      Left            =   720
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   21
      Left            =   720
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   20
      Left            =   1200
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   19
      Left            =   2400
      Top             =   6600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   18
      Left            =   3000
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   17
      Left            =   4080
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   16
      Left            =   4080
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   15
      Left            =   3600
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   14
      Left            =   3000
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   13
      Left            =   2400
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   12
      Left            =   3000
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   11
      Left            =   2400
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   10
      Left            =   120
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   9
      Left            =   720
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   8
      Left            =   1200
      Top             =   4800
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   7
      Left            =   1800
      Top             =   3000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   6
      Left            =   1800
      Top             =   3600
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   5
      Left            =   1800
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   4
      Left            =   2400
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   3
      Left            =   1800
      Top             =   6000
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   2
      Left            =   1800
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image imgMap 
      Height          =   570
      Index           =   1
      Left            =   1800
      Top             =   4800
      Width           =   555
   End
End
Attribute VB_Name = "frmMapaClassic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Public MouseX As Long
Public MouseY As Long

Public ListMaps As clsGraphicalList
Public ListNpcs As clsGraphicalList
Public ListCofres As clsGraphicalList

Public InvMapa As clsGrapchicalInventory
Private InvCofre As clsGrapchicalInventory

Private NpcIndex_Selected As Integer
Private MapaSelected As Integer
Private MapaSubSelected As Integer
Private MapaCofreSelected As Integer
Private Map_Coord As String


Private Sub Form_Click()
    ResetAll
End Sub

'Private clsFormulario As clsFormMovementManager

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
    g_Captions(eCaption.eMapaClassic) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    g_Captions(eCaption.cInvMapa) = wGL_Graphic.Create_Device_From_Display(picInv.hWnd, picInv.ScaleWidth, picInv.ScaleHeight)
    g_Captions(eCaption.cInvCofre1) = wGL_Graphic.Create_Device_From_Display(picCofreItem.hWnd, picCofreItem.ScaleWidth, picCofreItem.ScaleHeight)
  
    picInv.visible = False
  
    ' Listas Gráficas
    Set ListMaps = New clsGraphicalList
    Set ListNpcs = New clsGraphicalList
    Set ListCofres = New clsGraphicalList
    
    Set InvMapa = New clsGrapchicalInventory
    Set InvCofre = New clsGrapchicalInventory
    
    Call ListMaps.Initialize(picMaps, RGB(200, 190, 190), 14, 30)
    Call ListNpcs.Initialize(picNpcs, RGB(200, 190, 190), 14, 30)
    Call ListCofres.Initialize(picCofres, RGB(200, 190, 190), 14, 30)
    
    Call InvMapa.Initialize(picInv, 10, 10, eCaption.cInvMapa, , , , , , , , False, , True, , True)
    Call InvCofre.Initialize(picCofreItem, 18, 18, eCaption.cInvCofre1, , , , , , , , False, , True, , True)
    
    MapaSelected = UserMap
    MapaSubSelected = MapaSelected
    imgMap_Click MapaSelected

End Sub

Private Sub ResetAll()
    NpcIndex_Selected = 0
    picInv.visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
   'If MirandoObjetos Then
      ' FrmObject_Info.Close_Form
    'End If
     
     
     MirandoListaDrops = False
     MirandoListaCofres = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

      MirandoListaDrops = False
      MirandoListaCofres = False
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMapaClassic))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvMapa))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvCofre1))
    
   ' If MirandoObjetos Then
      '  FrmObject_Info.Close_Form

   'End If

End Sub

Private Sub UpdateListMaps()

    Dim A As Long
        
        
    ListMaps.Clear
    
    With MiniMap(MapaSelected)
        ListMaps.AddItem .Name

        If .SubMaps > 0 Then

            For A = 1 To .SubMaps
                ListMaps.AddItem MiniMap(.Maps(A)).Name
            Next A
        
        End If
    
    End With

End Sub

Private Sub UpdateListNpcs()

    Dim A As Long
    Dim Npcs As Integer
    
    picInv.visible = False
    ListNpcs.Clear
    
    With MiniMap(MapaSubSelected)

        If .NpcsNum > 0 Then
            
            For A = 1 To .NpcsNum
                'If NpcList(.Npcs(A).NpcIndex).MaxHp > 0 Then
                    Npcs = Npcs + 1
                    ListNpcs.AddItem NpcList(.Npcs(A).NpcIndex).Name
               ' End If
            Next A

        End If
    
        
        If Npcs > 0 Then
            NpcIndex_Selected = MiniMap(MapaSubSelected).Npcs(1).NpcIndex
        Else
            NpcIndex_Selected = 0
            picInv.visible = False
            Exit Sub
        End If
        
     If NpcList(NpcIndex_Selected).NroItems + NpcList(NpcIndex_Selected).NroDrops = 0 Then Exit Sub
     Call RenderScreen_DrawInventory
    End With

End Sub

Private Sub imgMap_Click(Index As Integer)

    Dim A As Long
    
    MapaSelected = Index
    MapaSubSelected = MapaSelected
    
    Call UpdateListMaps
    Call UpdateListNpcs
    Call RenderScreen_Chest(1)
End Sub

Private Sub imgMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Index = 0 Then Exit Sub
    Map_Coord = MiniMap(Index).Name & " - " & "Mapa " & Index
End Sub


Private Sub picCofreItem_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    MouseX = X
    MouseY = Y
    MirandoListaCofres = True

End Sub

Private Sub picCofres_Click()
    If ListCofres.ListIndex = -1 Then Exit Sub
    
    MapaCofreSelected = ListCofres.ListIndex + 1
    
    Chest_Reset
    Render_Chest_Drop MapaCofreSelected
End Sub

Private Sub picCofres_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
    MirandoListaDrops = True
End Sub

Private Sub picMaps_Click()

    If ListMaps.ListIndex = -1 Then Exit Sub
    
    If ListMaps.ListIndex = 0 Then
        MapaSubSelected = MapaSelected
    Else
        MapaSubSelected = MiniMap(MapaSelected).Maps(ListMaps.ListIndex)
    End If
    
    Call UpdateListNpcs
    Call RenderScreen_Chest(1)
End Sub

Private Sub picMaps_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub picNpcs_Click()

    If ListNpcs.ListIndex = -1 Then Exit Sub
    
    If ListNpcs.ListIndex = 0 Then
        NpcIndex_Selected = MiniMap(MapaSubSelected).Npcs(ListNpcs.ListIndex + 1).NpcIndex
    Else
        NpcIndex_Selected = MiniMap(MapaSubSelected).Npcs(ListNpcs.ListIndex + 1).NpcIndex
    End If
    
    
    picInv.visible = True
     'If NpcList(NpcIndex_Selected).NroItems + NpcList(NpcIndex_Selected).NroDrops = 0 Then Exit Sub
     Call RenderScreen_DrawInventory
    
    
End Sub

Private Sub Inventory_Reset()
      
    Dim A As Long
    
    For A = 1 To 10
        Call InvMapa.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
    Next A

End Sub
Private Sub Chest_Reset()
      
    Dim A As Long, b As Long
    
    For A = 1 To 15
        Call InvCofre.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
    Next A

End Sub

Private Sub Render_Chest_Drop(ByVal Index As Integer)

    Dim A As Long, b As Long
    Dim Drop As Long, Slot As Long
    
    With MiniMap(MapaSubSelected)
    
        ' Muestro los Drops del primer Cofre
        For A = 1 To ObjData(.Chest(Index)).Chest.NroDrop
            Drop = ObjData(.Chest(Index)).Chest.Drop(A)
                    
            With DropData(Drop)
    
                For b = 1 To .Last
                    Slot = Slot + 1
                    Call InvCofre.SetItem(Slot, .data(b).ObjIndex, 1, 0, ObjData(.data(b).ObjIndex).GrhIndex, ObjData(.data(b).ObjIndex).ObjType, 0, 0, 0, 0, 0, ObjData(.data(b).ObjIndex).Name & " (" & .data(b).Prob & "%)" & " [" & .data(b).Amount(0) & "-" & .data(b).Amount(1) & "]", 0, True, 0, 0, 0, 0)
                Next b
            End With
                    
        Next A
    End With
    
End Sub

Private Sub RenderScreen_Chest(ByVal Index As Integer)
    Call Chest_Reset
    
    Dim A        As Long, b As Long

    Dim Drop     As Integer

    Dim ObjIndex As Integer

    Dim Drops    As Byte

    Dim Slot     As Integer
    
    ListCofres.Clear
    
    With MiniMap(MapaSubSelected)
        
        If .ChestLast > 0 Then
            
            For A = 1 To .ChestLast
                ListCofres.AddItem ObjData(.Chest(A)).Name
                'Call InvCofre(0).SetItem(A, .Chest(A), 1, 0, ObjData(.Chest(A)).GrhIndex, ObjData(.Chest(A)).ObjType, 0, 0, 0, 0, 0, ObjData(.Chest(A)).Name, 0, True, 0, 0, 0, 0)
            Next A
             
             Call Render_Chest_Drop(Index)

        End If

    End With
    
    InvCofre.DrawInventory

End Sub

Private Sub RenderScreen_DrawInventory()
    
    Dim Slot As Byte

    Dim A    As Long
    
    Call Inventory_Reset
    
    With NpcList(NpcIndex_Selected)
        
        
        If .NroItems > 0 Then
            
            For A = 1 To .NroItems
                Slot = Slot + 1
                
                Call InvMapa.SetItem(Slot, .Object(A).ObjIndex, .Object(A).Amount, 0, ObjData(.Object(A).ObjIndex).GrhIndex, ObjData(.Object(A).ObjIndex).ObjType, 0, 0, 0, 0, 0, .Object(A).Name, 0, True, 0, 0, 0, 0)

            Next A

        End If
        
        If .NroDrops > 0 Then

            For A = 1 To .NroDrops
                Slot = Slot + 1
                
                Call InvMapa.SetItem(Slot, .Drop(A).ObjIndex, .Drop(A).Amount, 0, ObjData(.Drop(A).ObjIndex).GrhIndex, ObjData(.Drop(A).ObjIndex).ObjType, 0, 0, 0, 0, 0, .Drop(A).Name, 0, True, 0, 0, 0, 0)

            Next A
            
        End If
        
        InvMapa.DrawInventory

    End With

End Sub

Private Sub picNpcs_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub tUpdate_Timer()
    RenderScreen_GranMapa
    InvMapa.DrawInventory
    InvCofre.DrawInventory
End Sub


Public Sub RenderScreen_GranMapa()

    '<EhHeader>
    On Error GoTo RenderScreen_Err
    '</EhHeader>
        
    Dim FondoPNG As Long
    
    FondoPNG = 136
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.eMapaClassic))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, Me.ScaleWidth, Me.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
   
    Dim X          As Long
    Dim Y          As Long

    Dim Drawable   As Long
    Dim DrawableX  As Long
    Dim DrawableY  As Long

    Dim Divisor    As Long
    
    Dim XTemp As Long
    Dim YTemp As Long
    
    XTemp = 0
    YTemp = 0
    
    Call Draw_Texture_Graphic_Gui(FondoPNG, 0, 0, To_Depth(1), 800, 600, 0, 0, 800, 600, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
    
    
    X = 10
    Y = 575
    
    Call Draw_Text(eFonts.f_Verdana, 14, X, Y, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, Map_Coord, False, True)
   ' Call Draw_Text(eFonts.f_Verdana, 14, MouseX, MouseY, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, Map_Coord, False, True)
        
    
    Select Case PanelMapa
    
        Case ePanelMapa.eDefault
            
            If MapaSelected Then
                ' Config Basic
                Call RenderScreen_ConfigBasic
                
                ' List Npcs
               ' If MiniMap(MapaSelected).NpcsNum > 0 Then
                    'RenderScreen_ConfigNpcs
               ' End If
                
            End If
            
            If NpcIndex_Selected Then
                Call RenderScreen_InfoNpc
            
            
            End If
            
        Case ePanelMapa.eMapInfo
            
    End Select
    
    Call wGL_Graphic_Renderer.Flush

    Exit Sub

RenderScreen_Err:
    LogError err.Description & vbCrLf & _
       "in RenderScreen_GranMapa " & _
       "at line " & Erl

End Sub

Private Sub RenderScreen_InfoNpc()

    Dim X        As Long, Y As Long, Y_Avance As Long

    Dim A        As Long

    Dim FondoPNG As Long

    Dim Body     As Long

    Dim Head     As Long
    
    With MiniMap(MapaSubSelected)
            
        X = 240
        Y = 120
        FondoPNG = 137
        
        Call Draw_Texture_Graphic_Gui(FondoPNG, X, Y, To_Depth(3), 213, 263, 0, 0, 213, 263, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
            
        X = 325
        Y = 300
        Body = NpcList(NpcIndex_Selected).Body
        Head = NpcList(NpcIndex_Selected).Head
        
        If BodyData(Body).Walk(E_Heading.SOUTH).GrhIndex > 0 Then
            Call Draw_Grh(BodyData(Body).Walk(E_Heading.SOUTH), X, Y, To_Depth(5), 1, 0, 0)
                    
            If HeadData(Head).Head(E_Heading.SOUTH).GrhIndex > 0 Then
                Call Draw_Grh(HeadData(Head).Head(E_Heading.SOUTH), X + BodyData(Body).HeadOffset.X, Y + BodyData(Body).HeadOffset.Y, To_Depth(5), 1, 0, 0)

            End If
                
        End If
        
        X = 345
        Y = 110
        

        With NpcList(NpcIndex_Selected)
            Call Draw_Text(eFonts.f_Medieval, 25, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 170, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Name, False, True)
            
            
            X = 345
            Y = 200
        
            If .MinHit > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Hit " & .MinHit & "/" & .MaxHit, True, True)
                Y_Avance = Y_Avance + 15

            End If
            
            If .Def > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Def " & .Def, True, True)
                Y_Avance = Y_Avance + 15

            End If

            If .DefM > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Def Mag " & .DefM, True, True)
                Y_Avance = Y_Avance + 15

            End If

            If .MaxHp > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Hp " & PonerPuntos(.MaxHp), True, True)
                Y_Avance = Y_Avance + 15

            End If
            
            If .GiveExp > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "EXP " & PonerPuntos(.GiveExp), True, True)
                Y_Avance = Y_Avance + 15

            End If
        
            If .GiveGld > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "ORO " & PonerPuntos(.GiveGld), True, True)
                Y_Avance = Y_Avance + 15

            End If
        
            If .RespawnTime > 0 Then
                Call Draw_Text(eFonts.f_Verdana, 16, X, Y + Y_Avance, To_Depth(9), 0, ARGB(255, 0, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Respawn " & SecondsToHMS(.RespawnTime), True, True)
                Y_Avance = Y_Avance + 15
            
            End If
        End With
        
    End With

End Sub

Private Sub RenderScreen_ConfigNpcs()

    Dim X As Long, Y As Long, Y_Avance As Long

    Dim A As Long
        
    With MiniMap(MapaSubSelected)

        If .NpcsNum = 0 Then Exit Sub
            
        X = 620
        Y = 65
        
        For A = 1 To .NpcsNum
            Call Draw_Text(eFonts.f_Verdana, 6, X, Y + Y_Avance, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, NpcList(.Npcs(A).NpcIndex).Name, False)
            Y_Avance = Y_Avance + 15
        Next A
        
    End With

    End Sub
Private Sub RenderScreen_SubMaps()

    Dim X As Long, Y As Long, Y_Avance As Long

    Dim A As Long
        
    With MiniMap(MapaSelected)

        If .SubMaps = 0 Then Exit Sub
        X = 535
        Y = 75
        
        For A = 1 To .SubMaps
            Call Draw_Text(eFonts.f_Verdana, 6, X, Y + Y_Avance, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, MiniMap(.Maps(A)).Name, False)
            Y_Avance = Y_Avance + 15
        Next A
        
    End With

    End Sub

Private Sub RenderScreen_ConfigBasic()

    Dim X As Long, Y As Long, Y_Avance As Long

    Dim A As Long
    
    With MiniMap(MapaSubSelected)
    
        ' Mapa y Sub Mapas
        '  X = 535
        ' Y = 60
        'Call Draw_Text(eFonts.f_Verdana, 7, X, Y, To_Depth(2), 0, ARGB(50, 200, 35, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Name, False)
        'Call RenderScreen_SubMaps
        
        X = 720
        Y = 228
        
        If .Pk Then
            Call Draw_Text(eFonts.f_Verdana, 17, X, Y, To_Depth(2), 0, ARGB(200, 12, 12, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "ZONA INSEGURA", False, True)
        Else
            Call Draw_Text(eFonts.f_Verdana, 17, X, Y, To_Depth(2), 0, ARGB(200, 240, 190, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "ZONA SEGURA", False, True)

        End If
        
        X = 482
        Y = 228
        Call Draw_Text(eFonts.f_Verdana, 17, X, Y, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Nivel Permitido:", False, True)
        Call Draw_Text(eFonts.f_Verdana, 17, X + 115, Y, To_Depth(2), 0, ARGB(255, 210, 10, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, .LvlMin & "-" & .LvlMax, False, True)
      
        X = 480
        Y = 250
        Call Draw_Texture_Graphic_Gui(99, X, Y, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
        Call Draw_Texture_Graphic_Gui(99, X, Y + 20, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
        Call Draw_Texture_Graphic_Gui(99, X, Y + 40, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
        Call Draw_Texture_Graphic_Gui(99, X, Y + 60, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
        Call Draw_Texture_Graphic_Gui(99, X, Y + 80, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)

        Call Draw_Text(eFonts.f_Verdana, 13, X + 25, Y + 2, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "VALE RESU", False, True)
        Call Draw_Text(eFonts.f_Verdana, 13, X + 25, Y + 22, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "VALE INVOCAR", False, True)
        Call Draw_Text(eFonts.f_Verdana, 13, X + 25, Y + 42, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "VALE OCULTAR", False, True)
        Call Draw_Text(eFonts.f_Verdana, 13, X + 25, Y + 62, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "VALE INVI", False, True)
        Call Draw_Text(eFonts.f_Verdana, 13, X + 25, Y + 82, To_Depth(2), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "CAEN ITEMS", False, True)
        
        If .ResuSinEfecto = 0 Then
            Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 4, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)

        End If
        
        If .InvocarSinEfecto = 0 Then
            Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 24, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)

        End If
        
        If .OcultarSinEfecto = 0 Then
            Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 44, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)

        End If
        
        If .InviSinEfecto = 0 Then
            Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 64, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)

        End If
        
        If .CaenItem = 1 Then
            Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 84, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)

        End If
        
    End With
    
End Sub










'############################
' Lista Gráfica de Hechizos
Private Sub picMaps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picMaps.ScaleHeight / ListMaps.Pixel_Alto) * ListMaps.Pixel_Alto - 1 Then Y = Int(picMaps.ScaleHeight / ListMaps.Pixel_Alto) * ListMaps.Pixel_Alto - 1

If X < picMaps.ScaleWidth - 10 Then
    ListMaps.ListIndex = Int(Y / ListMaps.Pixel_Alto) + ListMaps.Scroll
    ListMaps.DownBarrita = 0

Else
    ListMaps.DownBarrita = Y - ListMaps.Scroll * (picMaps.ScaleHeight - ListMaps.BarraHeight) / (ListMaps.ListCount - ListMaps.VisibleCount)
End If
End Sub

Private Sub picMaps_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picMaps.ScaleHeight / ListMaps.Pixel_Alto) * ListMaps.Pixel_Alto - 1 Then yy = Int(picMaps.ScaleHeight / ListMaps.Pixel_Alto) * ListMaps.Pixel_Alto - 1
    If ListMaps.DownBarrita > 0 Then
        ListMaps.Scroll = (Y - ListMaps.DownBarrita) * (ListMaps.ListCount - ListMaps.VisibleCount) / (picMaps.ScaleHeight - ListMaps.BarraHeight)
    Else
        ListMaps.ListIndex = Int(yy / ListMaps.Pixel_Alto) + ListMaps.Scroll

       ' If ScrollArrastrar = 0 Then
           ' If (Y < yy) Then ListMaps.Scroll = ListMaps.Scroll - 1
           ' If (Y > yy) Then ListMaps.Scroll = ListMaps.Scroll + 1
        'End If
    End If
ElseIf Button = 0 Then
    ListMaps.ShowBarrita = X > picMaps.ScaleWidth - ListMaps.BarraWidth * 2
End If
End Sub

Private Sub picMaps_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListMaps.DownBarrita = 0
End Sub

'############################
' Lista Gráfica de Hechizos
Private Sub picNpcs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picNpcs.ScaleHeight / ListNpcs.Pixel_Alto) * ListNpcs.Pixel_Alto - 1 Then Y = Int(picNpcs.ScaleHeight / ListNpcs.Pixel_Alto) * ListNpcs.Pixel_Alto - 1

If X < picNpcs.ScaleWidth - 10 Then
    ListNpcs.ListIndex = Int(Y / ListNpcs.Pixel_Alto) + ListNpcs.Scroll
    ListNpcs.DownBarrita = 0

Else
    ListNpcs.DownBarrita = Y - ListNpcs.Scroll * (picNpcs.ScaleHeight - ListNpcs.BarraHeight) / (ListNpcs.ListCount - ListNpcs.VisibleCount)
End If
End Sub

Private Sub picNpcs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picNpcs.ScaleHeight / ListNpcs.Pixel_Alto) * ListNpcs.Pixel_Alto - 1 Then yy = Int(picNpcs.ScaleHeight / ListNpcs.Pixel_Alto) * ListNpcs.Pixel_Alto - 1
    If ListNpcs.DownBarrita > 0 Then
        ListNpcs.Scroll = (Y - ListNpcs.DownBarrita) * (ListNpcs.ListCount - ListNpcs.VisibleCount) / (picNpcs.ScaleHeight - ListNpcs.BarraHeight)
    Else
        ListNpcs.ListIndex = Int(yy / ListNpcs.Pixel_Alto) + ListNpcs.Scroll

       ' If ScrollArrastrar = 0 Then
           ' If (Y < yy) Then ListNpcs.Scroll = ListNpcs.Scroll - 1
           ' If (Y > yy) Then ListNpcs.Scroll = ListNpcs.Scroll + 1
        'End If
    End If
ElseIf Button = 0 Then
    ListNpcs.ShowBarrita = X > picNpcs.ScaleWidth - ListNpcs.BarraWidth * 2
End If
End Sub

Private Sub picNpcs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListNpcs.DownBarrita = 0
End Sub
'############################
' Lista Gráfica de Hechizos
Private Sub picCofres_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picCofres.ScaleHeight / ListCofres.Pixel_Alto) * ListCofres.Pixel_Alto - 1 Then Y = Int(picCofres.ScaleHeight / ListCofres.Pixel_Alto) * ListCofres.Pixel_Alto - 1

If X < picCofres.ScaleWidth - 10 Then
    ListCofres.ListIndex = Int(Y / ListCofres.Pixel_Alto) + ListCofres.Scroll
    ListCofres.DownBarrita = 0

Else
    ListCofres.DownBarrita = Y - ListCofres.Scroll * (picCofres.ScaleHeight - ListCofres.BarraHeight) / (ListCofres.ListCount - ListCofres.VisibleCount)
End If
End Sub

Private Sub picCofres_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picCofres.ScaleHeight / ListCofres.Pixel_Alto) * ListCofres.Pixel_Alto - 1 Then yy = Int(picCofres.ScaleHeight / ListCofres.Pixel_Alto) * ListCofres.Pixel_Alto - 1
    If ListCofres.DownBarrita > 0 Then
        ListCofres.Scroll = (Y - ListCofres.DownBarrita) * (ListCofres.ListCount - ListCofres.VisibleCount) / (picCofres.ScaleHeight - ListCofres.BarraHeight)
    Else
        ListCofres.ListIndex = Int(yy / ListCofres.Pixel_Alto) + ListCofres.Scroll

       ' If ScrollArrastrar = 0 Then
           ' If (Y < yy) Then ListNpcs.Scroll = ListNpcs.Scroll - 1
           ' If (Y > yy) Then ListNpcs.Scroll = ListNpcs.Scroll + 1
        'End If
    End If
ElseIf Button = 0 Then
    ListCofres.ShowBarrita = X > picCofres.ScaleWidth - ListCofres.BarraWidth * 2
End If
End Sub

Private Sub picCofres_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ListCofres.DownBarrita = 0
End Sub
