VERSION 5.00
Begin VB.Form frmCriatura_Quest 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCriature_Quest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tSecond 
      Interval        =   1050
      Left            =   5880
      Top             =   480
   End
   Begin VB.Timer tUpdate 
      Interval        =   200
      Left            =   6360
      Top             =   480
   End
   Begin VB.Image imgFabricar 
      Height          =   375
      Left            =   2160
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgCommerce 
      Height          =   375
      Left            =   240
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Realiza doble clic sobre el objeto y podrás saber su información"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   555
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   6090
   End
   Begin VB.Image imgNpcInfo 
      Height          =   1485
      Left            =   7200
      Top             =   2400
      Width           =   2220
   End
   Begin VB.Image NextQuest 
      Height          =   480
      Left            =   7320
      Top             =   4800
      Width           =   240
   End
   Begin VB.Image LastQuest 
      Height          =   480
      Left            =   7320
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image NextNpc 
      Height          =   270
      Left            =   7440
      Top             =   5880
      Width           =   270
   End
   Begin VB.Image LastNpc 
      Height          =   270
      Left            =   7200
      Top             =   5880
      Width           =   270
   End
   Begin VB.Image Npc 
      Height          =   480
      Left            =   7200
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image ItemRequired 
      Height          =   480
      Index           =   4
      Left            =   6510
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image ItemRequired 
      Height          =   480
      Index           =   3
      Left            =   6045
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image ItemRequired 
      Height          =   480
      Index           =   2
      Left            =   5550
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image ItemRequired 
      Height          =   480
      Index           =   1
      Left            =   5070
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image ItemRequired 
      Height          =   480
      Index           =   0
      Left            =   4605
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   1
      Left            =   1200
      Top             =   2400
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   0
      Left            =   1200
      Top             =   1920
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   2
      Left            =   1200
      Top             =   2880
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   3
      Left            =   1200
      Top             =   3360
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   4
      Left            =   1200
      Top             =   3840
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   5
      Left            =   1200
      Top             =   4320
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   6
      Left            =   1200
      Top             =   4800
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   7
      Left            =   1200
      Top             =   5280
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   8
      Left            =   1200
      Top             =   5760
      Width           =   3315
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   9
      Left            =   1200
      Top             =   6240
      Width           =   3315
   End
   Begin VB.Image ImgUnload 
      Height          =   435
      Left            =   7320
      MouseIcon       =   "frmCriature_Quest.frx":000C
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "frmCriatura_Quest"
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

' Determina el Panel seleccionado en el que esta posicionado
Public Enum ePanelCommerce
    eList = 0 '  Predeterminada con Lista de Objetos
    eRequiredObj = 1 ' Objetos que requiere para comprar y/o construir el obj
    eRequiredNpc = 2 ' Panel de Criaturas requeridas para completar la misión
    eRequiredObj_Selected = 3 ' Panel del objeto seleccionado que requiere (Ej un Lingote)

End Enum

Public PanelCommerce As ePanelCommerce
        
Private Heading As Byte
Private clsFormulario As clsFormMovementManager

Public LastIndex1     As Integer

Public LastIndex2     As Integer

Public LasActionBuy   As Boolean

Private ClickNpcInv   As Boolean

Private lIndex        As Byte


Private Sub Form_Load()
    
    Dim A As Long
    
    ' Handles Form movement (drag and drop).
    'Set clsFormulario = New clsFormMovementManager
    'clsFormulario.Initialize Me

    g_Captions(eCaption.cPivQuest) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    
    QuestIndex = 0
    QuestNpcIndex = 0
    QuestObjIndex = 0
    PanelCommerce = eList
    
    For A = 1 To Item.UBound
        Item(A).Top = 128 + (A * 32)
        Item(A).Left = Item(1).Left
        
    Next A
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
  '  frmMain.SetFocus
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cPivQuest))
    
    
   ' If MirandoObjetos Then
        'FrmObject_Info.Close_Form
    'End If
    
End Sub

Private Sub imgCommerce_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
    If InitQuest Then
        InitQuest = False
        frmComerciar.visible = True
        Unload Me
    Else
        Call MsgBox("La criatura no comercia objetos")
    End If
End Sub

Private Sub imgFabricar_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If QuestIndex <= 0 Then
        Call MsgBox("Selecciona el objeto a construir...")
        Exit Sub
    End If
    
     If Crafting_Checking_Object(QuestNpc(QuestIndex)) Then
        WriteCraftBlacksmith QuestNpc(QuestIndex)
    End If
End Sub

Private Sub imgNpcInfo_Click()
    If QuestIndex <= 0 Then Exit Sub
    If QuestNpcIndex <= 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    SelectedNpcIndex = QuestList(QuestNpc(QuestIndex)).Npcs(QuestNpcIndex).NpcIndex
    
    Call Invalidate(FrmCriatura_Info.hWnd)
    If Not FrmCriatura_Info.visible Then
        Call FrmCriatura_Info.Show(, frmCriatura_Quest)
    End If
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If InitQuest Then
        InitQuest = False
        MirandoComerciar = False
        Call WriteCommerceEnd
        FrmMain.SetFocus
        Unload Me
        
    Else
        Unload Me
    End If
    
End Sub

Private Sub Item_Click(Index As Integer)
    If (Index + 1) > QuestLast Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)

    QuestIndex = Index + 1
    QuestNpcIndex = 0
    QuestObjIndex = 0
    
    Render_QuestPanel
    
    Dim A As Long
    
    For A = ItemRequired.LBound To ItemRequired.UBound
        ItemRequired(A).Top = Item(Index).Top
        ItemRequired(A).Left = 238 + (A * 32)
        'ItemRequired(A).Left = (A * 32) + (Item(Index).Left + 32)
    Next A
            
    Npc.Top = Item(Index).Top
    LastQuest.Top = Item(Index).Top
    NextQuest.Top = Item(Index).Top
        
    LastQuest.Left = (Npc.Left) - 32
    NextQuest.Left = (Npc.Left) - 16
    PanelCommerce = ePanelCommerce.eRequiredObj
    
    
End Sub

' # Comprueba de tener los recursos necesarios que necesita el objeto para ser creado/mejorado
Public Function Crafting_Checking_Object(ByVal QuestIndex As Integer) As Boolean
    Dim A As Long
    Dim Temp As String
    
    Crafting_Checking_Object = True
    
    With QuestList(QuestIndex)
        For A = 1 To .Obj
            If Not TieneObjetos(.Objs(A).ObjIndex) >= .Objs(A).Amount Then
                Temp = Temp & "* " & ObjData(.Objs(A).ObjIndex).Name & " (x" & .Objs(A).Amount & ")" & vbCrLf
                Crafting_Checking_Object = False
            End If
        Next A
        
        If Not Crafting_Checking_Object Then
            Call ShowConsoleMsg("Te faltan recursos: " & vbCrLf & Temp)
        End If
        
    End With
    
    
End Function

Private Sub Item_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QuestIndex <= 0 Then Exit Sub
    If QuestList(QuestNpc(QuestIndex)).Obj < (Index + 1) Then Exit Sub
    
    QuestObjIndex = Index + 1
    
    Call ShowInfoItem(QuestList(QuestNpc(QuestIndex)).RewardObjs(1).ObjIndex)
    
End Sub
Private Sub ItemRequired_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If QuestIndex <= 0 Then Exit Sub
  If QuestList(QuestNpc(QuestIndex)).Obj < (Index + 1) Then Exit Sub
    
    QuestObjIndex = Index + 1
    'PanelCommerce = ePanelCommerce.eRequiredObj_Selected
    Call ShowInfoItem(QuestList(QuestNpc(QuestIndex)).Objs(QuestObjIndex).ObjIndex)
End Sub

Private Sub LastNpc_Click()
    If QuestIndex <= 0 Then Exit Sub
    If QuestList(QuestNpc(QuestIndex)).Npc = 1 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    QuestNpcIndex = QuestNpcIndex - 1
    If QuestNpcIndex <= 0 Then QuestNpcIndex = 1
    Render_QuestPanel
End Sub

Private Sub LastQuest_Click()
    If QuestIndex <= 0 Then Exit Sub
    If QuestList(QuestNpc(QuestIndex)).LastQuest <= 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    
    Dim Last As Byte
    QuestNpcIndex = 0
    QuestObjIndex = 0
    Last = QuestNpc(QuestIndex)
    
    QuestNpc(QuestIndex) = QuestList(Last).LastQuest
End Sub

Private Sub NextNpc_Click()
    If QuestIndex <= 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    QuestNpcIndex = QuestNpcIndex + 1
    If QuestNpcIndex > QuestList(QuestNpc(QuestIndex)).Npc Then QuestNpcIndex = QuestList(QuestNpc(QuestIndex)).Npc
    Render_QuestPanel
End Sub

Private Sub NextQuest_Click()
    If QuestIndex <= 0 Then Exit Sub
    If QuestList(QuestNpc(QuestIndex)).NextQuest <= 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Last As Byte
    QuestNpcIndex = 0
    QuestObjIndex = 0
    Last = QuestNpc(QuestIndex)
    
    QuestNpc(QuestIndex) = QuestList(Last).NextQuest
End Sub

Private Sub Npc_Click()
    If QuestIndex <= 0 Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    PanelCommerce = ePanelCommerce.eRequiredNpc
    QuestNpcIndex = 1
    Render_QuestPanel

    
End Sub

Private Sub tUpdate_Timer()
  
    Render_QuestPanel
End Sub

