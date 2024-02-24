VERSION 5.00
Begin VB.Form FrmCriatura_Info 
   BorderStyle     =   0  'None
   Caption         =   "Información de la Criatura"
   ClientHeight    =   5250
   ClientLeft      =   8535
   ClientTop       =   3630
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCriatura_Info.frx":0000
   LinkTopic       =   "Info Criature"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tUpdate 
      Interval        =   200
      Left            =   1680
      Top             =   735
   End
   Begin VB.Image Drop 
      Height          =   480
      Index           =   5
      Left            =   2580
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Drop 
      Height          =   480
      Index           =   4
      Left            =   2100
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Drop 
      Height          =   480
      Index           =   3
      Left            =   1650
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Drop 
      Height          =   480
      Index           =   2
      Left            =   1155
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Drop 
      Height          =   480
      Index           =   1
      Left            =   690
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Drop 
      Height          =   480
      Index           =   0
      Left            =   210
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   5
      Left            =   2595
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   4
      Left            =   2100
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   3
      Left            =   1635
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   2
      Left            =   1140
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   1
      Left            =   690
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Item 
      Height          =   480
      Index           =   0
      Left            =   210
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image imgUnload 
      Height          =   435
      Left            =   2835
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "FrmCriatura_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormMovement As clsFormMovementManager



Private Sub Form_Load()

    Set FormMovement = New clsFormMovementManager
    
    Call FormMovement.Initialize(Me, 32)
    
  '  g_Captions(eCaption.cCriaturaInfo) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cCriaturaInfo))
End Sub



Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If frmCriatura_Quest.visible Then
        frmCriatura_Quest.SetFocus
    End If
    
    Unload Me
End Sub

Private Sub Item_Click(Index As Integer)
    If NpcList(SelectedNpcIndex).Object(Index + 1).ObjIndex = 0 Then Exit Sub
    
    SelectedObjIndex = NpcList(SelectedNpcIndex).Object(Index + 1).ObjIndex
    
    Call SelectedObjIndex_Update
 
End Sub
Private Sub Drop_Click(Index As Integer)
    If NpcList(SelectedNpcIndex).Drop(Index + 1).ObjIndex = 0 Then Exit Sub
    
    SelectedObjIndex = NpcList(SelectedNpcIndex).Drop(Index + 1).ObjIndex
    
    Call SelectedObjIndex_Update
End Sub


Private Sub tUpdate_Timer()
    If SelectedNpcIndex = 0 Then Exit Sub
    Render_CriaturaInfo
End Sub
