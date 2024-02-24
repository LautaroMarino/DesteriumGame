VERSION 5.00
Begin VB.Form FrmMercader_ListInfo 
   BorderStyle     =   0  'None
   Caption         =   "Información de publicación"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercader_ListInfo.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicDraw 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   6900
      Left            =   570
      ScaleHeight     =   460
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   737
      TabIndex        =   0
      Top             =   1320
      Width           =   11055
   End
   Begin VB.Image imgConfirm 
      Height          =   435
      Left            =   8880
      Top             =   8430
      Width           =   2505
   End
   Begin VB.Image imgUnload 
      Height          =   465
      Left            =   10650
      Top             =   720
      Width           =   915
   End
   Begin VB.Image imgSpells 
      Height          =   405
      Left            =   2490
      Top             =   8490
      Width           =   915
   End
   Begin VB.Image imgBov 
      Height          =   435
      Left            =   1440
      Top             =   8520
      Width           =   855
   End
   Begin VB.Image imgInventory 
      Height          =   435
      Left            =   480
      Top             =   8490
      Width           =   855
   End
   Begin VB.Image imgStats 
      Height          =   435
      Left            =   3450
      Top             =   8490
      Width           =   855
   End
End
Attribute VB_Name = "FrmMercader_ListInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    g_Captions(eCaption.eMercader_List) = wGL_Graphic.Create_Device_From_Display(FrmMercader_ListInfo.PicDraw.hWnd, FrmMercader_ListInfo.PicDraw.ScaleWidth, FrmMercader_ListInfo.PicDraw.ScaleHeight)
    SwitchMap (1)
    MirandoMercader = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMercader_List))
End Sub
Private Sub imgInventory_Click()
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Inv.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgSpells_Click()
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Meditations.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgStats_Click()
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Other.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgBov_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    
    Call FrmMercader_Bank.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgConfirm_Click()
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_OfferSend.Show(vbModeless, FrmMercader_ListInfo)
End Sub


Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    MirandoMercader = False
    Unload Me
    
   ' FrmMercader_List.SetFocus
End Sub



