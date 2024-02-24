VERSION 5.00
Begin VB.Form FrmMercader_Offer 
   BorderStyle     =   0  'None
   Caption         =   "Oferta seleccionada"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercader_Offer.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicDraw 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   6900
      Left            =   510
      ScaleHeight     =   460
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   737
      TabIndex        =   0
      Top             =   1320
      Width           =   11055
   End
   Begin VB.Image imgAccept 
      Height          =   405
      Left            =   10680
      Top             =   8460
      Width           =   855
   End
   Begin VB.Image imgRechace 
      Height          =   405
      Left            =   9630
      Top             =   8460
      Width           =   915
   End
   Begin VB.Image imgSpells 
      Height          =   405
      Left            =   3450
      Top             =   8490
      Width           =   915
   End
   Begin VB.Image imgBov 
      Height          =   435
      Left            =   2490
      Top             =   8490
      Width           =   855
   End
   Begin VB.Image imgInventory 
      Height          =   435
      Left            =   1470
      Top             =   8460
      Width           =   855
   End
   Begin VB.Image imgStats 
      Height          =   435
      Left            =   480
      Top             =   8490
      Width           =   855
   End
   Begin VB.Image imgUnload 
      Height          =   435
      Left            =   10680
      Top             =   720
      Width           =   915
   End
End
Attribute VB_Name = "FrmMercader_Offer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SwitchMap (1)
    g_Captions(eCaption.eMercader_ListOffer) = wGL_Graphic.Create_Device_From_Display(picDraw.hWnd, picDraw.ScaleWidth, picDraw.ScaleHeight)
    Render_Mercader_Chars_Offer
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMercader_ListOffer))
End Sub
Private Sub imgAccept_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    'Prepare_And_Connect E_MODO.e_MercaderAcceptOffer
    Call WriteMercader_Required(5)
End Sub



Private Sub imgRechace_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    'Prepare_And_Connect E_MODO.e_MercaderRemoveOffer
    Call WriteMercader_Required(6)
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub

Private Sub imgInventory_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Inv.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgSpells_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Meditations.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgStats_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Other.Show(vbModeless, FrmMercader_ListInfo)
End Sub

Private Sub imgBov_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call FrmMercader_Bank.Show(vbModeless, FrmMercader_ListInfo)
End Sub
