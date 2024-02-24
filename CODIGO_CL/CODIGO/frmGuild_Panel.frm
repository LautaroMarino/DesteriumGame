VERSION 5.00
Begin VB.Form frmGuild_Panel 
   BorderStyle     =   0  'None
   Caption         =   "Clanes"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmGuild_Panel.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
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
      Height          =   7110
      Left            =   210
      ScaleHeight     =   474
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   514
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7710
   End
   Begin VB.Image imgUnload 
      Height          =   585
      Left            =   7890
      Top             =   450
      Width           =   315
   End
End
Attribute VB_Name = "frmGuild_Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MoveForm As clsFormMovementManager

Private Sub Form_Load()
    
    ' Movimiento de pantalla
    Set MoveForm = New clsFormMovementManager
    MoveForm.Initialize Me
    
    'g_Captions(eCaption.eGuildPanel) = wGL_Graphic.Create_Device_From_Display(picDraw.hWnd, picDraw.ScaleWidth, picDraw.ScaleHeight)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    ' call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eGuildPanel))
End Sub
Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call WriteGuilds_Required(0)
    MirandoGuildPanel = False
    Unload Me
End Sub

