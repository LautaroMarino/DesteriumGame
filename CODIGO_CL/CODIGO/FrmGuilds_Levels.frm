VERSION 5.00
Begin VB.Form FrmGuilds_Levels 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Niveles"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgReturn 
      Height          =   375
      Left            =   2040
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmGuilds_Levels.frx":0000
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmGuilds_Levels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterface & "menucompacto\guilds_levels.jpg")
End Sub

Private Sub imgReturn_Click()
    imgUnload_Click
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteGuilds_Required(0)
    Unload Me
End Sub
