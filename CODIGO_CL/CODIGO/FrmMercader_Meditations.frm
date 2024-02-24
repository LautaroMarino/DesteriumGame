VERSION 5.00
Begin VB.Form FrmMercader_Meditations 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meditaciones de Personajes"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMeditations 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1965
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3540
      Top             =   210
   End
   Begin VB.PictureBox PicDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   90
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   306
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   990
      Width           =   4590
   End
   Begin VB.ComboBox cmbChars 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   1965
   End
End
Attribute VB_Name = "FrmMercader_Meditations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbChars_Click()
    
    MercaderSelectedChar = cmbChars.ListIndex
    Selected
End Sub

Private Sub cmbMeditations_Click()
    
    MercaderSelectedMeditation = Val(cmbMeditations.List(cmbMeditations.ListIndex))
    Call InitGrh(GrhMeditationForm, FxData(MercaderSelectedMeditation).Animacion)
End Sub

Private Sub Form_Load()

    Dim A As Long
    
    MercaderSelectedChar = -1
    MercaderSelectedMeditation = -1
    
    cmbChars.Clear
    cmbMeditations.Clear
    
    For A = 0 To 4
        If MercaderChars(A).Name <> vbNullString Then
            cmbChars.AddItem MercaderChars(A).Name
        End If
    Next A

    g_Captions(eCaption.eMercader_Meditations) = wGL_Graphic.Create_Device_From_Display(picDraw.hWnd, picDraw.ScaleWidth, picDraw.ScaleHeight)
    
    
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMercader_Meditations))
End Sub
Private Sub Timer1_Timer()
    Render_Mercader_Chars_Meditations
End Sub


Private Sub Selected()

    Dim A As Long
    
    cmbMeditations.Clear
    
    For A = 1 To MAX_MEDITATION
        
        If MercaderChars(MercaderSelectedChar).Meditations(A) <> 0 Then
            cmbMeditations.AddItem (MercaderChars(MercaderSelectedChar).Meditations(A))
            
        End If
    Next A
    
    
End Sub
