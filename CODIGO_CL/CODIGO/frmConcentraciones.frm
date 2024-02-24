VERSION 5.00
Begin VB.Form frmConcentraciones 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   Picture         =   "frmConcentraciones.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   8145
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
      Height          =   4695
      Index           =   0
      Left            =   350
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   9
      Top             =   2100
      Width           =   4215
   End
   Begin VB.ComboBox cmbMeditation 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   390
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1710
      Width           =   4155
   End
   Begin VB.Label lblPremium 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5520
      TabIndex        =   7
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label lblOro 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5070
      TabIndex        =   6
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label lblPlata 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5220
      TabIndex        =   5
      Top             =   4230
      Width           =   225
   End
   Begin VB.Label lblBronce 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5370
      TabIndex        =   4
      Top             =   3930
      Width           =   225
   End
   Begin VB.Label lblFaction 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5430
      TabIndex        =   3
      Top             =   3330
      Width           =   1605
   End
   Begin VB.Label lblBlue 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   5430
      TabIndex        =   2
      Top             =   3030
      Width           =   1365
   End
   Begin VB.Label lblRed 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5100
      TabIndex        =   1
      Top             =   2700
      Width           =   1365
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   5250
      TabIndex        =   0
      Top             =   2400
      Width           =   405
   End
   Begin VB.Image imgReclamar 
      Height          =   630
      Left            =   6000
      Picture         =   "frmConcentraciones.frx":23873
      Top             =   6480
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image imgUnload 
      Height          =   405
      Left            =   7950
      Top             =   540
      Width           =   225
   End
   Begin VB.Image imgUse 
      Height          =   435
      Left            =   6120
      Top             =   6540
      Width           =   1695
   End
End
Attribute VB_Name = "frmConcentraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub cmbMeditation_Click()
    Call SetInfo(cmbMeditation.ListIndex)
End Sub

Private Sub SetInfo(ByVal Selected As Integer)
    
    MeditationSelected = Selected

    With Meditation(MeditationSelected)
        lblLvl.Caption = .RequiredLvl
        lblRed.Caption = .RequiredRed
        lblBlue.Caption = .RequiredBlue
        
        If .RequiredFaction = 1 Then
            lblFaction.Caption = "<Armada Real>"
            lblFaction.ForeColor = vbCyan
        ElseIf .RequiredFaction = 2 Then
            lblFaction.Caption = "<Legión Oscura>"
            lblFaction.ForeColor = vbRed
        Else
            lblFaction.Caption = "NO"
            lblFaction.ForeColor = vbWhite
        End If
        
        If .RequiredBronce = 1 Then
            lblBronce.Caption = "SI"
            lblBronce.ForeColor = vbGreen
        Else
            lblBronce.Caption = "NO"
            lblBronce.ForeColor = vbRed
        End If
        
        If .RequiredPlata = 1 Then
            lblPlata.Caption = "SI"
            lblPlata.ForeColor = vbGreen
        Else
            lblPlata.Caption = "NO"
            lblPlata.ForeColor = vbRed
        End If
        
        If .RequiredGold = 1 Then
            lblOro.Caption = "SI"
            lblOro.ForeColor = vbGreen
        Else
            lblOro.Caption = "NO"
            lblOro.ForeColor = vbRed
        End If
        
        If .RequiredPremium = 1 Then
            lblPremium.Caption = "SI"
            lblPremium.ForeColor = vbGreen
        Else
            lblPremium.Caption = "NO"
            lblPremium.ForeColor = vbRed
        End If
        
        If .Learned = 0 And Selected <> 0 Then
            imgReclamar.visible = True
        Else
            imgReclamar.visible = False
        End If
        
        Call InitGrh(GrhMeditationForm, FxData(.FX).Animacion)
    End With
    
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    MeditationSelected = 0
    
    
    ' Ventana de Meditaciones
    g_Captions(eCaption.Meditations) = wGL_Graphic.Create_Device_From_Display(frmConcentraciones.PicDraw(0).hWnd, frmConcentraciones.PicDraw(0).ScaleWidth, frmConcentraciones.PicDraw(0).ScaleHeight)


End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.Meditations))
End Sub
Private Sub imgReclamar_Click()
    
    If MeditationSelected <= 0 Then Exit Sub
    Call WriteLearnMeditation(0, MeditationSelected)
    
    imgUnload_Click
End Sub

Private Sub imgUnload_Click()
    MeditationSelected = 0
    MirandoConcentracion = False
    Unload Me
End Sub

Private Sub imgUse_Click()

    If MeditationSelected < 0 Then Exit Sub
    Call WriteLearnMeditation(1, MeditationSelected)
    
    imgUnload_Click
End Sub

