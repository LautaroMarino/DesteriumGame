VERSION 5.00
Begin VB.Form FrmShop_Dsp 
   BorderStyle     =   0  'None
   Caption         =   "Cargar DSP"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
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
   Picture         =   "FrmShop_Dsp.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCargar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARGAR DSP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   7830
      TabIndex        =   2
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Image imgDSP 
      Height          =   825
      Left            =   7680
      Picture         =   "FrmShop_Dsp.frx":2F619
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   3765
      TabIndex        =   1
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Label lblTiers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIERS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   5880
      TabIndex        =   0
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   11640
      Picture         =   "FrmShop_Dsp.frx":334BB
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgTier 
      Height          =   825
      Left            =   5520
      Picture         =   "FrmShop_Dsp.frx":3456D
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Image imgGeneral 
      Height          =   825
      Left            =   3360
      Picture         =   "FrmShop_Dsp.frx":382D9
      Top             =   1080
      Width           =   2010
   End
End
Attribute VB_Name = "FrmShop_Dsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgGeneral_Click()
    FrmShop.Show , FrmMain
    Unload Me
End Sub
Private Sub lblGeneral_Click()
imgGeneral_Click
End Sub

Private Sub imgUnload_Click()
        Call Audio.PlayInterface(SND_CLICK)
    
    If MsgBox("¿Estás seguro que deseas salir?", vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub



Private Sub lblTiers_Click()
imgTier_Click
End Sub
Private Sub imgTier_Click()
   Call Audio.PlayInterface(SND_CLICK)
      
    imgGeneral.Picture = LoadPicture(DirInterface & "\generic\generic_hover.jpg")
    
    FrmShop_Tiers.Show , FrmMain
    Unload Me
End Sub
