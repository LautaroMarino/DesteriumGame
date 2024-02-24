VERSION 5.00
Begin VB.Form FrmRuleta 
   BorderStyle     =   0  'None
   Caption         =   "Ruleta"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
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
   Picture         =   "FrmRuleta.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   525
      Left            =   465
      MousePointer    =   4  'Icon
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5010
      Width           =   4200
   End
   Begin VB.PictureBox picAccount 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   240
      MousePointer    =   99  'Custom
      Picture         =   "FrmRuleta.frx":5CCB8
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2700
      Width           =   4695
   End
   Begin VB.Image imgDsp 
      Height          =   495
      Left            =   2880
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image imgGld 
      Height          =   495
      Left            =   600
      Top             =   6240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmRuleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InvRuleta As clsGrapchicalInventory
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        FrmMain.SetFocus
        Unload Me
        
    End If
End Sub

Private Sub Form_Load()
    
    Set InvRuleta = New clsGrapchicalInventory
    
    Dim A As Long
    
    With RuletaConfig
        g_Captions(eCaption.eInvRuleta) = wGL_Graphic.Create_Device_From_Display(picInv.hWnd, picInv.ScaleWidth, picInv.ScaleHeight)
        InvRuleta.Initialize picInv, .ItemLast, .ItemLast, eCaption.eInvRuleta
        
        For A = 1 To .ItemLast
        
            InvRuleta.SetItem A, .Items(A).ObjIndex, .Items(A).Amount, 0, ObjData(.Items(A).ObjIndex).GrhIndex, ObjData(.Items(A).ObjIndex).ObjType, 0, 0, 0, 0, 0, ObjData(.Items(A).ObjIndex).Name, 0, True, 0, 0, 0, 0
        Next A
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set InvRuleta = Nothing
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eInvRuleta))
End Sub

Private Sub imgDsp_Click()
    WriteTirarRuleta 2
End Sub

Private Sub imgGld_Click()
    WriteTirarRuleta 1
End Sub

