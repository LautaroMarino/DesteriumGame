VERSION 5.00
Begin VB.Form FrmMercader_Inv 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inventario de Personajes"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1965
   End
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
      Height          =   2880
      Left            =   1200
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      Top             =   540
      Width           =   2400
   End
End
Attribute VB_Name = "FrmMercader_Inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbChars_Click()
    If MercaderSelectedChar <> -1 Then
        Set Mercader_Inv = Nothing
    End If
    
    MercaderSelectedChar = cmbChars.ListIndex
    
    Dim A As Long
    
    
    g_Captions(eCaption.eMercader_Inv) = wGL_Graphic.Create_Device_From_Display(FrmMercader_Inv.picInv.hWnd, FrmMercader_Inv.picInv.ScaleWidth, FrmMercader_Inv.picInv.ScaleHeight)
    Set Mercader_Inv = New clsGrapchicalInventory

    Call Mercader_Inv.Initialize(FrmMercader_Inv.picInv, MAX_INVENTORY_SLOTS, eCaption.eMercader_Inv)

    For A = 1 To MAX_INVENTORY_SLOTS
        With MercaderChars(MercaderSelectedChar)
            Call Mercader_Inv.SetItem(A, .Object(A).ObjIndex, .Object(A).Amount, 0, .Object(A).GrhIndex, 0, 0, 0, 0, 0, 0, .Object(A).Name, 0, True, 0, 0, 0, 0)
        End With
    
    Next A
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMercader_Inv))
End Sub
Private Sub Form_Activate()
    
    Dim A As Long
    
    MercaderSelectedChar = -1
    cmbChars.Clear
    
    For A = 0 To 4
        With MercaderChars(A)
            If .Name <> vbNullString Then
                cmbChars.AddItem (.Name)
            End If
        End With
    Next A
 
End Sub

