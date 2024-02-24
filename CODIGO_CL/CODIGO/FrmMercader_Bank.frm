VERSION 5.00
Begin VB.Form FrmMercader_Bank 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Boveda de Personajes"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
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
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   1965
   End
   Begin VB.PictureBox PicBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   390
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   258
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   3870
   End
End
Attribute VB_Name = "FrmMercader_Bank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbChars_Click()
    If MercaderSelectedChar <> -1 Then
        Set Mercader_Bank = Nothing
    End If
    
    MercaderSelectedChar = cmbChars.ListIndex
    
    Dim A As Long
    
    
    g_Captions(eCaption.eMercader_Bank) = wGL_Graphic.Create_Device_From_Display(PicBank.hWnd, PicBank.ScaleWidth, PicBank.ScaleHeight)
    Set Mercader_Bank = New clsGrapchicalInventory

    Call Mercader_Bank.Initialize(FrmMercader_Bank.PicBank, MAX_BANCOINVENTORY_SLOTS, eCaption.eMercader_Bank)

    For A = 1 To MAX_BANCOINVENTORY_SLOTS
        With MercaderChars(MercaderSelectedChar)
            Call Mercader_Bank.SetItem(A, .ObjectBank(A).ObjIndex, .ObjectBank(A).Amount, 0, .ObjectBank(A).GrhIndex, 0, 0, 0, 0, 0, 0, .ObjectBank(A).Name, 0, True, 0, 0, 0, 0)
        End With
    
    Next A
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMercader_Bank))
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

