VERSION 5.00
Begin VB.Form FrmMercader_Offers 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas recibidas"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercader_Offers.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPublication 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   5070
      Left            =   930
      TabIndex        =   0
      Top             =   1890
      Width           =   10185
   End
   Begin VB.Image imgUnload 
      Height          =   405
      Left            =   10650
      Top             =   750
      Width           =   915
   End
   Begin VB.Image imgInfo 
      Height          =   465
      Left            =   8430
      Top             =   7590
      Width           =   2505
   End
   Begin VB.Image imgUpdate 
      Height          =   465
      Left            =   5760
      Top             =   7620
      Width           =   2505
   End
End
Attribute VB_Name = "FrmMercader_Offers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub imgInfo_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    
    If lstPublication.ListIndex = -1 Then Exit Sub
    
    'Prepare_And_Connect E_MODO.e_MercaderListInfoOffer
    Call WriteMercader_Required(4)
     
    If FrmMercader_ListInfo.Visible Then
        Unload FrmMercader_ListInfo
    End If
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub

Private Sub imgUpdate_Click()
   ' Call Audio.PlayInterface(SND_CLICK)

End Sub

Private Sub lstPublication_Click()
    MercaderSelected = Val(ReadField(1, lstPublication.List(lstPublication.ListIndex), Asc("°")))
End Sub
