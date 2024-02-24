VERSION 5.00
Begin VB.Form FrmSubasta 
   BorderStyle     =   0  'None
   Caption         =   "Subastas"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmSubasta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEldhir 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   2835
      TabIndex        =   3
      Text            =   "1"
      Top             =   2520
      Width           =   960
   End
   Begin VB.TextBox txtGld 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   2835
      TabIndex        =   2
      Text            =   "1"
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   5670
      TabIndex        =   1
      Text            =   "1"
      Top             =   4230
      Width           =   750
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
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
      Height          =   1440
      Left            =   810
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   2985
      Width           =   4800
   End
   Begin VB.Image imgSubastar 
      Height          =   795
      Left            =   1470
      Top             =   210
      Width           =   3540
   End
   Begin VB.Image imgUnload 
      Height          =   510
      Left            =   6195
      MouseIcon       =   "FrmSubasta.frx":000C
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FrmSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private Inv As clsGrapchicalInventory
Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    g_Captions(eCaption.cSubastar) = wGL_Graphic.Create_Device_From_Display(picInv.hWnd, picInv.ScaleWidth, picInv.ScaleHeight)
    
    Dim A As Long
    
    
    Set Inv = New clsGrapchicalInventory
    
    Call Inv.Initialize(picInv, MAX_INVENTORY_SLOTS, MAX_INVENTORY_SLOTS, eCaption.cSubastar, , , , , , , , , , True)
    
    For A = 1 To MAX_INVENTORY_SLOTS
        With Inventario
            Call Inv.SetItem(A, .ObjIndex(A), .Amount(A), 0, .GrhIndex(A), .ObjType(A), 0, 0, 0, 0, .Valor(A), .ItemName(A), 0, .CanUse(A), 0, 0, 0, 0)
        End With
    Next A
    
    Inv.DrawInventory
    
    Me.Picture = LoadPicture(App.path & "\resource\interface\bank\bank_auction.jpg")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
     
    Set Inv = Nothing
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cSubastar))
    
End Sub

Private Sub imgSubastar_Click()
    If Not CheckData Then Exit Sub
    
    If MsgBox("¿Estás seguro que deseas subastar el objeto " & Inventario.ItemName(Inv.SelectedItem) & " (x" & Val(txtAmount.Text) & ")?", vbYesNo) = vbYes Then
        Call WriteAuction_New(Inv.SelectedItem, Val(txtAmount.Text), Val(txtGld.Text), Val(txtEldhir.Text))
        Unload Me
    End If
    
End Sub

Private Function CheckData() As Boolean
    
    If Inv.SelectedItem = 0 Then Exit Function
    If Val(txtGld.Text) < 0 Then Exit Function
    If Val(txtEldhir.Text) < 0 Then Exit Function
    If Val(txtAmount.Text) <= 0 Then Exit Function
    
    If Val(txtAmount.Text) > 10000 Then
        Call MsgBox("No puedes subastar más de 10.000 Objetos.")
        Exit Function
    End If
    
    If Val(txtGld.Text) > 100000000 Then
        Call MsgBox("El máximo de Oro que puedes pedir es de 100.000.000")
        Exit Function
    End If
    
    If Val(txtEldhir.Text) > 1000 Then
        Call MsgBox("El máximo de Eldhires que puedes pedir es de 1.000")
        Exit Function
    End If
    
    If UserGLD < 20000 Then
        Call MsgBox("El servidor te cobrá 20.000 Monedas de Oro para realizar la subasta.")
        Exit Function
    End If
    
    
    CheckData = True
End Function
Private Sub imgUnload_Click()
    Unload Me
End Sub

