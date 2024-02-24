VERSION 5.00
Begin VB.Form FrmPremios 
   BorderStyle     =   0  'None
   Caption         =   "Premios"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Premios Desterium"
   ScaleHeight     =   7605
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmPremios.frx":0000
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgPagination 
      Height          =   375
      Index           =   1
      Left            =   2880
      Top             =   6720
      Width           =   255
   End
   Begin VB.Image imgPagination 
      Height          =   375
      Index           =   0
      Left            =   1920
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label lblPagination 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   6800
      Width           =   255
   End
End
Attribute VB_Name = "FrmPremios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Const MAX_PAGINATION As Byte = 1
Public LastPagination As Integer

Private Sub Form_Load()


    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    Dim FilePath As String
    
    FilePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(FilePath & "Premios1.jpg")
    
    LastPagination = 1
    
End Sub

Private Sub imgPagination_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    
    Select Case Index
        Case 0
            If Not LastPagination > 1 Then Exit Sub
            LastPagination = LastPagination - 1
        Case 1
            If Not LastPagination < MAX_PAGINATION Then Exit Sub
            LastPagination = LastPagination + 1
    End Select
    
    
    Dim FilePath As String
    
    FilePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(FilePath & "Premios" & LastPagination & ".jpg")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub
