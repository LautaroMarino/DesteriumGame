VERSION 5.00
Begin VB.Form frmParty 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   Icon            =   "frmParty.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "frmParty.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblRewardExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   4
      Left            =   3255
      TabIndex        =   19
      Top             =   6105
      Width           =   1155
   End
   Begin VB.Label lblRewardExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   3
      Left            =   3255
      TabIndex        =   18
      Top             =   5655
      Width           =   1155
   End
   Begin VB.Label lblRewardExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   2
      Left            =   3255
      TabIndex        =   17
      Top             =   5205
      Width           =   1155
   End
   Begin VB.Label lblRewardExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   3255
      TabIndex        =   16
      Top             =   4755
      Width           =   1155
   End
   Begin VB.Image imgCheck 
      Height          =   315
      Left            =   585
      Top             =   3630
      Width           =   315
   End
   Begin VB.Label lblRewardExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   0
      Left            =   3270
      TabIndex        =   15
      Top             =   4305
      Width           =   1155
   End
   Begin VB.Label lblReward 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   14
      Top             =   6105
      Width           =   795
   End
   Begin VB.Label lblReward 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   13
      Top             =   5655
      Width           =   795
   End
   Begin VB.Label lblReward 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   12
      Top             =   5190
      Width           =   795
   End
   Begin VB.Label lblReward 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   11
      Top             =   4755
      Width           =   795
   End
   Begin VB.Label lblReward 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   795
   End
   Begin VB.Image imgSavePorc 
      Height          =   495
      Left            =   3000
      Picture         =   "frmParty.frx":10BE
      Top             =   6600
      Width           =   1680
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   9
      Top             =   3270
      Width           =   795
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   8
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   2370
      Width           =   795
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   795
   End
   Begin VB.Image imgAbandonate 
      Height          =   495
      Left            =   600
      Picture         =   "frmParty.frx":6A8C
      Top             =   6600
      Width           =   1680
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   4
      Left            =   3600
      TabIndex        =   5
      Top             =   3270
      Width           =   525
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   4
      Top             =   2820
      Width           =   525
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   2370
      Width           =   525
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   0
      Left            =   3585
      TabIndex        =   1
      Top             =   1470
      Width           =   525
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "<Vacio>"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1470
      Width           =   2535
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   0
      Left            =   3720
      Top             =   7500
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public picCheckBox          As Picture
Public picCheckBoxNulo      As Picture

Public cRewardExp As Boolean

Private BotonAbandonar As clsGraphicalButton
Private BotonGuardar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton
Private clsFormulario As clsFormMovementManager

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterface & "menucompacto\group.jpg")
    
    #If ModoBig = 0 Then
            ' Handles Form movement (drag and dr|op).
            Set clsFormulario = New clsFormMovementManager
            clsFormulario.Initialize Me, 120
    #End If
    
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    GrhPath = DirInterface & "menucompacto\buttons\"
    
    Set picCheckBox = LoadPicture(GrhPath & "check.jpg")
    Set picCheckBoxNulo = LoadPicture(GrhPath & "nocheck.jpg")
    
    Set BotonAbandonar = New clsGraphicalButton
    Set BotonGuardar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call BotonAbandonar.Initialize(imgAbandonate, GrhPath & "abandonar.jpg", GrhPath & "abandonar_hover.jpg", GrhPath & "abandonar.jpg", Me)
    Call BotonGuardar.Initialize(imgSavePorc, GrhPath & "guardar.jpg", GrhPath & "guardar_hover.jpg", GrhPath & "guardar.jpg", Me)
  
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAbandonate_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Protocol.WritePartyClient 4
    Unload Me
End Sub

Private Sub imgCheck_Click()
        
        Call Audio.PlayInterface(SND_CLICK)
        
        If UCase$(lblUser(0).Caption) <> UCase$(FrmMain.Label8(0).Caption) Then
            Call ShowConsoleMsg("Solo el Líder puede cambiar esta opción...")
            Exit Sub
        End If
        
        cRewardExp = Not cRewardExp
        
        If cRewardExp Then
            imgCheck.Picture = picCheckBox
        Else
            imgCheck.Picture = picCheckBoxNulo
        End If
        
        WritePartyClient 2
End Sub

Private Sub imgSavePorc_Click()
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Exp(4)   As Byte

    Dim A        As Byte

    Dim TotalExp As Integer
    
    For A = 0 To 4
        Exp(A) = lblExp(A)
        
        TotalExp = TotalExp + Exp(A)
        
    Next A
    
    If TotalExp <> 100 Then Exit Sub
    
    WriteGroupChangePorc Exp
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub lblExp_Click(Index As Integer)

    Dim Temp

    Temp = InputBox("Elige el porcentaje de experiencia que deseas para este personaje", "Grupos Desterium: Edición de porcentaje", lblExp(Index).Caption)

    lblExp(Index).Caption = Val(Temp)
End Sub
