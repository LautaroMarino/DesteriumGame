VERSION 5.00
Begin VB.Form FrmBank_Deposit 
   BorderStyle     =   0  'None
   Caption         =   "Deposito de Monedas"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleMode       =   0  'User
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptionGld 
      BackColor       =   &H80000009&
      Caption         =   "Option1"
      Height          =   195
      Left            =   4305
      TabIndex        =   2
      Top             =   1785
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.OptionButton OptionEldhir 
      BackColor       =   &H80000009&
      Caption         =   "Option1"
      Height          =   195
      Left            =   4305
      TabIndex        =   1
      Top             =   2820
      Width           =   195
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   210
      Left            =   3990
      TabIndex        =   0
      Text            =   "0"
      Top             =   3675
      Width           =   1515
   End
   Begin VB.Image imgUnload 
      Height          =   435
      Left            =   6195
      Top             =   0
      Width           =   435
   End
   Begin VB.Label lblGld 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9.999.999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2100
      TabIndex        =   4
      Top             =   2310
      Width           =   2475
   End
   Begin VB.Label lblEldhir 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9.999.999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   2100
      TabIndex        =   3
      Top             =   3255
      Width           =   2475
   End
   Begin VB.Image imgAdd 
      Height          =   315
      Left            =   5565
      Top             =   3360
      Width           =   585
   End
   Begin VB.Image imgRemove 
      Height          =   345
      Left            =   5565
      Top             =   3885
      Width           =   615
   End
   Begin VB.Image imgGld 
      Height          =   360
      Left            =   2310
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Image imgEldhir 
      Height          =   345
      Left            =   2415
      Top             =   2730
      Width           =   1770
   End
End
Attribute VB_Name = "FrmBank_Deposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\resource\interface\bank\bank_deposit.jpg")

    lblGld.Caption = IIf(UserBankGold > 0, Format$(UserBankGold, "##,##"), "0")
    lblEldhir.Caption = IIf(UserBankEldhir > 0, Format$(UserBankEldhir, "##,##"), "0")
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub

Private Sub imgAdd_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Val(txtAmount.Text) <= 0 Then
        Exit Sub
    End If

    If OptionGld.Value Then
        Call WriteBankGold(Val(txtAmount.Text), 0, False)
    ElseIf OptionEldhir.Value Then
        
        Call WriteBankGold(Val(txtAmount.Text), 1, False)
    End If
    
    
    
End Sub

Private Sub imgEldhir_Click()
    OptionEldhir.Value = True
End Sub

Private Sub imgGld_Click()
    OptionGld.Value = True
End Sub

Private Sub imgRemove_Click()
    Call Audio.PlayInterface(SND_CLICK)
    If Val(txtAmount.Text) <= 0 Then
        Exit Sub
    End If
    
    If OptionGld.Value Then
        Call WriteBankGold(Val(txtAmount.Text), 0, True)
    ElseIf OptionEldhir.Value Then
        Call WriteBankGold(Val(txtAmount.Text), 1, True)
    End If
    
End Sub
