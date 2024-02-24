VERSION 5.00
Begin VB.Form FrmGuilds_Found 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   3960
      Width           =   1665
   End
   Begin VB.ComboBox cmbAlineation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblGuild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Image imgFound 
      Height          =   495
      Left            =   960
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmGuilds_Found.frx":0000
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgReturn 
      Height          =   375
      Left            =   2040
      Top             =   6720
      Width           =   1215
   End
End
Attribute VB_Name = "FrmGuilds_Found"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Me.Picture = LoadPicture(DirInterface & "menucompacto\guilds_found.jpg")
    
    cmbAlineation.AddItem "Neutral"
    cmbAlineation.AddItem "Armada"
    cmbAlineation.AddItem "Legion"
    
    
    
    cmbAlineation.ListIndex = 0
    
    
    lblName.Caption = UserName
    lblGuild.Caption = "<" & txtName.Text & ">"
End Sub

Private Sub imgFound_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
     If txtName.Text = vbNullString Then
        Call MsgBox("Debes escoger el nombre que tendrá tu nueva Alianza.")
        Exit Sub
    End If
        
    If UserLvl < 35 Then
        Call MsgBox("Debes ser nivel 35 para fundar un clan")
        Exit Sub
    End If
    
    TempAlineation = cmbAlineation.ListIndex
    
    If MsgBox("¿Estás seguro de fundar el clan '" & txtName.Text & "'", vbYesNo) = vbYes Then
        Call WriteGuilds_Found(txtName.Text, TempAlineation)
    End If
End Sub

Private Sub imgReturn_Click()
    
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteGuilds_Required(0)
    Unload Me
End Sub

Private Sub imgUnload_Click()
    imgReturn_Click
End Sub

Private Sub txtName_Change()
    lblGuild.Caption = "<" & txtName.Text & ">"
End Sub
