VERSION 5.00
Begin VB.Form FrmGuilds_Leader 
   BorderStyle     =   0  'None
   Caption         =   "Panel de Lider"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMembers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2340
      Left            =   525
      TabIndex        =   0
      Top             =   1295
      Width           =   1995
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmGuilds_Leader.frx":0000
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblElv 
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
      Left            =   3435
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   3435
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image imgKick 
      Height          =   255
      Left            =   3240
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgReturn 
      Height          =   375
      Left            =   2040
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "FrmGuilds_Leader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterface & "menucompacto\guilds_leader.jpg")
    
    Call ListarMiembros
End Sub


Private Sub ListarMiembros()
    Dim A As Long

    With GuildsInfo(Selected_GuildIndex)
        For A = 1 To MAX_GUILD_MEMBER
            If .Members(A).Elv > 0 Then
                lstMembers.AddItem .Members(A).Name
            End If
        Next A
    End With
End Sub

Private Sub imgKick_Click()

    If lstMembers.ListIndex = -1 Then Exit Sub
    
    Dim Name As String
    Name = lstMembers.List(lstMembers.ListIndex)
    
    If MsgBox("¿Estás seguro que deseas quitar de tu clan a " & Name & "?") = vbYes Then
        'Call ParseUserCommand("/KICK " & Name)
        Call WriteGuilds_Kick(Name)
    End If
End Sub

Private Sub imgReturn_Click()
    imgUnload_Click
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteGuilds_Required(0)
    Unload Me
End Sub

Private Sub lstMembers_Click()
    
    
    Dim Slot As Integer
    Slot = SearchMember(UCase$(Name))
    
    If Slot > 0 Then
        lblElv.Caption = GuildsInfo(Selected_GuildIndex).Members(Slot).Elv
    End If
End Sub

Private Function SearchMember(ByVal Name As String)
    Dim A As Long

    With GuildsInfo(Selected_GuildIndex)
        For A = 1 To MAX_GUILD_MEMBER
            If StrComp(UCase$(.Members(A).Name), Name) = 0 Then
                SearchMember = A
                Exit Function
            End If
        Next A
    End With
End Function
