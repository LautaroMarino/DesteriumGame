VERSION 5.00
Begin VB.Form frmGuild 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Clanes"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuild.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Alianzas de Combate"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   180
      Left            =   2520
      TabIndex        =   1
      Top             =   5355
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Timer tUpdate 
      Interval        =   40
      Left            =   210
      Top             =   210
   End
   Begin VB.TextBox txtGuild 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   525
      TabIndex        =   0
      Top             =   5250
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image imgAlineation 
      Height          =   1695
      Index           =   2
      Left            =   4725
      Top             =   2940
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgAlineation 
      Height          =   1695
      Index           =   1
      Left            =   2730
      Top             =   3045
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgAlineation 
      Height          =   1695
      Index           =   0
      Left            =   840
      Top             =   3045
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRequired 
      Height          =   435
      Index           =   4
      Left            =   4095
      Top             =   1785
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgRequired 
      Height          =   435
      Index           =   3
      Left            =   3570
      Top             =   1785
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgRequired 
      Height          =   435
      Index           =   2
      Left            =   3150
      Top             =   1785
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgRequired 
      Height          =   435
      Index           =   1
      Left            =   2625
      Top             =   1785
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgRequired 
      Height          =   435
      Index           =   0
      Left            =   2205
      Top             =   1785
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgEventos 
      Height          =   435
      Index           =   1
      Left            =   210
      Top             =   3990
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgEventos 
      Height          =   435
      Index           =   0
      Left            =   210
      Top             =   1155
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image lblMenu 
      Height          =   330
      Index           =   3
      Left            =   4200
      Top             =   630
      Width           =   2115
   End
   Begin VB.Image lblMenu 
      Height          =   330
      Index           =   2
      Left            =   210
      Top             =   630
      Width           =   2010
   End
   Begin VB.Image lblMenu 
      Height          =   330
      Index           =   0
      Left            =   1155
      Top             =   105
      Width           =   4110
   End
   Begin VB.Image imgMove 
      Height          =   285
      Index           =   1
      Left            =   5985
      Top             =   5880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgMove 
      Height          =   285
      Index           =   0
      Left            =   5985
      Top             =   5565
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgSearch 
      Height          =   330
      Left            =   525
      Top             =   4620
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgUnload 
      Height          =   375
      Left            =   6195
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private MoveForm As clsFormMovementManager


Private Sub Form_Load()
    ' Movimiento de pantalla
    Set MoveForm = New clsFormMovementManager
    MoveForm.Initialize Me
    
    g_Captions(eCaption.eguildpanel) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)


    LastSelectedGuild = 0
    GuildPanel = rGuildPanel.rList
    
    Dim A As Long
    
    'For A = 1 To MAX_GUILDS
        'GuildsInfo(A) = GuildsInfo_Copy(A)
    'Next A
    
    lblMenu_Click (0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MirandoGuildPanel = False
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eguildpanel))
End Sub

Private Sub imgAlineation_Click(Index As Integer)
    TempAlineation = Index
End Sub

Private Sub imgDetails_Click(Index As Integer)
    If GuildPanel <> rGuildPanel.rList Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    GuildPanel = rGuildPanel.rPerfil
    GuildSelected = GuildsInfo(Index).Index
    
    Call WriteGuilds_Required(GuildSelected)
End Sub

Private Sub imgFound_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call Guilds_FounderNew
    
    Unload Me
End Sub

Private Sub imgLeader_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteGuilds_Required(1000)
    
    Unload Me
End Sub

Private Sub imgEventos_Click(Index As Integer)
    If GuildPanel <> rGuildPanel.rEvents Then Exit Sub
    
    Select Case Index
    
        Case 0 ' Castillos Norte y Sur
        
        
        Case 1 ' Torneo de Clanes
        
        
    End Select
End Sub

Private Sub imgMove_Click(Index As Integer)
    If GuildPanel <> rGuildPanel.rList Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim A As Long
    
    'For A = 1 To MAX_GUILDS
        'GuildsInfo(A) = GuildsInfo_Copy(A)
    'Next A
    
    Select Case Index
    
        Case 0  ' Down
            If LastSelectedGuild > 0 Then
                LastSelectedGuild = LastSelectedGuild - 1
            End If
        Case 1  ' Up
            'If GuildsInfo(LastSelectedGuild + 11).Name = vbNullString Then Exit Sub
           
            If LastSelectedGuild < MAX_GUILDS Then
                 LastSelectedGuild = LastSelectedGuild + 1
            End If
    End Select
    
End Sub


Private Sub imgSearch_Click()
    If GuildPanel <> rGuildPanel.rList Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    If txtGuild.visible Then
        txtGuild.visible = False
    Else
        txtGuild.visible = True
    End If
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    FrmMain.SetFocus
    
    ' Volvemos al Principal antes de cerrar el formulario!!
    If GuildPanel <> rGuildPanel.rList Then
        GuildPanel = rList
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub lblMenu_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    If GuildPanel = rGuildPanel.rFound And Index = rGuildPanel.rFound Then
       
        
        Exit Sub
    End If
    
    GuildPanel = Index
    
    Dim A As Long
    
    For A = imgRequired.LBound To imgRequired.UBound
        imgRequired(A).visible = False
    Next A
    
    For A = imgEventos.LBound To imgEventos.UBound
        imgEventos(A).visible = False
    Next A
    
    For A = imgMove.LBound To imgMove.UBound
        imgMove(A).visible = False
    Next A
    
    For A = imgAlineation.LBound To imgAlineation.UBound
        imgAlineation(A).visible = False
    Next A
    
    txtGuild.visible = False
    txtName.visible = False
    
    Select Case GuildPanel
    
        Case rGuildPanel.rList
            For A = imgMove.LBound To imgMove.UBound
                imgMove(A).visible = True
            Next A
            
            imgSearch.visible = True
            
        Case rGuildPanel.rFound
            For A = imgRequired.LBound To imgRequired.UBound
                imgRequired(A).visible = True
            Next A
            
            For A = imgAlineation.LBound To imgAlineation.UBound
                imgAlineation(A).visible = True
            Next A
            
            txtName.visible = True
            
        Case rGuildPanel.rEvents
                
            For A = imgEventos.LBound To imgEventos.UBound
                imgEventos(A).visible = True
            Next A
    End Select

End Sub



