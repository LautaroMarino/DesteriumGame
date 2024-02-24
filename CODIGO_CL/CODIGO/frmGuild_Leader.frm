VERSION 5.00
Begin VB.Form frmGuild_Leader 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
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
   Picture         =   "frmGuild_Leader.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1650
      Left            =   3720
      TabIndex        =   6
      Top             =   3990
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
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
      Height          =   1920
      Left            =   5700
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1920
   End
   Begin VB.ListBox lstUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3270
      Left            =   510
      TabIndex        =   0
      Top             =   1770
      Width           =   2295
   End
   Begin VB.Image imgCodex 
      Height          =   615
      Left            =   6390
      Top             =   6510
      Width           =   1635
   End
   Begin VB.Image imgRange 
      Height          =   525
      Left            =   6330
      Top             =   5820
      Width           =   1605
   End
   Begin VB.Label lblPoints 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   4770
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblRaze 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo Oscuro"
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
      Height          =   240
      Left            =   3630
      TabIndex        =   4
      Top             =   2550
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
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
      Height          =   240
      Left            =   3630
      TabIndex        =   3
      Top             =   2190
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblElv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "47"
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
      Height          =   240
      Left            =   3630
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgKick 
      Height          =   555
      Left            =   180
      Top             =   6540
      Width           =   1605
   End
   Begin VB.Image imgUnload 
      Height          =   525
      Left            =   7950
      Top             =   510
      Width           =   255
   End
End
Attribute VB_Name = "frmGuild_Leader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MoveForm As clsFormMovementManager

Private Sub Form_Load()
    ' Movimiento de pantalla
    Set MoveForm = New clsFormMovementManager
    MoveForm.Initialize Me
    
       ' g_Captions(eCaption.eGuildLeader) = wGL_Graphic.Create_Device_From_Display(picDraw.hWnd, picDraw.ScaleWidth, picDraw.ScaleHeight)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
   ' Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eGuildLeader))
End Sub

Private Sub imgCodex_Click()
    Call Audio.PlayInterface(SND_CLICK)
    frmGuild_Codex.Show vbModeless, frmMain
    Unload Me
End Sub

Private Sub imgKick_Click()
    If lstUser.ListIndex = -1 Then Exit Sub
    If lstUser.List(lstUser.ListIndex) = "<VACIO>" Then Exit Sub
    
    If lstUser.List(lstUser.ListIndex) = UCase$(frmMain.Label8(0).Caption) Then
        Call MsgBox("¡No puedes expulsarte a ti mismo!")
        Exit Sub
    End If
    
    If lstRange.ListIndex = eGuildRange.rFound Then
        Call MsgBox("¡El fundador no puede ser expulsado!")
        Exit Sub
    End If
    
    If MsgBox("¿Estás seguro que deseas expulsar al miembro " & lstUser.List(lstUser.ListIndex) & "?", vbYesNo) = vbYes Then
        
        Call WriteGuilds_Kick(lstUser.ListIndex + 1)
        
    End If
    
End Sub

Private Sub imgRange_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If lstUser.ListIndex = -1 Then Exit Sub
    If lstUser.List(lstUser.ListIndex) = "<VACIO>" Then Exit Sub
    If lstRange.ListIndex = -1 Then Exit Sub
    
     Dim Key As String
     
    If lstRange.ListIndex = eGuildRange.rFound Then
        Key = InputBox("Para seleccionar este personaje como fundador debes tener en cuenta que el anterior desaparecerá y será expulsado del clan.")
        
        If Len(Key) < ACCOUNT_MIN_CHARACTER_KEY Then
            Call MsgBox("La clave de seguridad es enviada al momento de crear la cuenta y tiene " & ACCOUNT_MIN_CHARACTER_KEY & " caracteres.")
            Exit Sub
        End If
    End If
    
    Call WriteGuilds_SetRange(lstUser.ListIndex + 1, lstRange.ListIndex, Key)
    Unload Me
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    frmGuild.Show vbModeless, frmMain
    MirandoGuildRank = True
    Unload Me
End Sub

Private Sub Guilds_UpdateInfo(ByVal Index As Integer)
    
    
    With GuildView.Members(Index)
        lblElv = .Elv
        lblClass.Caption = ListaClases(.Class)
        lblRaze.Caption = ListaRazas(.Raze)
        
        lblPoints.Caption = .Points
        lstRange.ListIndex = .Range
        
        
        Guilds_SetLabels True
    End With
    
End Sub
Private Sub Guilds_SetLabels(ByVal value As Boolean)
    lblElv.Visible = value
    lblClass.Visible = value
    lblRaze.Visible = value
        
    lblPoints.Visible = value
    lstRange.Visible = value
End Sub

Private Sub lstUser_Click()
    If lstUser.ListIndex = -1 Then Exit Sub
    
    Dim A As Long
    
    If lstUser.List(lstUser.ListIndex) = "<VACIO>" Then
        Guilds_SetLabels False
        Exit Sub
    End If
    
    Call Guilds_UpdateInfo(lstUser.ListIndex + 1)
End Sub

