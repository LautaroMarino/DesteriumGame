VERSION 5.00
Begin VB.Form frmGuild_New 
   BorderStyle     =   0  'None
   Caption         =   "Clanes"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuild_New.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   6
      Top             =   6030
      Width           =   5950
   End
   Begin VB.ComboBox cmbAlineation 
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
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4080
      Width           =   2505
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   2220
      TabIndex        =   4
      Top             =   3720
      Width           =   2505
   End
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   5640
      Width           =   5950
   End
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   5250
      Width           =   5950
   End
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   4860
      Width           =   5950
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
      Left            =   5370
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2730
      Width           =   1920
   End
   Begin VB.Image imgFound 
      Height          =   555
      Left            =   6360
      Top             =   6540
      Width           =   1605
   End
   Begin VB.Image imgClose 
      Height          =   615
      Left            =   7920
      Top             =   420
      Width           =   285
   End
End
Attribute VB_Name = "frmGuild_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MoveForm As clsFormMovementManager

Private Sub Form_Load()
    
    ' Movimiento de pantalla
    Set MoveForm = New clsFormMovementManager
    MoveForm.Initialize Me
    
    'g_Captions(eCaption.eGuildFound) = wGL_Graphic.Create_Device_From_Display(picDraw.hWnd, picDraw.ScaleWidth, picDraw.ScaleHeight)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
   ' Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eGuildFound))
End Sub
Private Sub imgClose_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    MirandoGuildFound = False
    Unload Me
End Sub

Private Sub imgFound_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If MsgBox("Estás a punto de fundar el clan " & txtName.Text & " ¿Es correcto?", vbYesNo, App.Title) = vbYes Then
        
        Dim Name As String
        Dim Codex(1 To MAX_GUILD_CODEX) As String
        Dim Alineation As eGuildAlineation
        
        Dim A As Long
        
        Name = txtName.Text
        Alineation = cmbAlineation.ListIndex + 1
        
        If Len(Name) <= 0 Then
            Call MsgBox("Elige un nombre más largo.")
            Exit Sub
        End If
        
        If Right$(Name, 1) = " " Then
            Name = RTrim$(Name)
            MsgBox "Clan invalido, se han removido los espacios al final del nombre"
        End If
        
        For A = 1 To MAX_GUILD_CODEX
            Codex(A) = txtCodex(A).Text
            
            If Len(Codex(A)) <= 0 Then
                Call MsgBox("Una de las descripciones está vacía")
                Exit Sub
            End If
        Next A
        
        Call WriteGuilds_Found(Name, Alineation, Codex)
        Unload Me
    End If
End Sub

Private Sub txtCodex_Change(Index As Integer)

    If txtCodex(Index).Text = vbNullString Then Exit Sub
    
    If Len(txtCodex(Index)) > MAX_GUILD_LEN_CODEX Then
        txtCodex(Index).Text = Left(txtCodex(Index).Text, Len(txtCodex(Index).Text) - 1)
        txtCodex(Index).SelStart = Len(txtCodex(Index))
    End If
    
    If Not AsciiValidos_Name(txtCodex(Index).Text) Then
        txtCodex(Index).Text = Left(txtCodex(Index).Text, Len(txtCodex(Index).Text) - 1)
        txtCodex(Index).SelStart = Len(txtCodex(Index))
    End If
    
End Sub

Private Sub txtName_Change()

    If txtName.Text = vbNullString Then Exit Sub
    txtName.Text = LTrim$(txtName.Text)
    
    If Len(txtName) > MAX_GUILD_LEN Then
        txtName.Text = Left(txtName.Text, Len(txtName.Text) - 1)
        txtName.SelStart = Len(txtName)
    End If
    
    If Not AsciiValidos_Name(txtName.Text) Then
        txtName.Text = Left(txtName.Text, Len(txtName.Text) - 1)
        txtName.SelStart = Len(txtName)
    End If
End Sub
