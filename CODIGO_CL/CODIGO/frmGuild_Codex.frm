VERSION 5.00
Begin VB.Form frmGuild_Codex 
   BorderStyle     =   0  'None
   Caption         =   "Descripcion del Clan"
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
   Picture         =   "frmGuild_Codex.frx":0000
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
      Index           =   1
      Left            =   1650
      TabIndex        =   3
      Top             =   3840
      Width           =   5950
   End
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   1650
      TabIndex        =   2
      Top             =   4230
      Width           =   5950
   End
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   1650
      TabIndex        =   1
      Top             =   4620
      Width           =   5950
   End
   Begin VB.TextBox txtCodex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   1650
      TabIndex        =   0
      Top             =   5010
      Width           =   5950
   End
   Begin VB.Image imgSetting 
      Height          =   615
      Left            =   6360
      Top             =   6510
      Width           =   1635
   End
   Begin VB.Image imgUnload 
      Height          =   615
      Left            =   7860
      Top             =   450
      Width           =   375
   End
End
Attribute VB_Name = "frmGuild_Codex"
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
    
End Sub

Private Sub imgSetting_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim A As Long
    Dim Codex(1 To MAX_GUILD_CODEX) As String
    
    For A = 1 To MAX_GUILD_CODEX
        Codex(A) = txtCodex(A).Text
            
        If Len(Codex(A)) <= 0 Then
            Call MsgBox("Una de las descripciones está vacía")
            Exit Sub
        End If
    Next A
    
    Call WriteGuilds_Codex("Lautaro Marino" & "-" & Codex(1) & "-" & Codex(2) & "-" & Codex(3) & "-" & Codex(4))
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub
