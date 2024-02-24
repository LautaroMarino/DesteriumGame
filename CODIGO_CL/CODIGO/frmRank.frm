VERSION 5.00
Begin VB.Form frmRank 
   BorderStyle     =   0  'None
   Caption         =   "Ranking de Usuarios"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   Picture         =   "frmRank.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4980
      Top             =   1950
   End
   Begin VB.ComboBox cmbRanking 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2190
      Width           =   2745
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   9
      Left            =   1050
      TabIndex        =   10
      Top             =   5280
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   8
      Left            =   1050
      TabIndex        =   9
      Top             =   5010
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   7
      Left            =   1050
      TabIndex        =   8
      Top             =   4740
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   6
      Left            =   1050
      TabIndex        =   7
      Top             =   4470
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   5
      Left            =   1050
      TabIndex        =   6
      Top             =   4200
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   4
      Left            =   1050
      TabIndex        =   5
      Top             =   3930
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   3
      Left            =   1050
      TabIndex        =   4
      Top             =   3660
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   2
      Left            =   1050
      TabIndex        =   3
      Top             =   3390
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   1
      Left            =   1050
      TabIndex        =   2
      Top             =   3120
      Width           =   3885
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
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
      Height          =   315
      Index           =   0
      Left            =   1050
      TabIndex        =   1
      Top             =   2850
      Width           =   3885
   End
   Begin VB.Image imgWeb 
      Height          =   315
      Left            =   3990
      Top             =   1590
      Width           =   1005
   End
   Begin VB.Image imgUnload 
      Height          =   525
      Left            =   5730
      Top             =   270
      Width           =   285
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MoveForm As clsFormMovementManager

Private Enum eRank
    eElv = 1
    eFrags = 2
    eRetos1 = 3
    eRetosPlantes = 4
    eDesafios = 5
    eEventos = 6
End Enum
Private Sub cmbRanking_Click()
    
    If Timer1.Enabled Then Exit Sub
    
    Dim ListIndex As Integer
    
    ListIndex = cmbRanking.ListIndex
    
    WriteRequestRanking ListIndex + 1
    
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    ' Movimiento de pantalla
    Set MoveForm = New clsFormMovementManager
    MoveForm.Initialize Me
    
    cmbRanking.AddItem "Top Nivel"
    cmbRanking.AddItem "Top Frags"
    cmbRanking.AddItem "Top Retos"
    cmbRanking.AddItem "Top Plantes"
    cmbRanking.AddItem "Top Desafios"
    cmbRanking.AddItem "Top Eventos"
    
    Dim A As Long
    
    For A = lblUser.LBound To lblUser.UBound
        lblUser(A).Caption = "<Disponible>"
    Next A
    
End Sub


Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub

Private Sub imgWeb_Click()
    Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/rank/", vbNullString, vbNullString, 1)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub
