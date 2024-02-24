VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmTravelQuest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   Picture         =   "frmTravelQuest.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstQuest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
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
      Height          =   1785
      Left            =   750
      TabIndex        =   0
      Top             =   1890
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox RecTxtQuest 
      Height          =   1470
      Left            =   4050
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   3330
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   2593
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmTravelQuest.frx":34B7E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgReclamar 
      Height          =   600
      Left            =   8170
      Picture         =   "frmTravelQuest.frx":34BFB
      Top             =   5790
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgAccept 
      Height          =   615
      Left            =   210
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image imgAbandonate 
      Height          =   600
      Left            =   220
      Picture         =   "frmTravelQuest.frx":391C1
      Top             =   6540
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblGld 
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
      Height          =   195
      Left            =   4500
      TabIndex        =   11
      Top             =   6450
      Width           =   3225
   End
   Begin VB.Label lblObj 
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
      Height          =   195
      Index           =   1
      Left            =   4830
      TabIndex        =   10
      Top             =   7050
      Width           =   5115
   End
   Begin VB.Label lblEldhir 
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
      Height          =   195
      Left            =   4710
      TabIndex        =   9
      Top             =   6750
      Width           =   1635
   End
   Begin VB.Label lblExp 
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
      Height          =   195
      Left            =   5130
      TabIndex        =   8
      Top             =   6150
      Width           =   3225
   End
   Begin VB.Label lblPremium 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
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
      Height          =   195
      Left            =   1590
      TabIndex        =   7
      Top             =   5220
      Width           =   555
   End
   Begin VB.Label lblOro 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
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
      Height          =   195
      Left            =   1170
      TabIndex        =   6
      Top             =   4920
      Width           =   555
   End
   Begin VB.Label lblPlata 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
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
      Height          =   195
      Left            =   1260
      TabIndex        =   5
      Top             =   4600
      Width           =   555
   End
   Begin VB.Label lblBronce 
      BackStyle       =   0  'Transparent
      Caption         =   "SI"
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
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   4300
      Width           =   555
   End
   Begin VB.Label lblElv 
      BackStyle       =   0  'Transparent
      Caption         =   "47"
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
      Height          =   195
      Left            =   1230
      TabIndex        =   3
      Top             =   4000
      Width           =   555
   End
   Begin VB.Label lblObj 
      BackStyle       =   0  'Transparent
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
      Height          =   855
      Index           =   0
      Left            =   4110
      TabIndex        =   2
      Top             =   5190
      Width           =   5745
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la misión"
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
      Height          =   585
      Left            =   4080
      TabIndex        =   1
      Top             =   2130
      Width           =   5745
   End
   Begin VB.Image imgHide 
      Height          =   555
      Left            =   10080
      Top             =   450
      Width           =   285
   End
End
Attribute VB_Name = "frmTravelQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call SetWindowLong(RecTxtQuest.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub

Private Sub imgAbandonate_Click()
    If lstQuest.ListIndex = -1 Then Exit Sub
    
    Call WriteTravel_QuestAbandonate(lstQuest.ListIndex + 1)
    
    imgHide_Click
End Sub

Private Sub imgAccept_Click()
    If lstQuest.ListIndex = -1 Then Exit Sub
    
    If imgAbandonate.Visible Then
        Call ShowConsoleMsg("No puedes volver a aceptar la misión. ¡Terminala!")
        Exit Sub
    End If
    
    Call WriteTravel_QuestAccept(lstQuest.ListIndex + 1)
    
    imgHide_Click
End Sub

Private Sub imgHide_Click()
    'frmTravel.Show vbModeless, frmMain
    Unload Me
End Sub

Private Sub imgReclamar_Click()
    If lstQuest.ListIndex = -1 Then Exit Sub
    
    Call WriteTravel_QuestReward(lstQuest.ListIndex + 1)
    
    imgHide_Click
End Sub

Private Sub lstQuest_Click()
    Call WriteTravel_RequiredQuestInfo(lstQuest.ListIndex + 1)
End Sub
