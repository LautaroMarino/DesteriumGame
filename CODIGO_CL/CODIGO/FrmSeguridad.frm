VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form FrmSeguridad 
   BackColor       =   &H80000008&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguridad ExodoAO"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   Icon            =   "FrmSeguridad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Caption         =   "Control de Procesos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2685
      Left            =   240
      TabIndex        =   5
      Top             =   5610
      Width           =   11415
      Begin VB.CommandButton cmbUpdateProcess 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9540
         TabIndex        =   20
         Top             =   2280
         Width           =   1785
      End
      Begin VB.ListBox lstProcess 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1950
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   11175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Caption         =   "Control de Ventanas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1995
      Left            =   240
      TabIndex        =   3
      Top             =   3510
      Width           =   11385
      Begin VB.CommandButton cmbUpdateCaptions 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         TabIndex        =   19
         Top             =   1620
         Width           =   1785
      End
      Begin VB.ListBox lstCaptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1230
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   11175
      End
   End
   Begin VB.Frame f 
      BackColor       =   &H80000008&
      Caption         =   "Velocidades"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2925
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   11355
      Begin VB.CheckBox chkAutomatic 
         BackColor       =   &H80000009&
         Caption         =   "Borrar automáticamente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   3255
      End
      Begin VB.CommandButton cmbSeguir 
         Caption         =   "Dejar de seguir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   21
         Top             =   2520
         Width           =   1785
      End
      Begin VB.CommandButton cmbClearConsole 
         Caption         =   "Limpiar Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9390
         TabIndex        =   18
         Top             =   2460
         Width           =   1785
      End
      Begin VB.CommandButton cmbClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   4800
         TabIndex        =   16
         Top             =   2160
         Width           =   2505
      End
      Begin VB.ListBox lstAttack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   1470
         Left            =   6000
         TabIndex        =   14
         Top             =   480
         Width           =   1425
      End
      Begin VB.ListBox lstSpell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   1470
         Left            =   4440
         TabIndex        =   12
         Top             =   480
         Width           =   1425
      End
      Begin VB.ListBox lstClick 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   1470
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   1305
      End
      Begin VB.ListBox lstU 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   1470
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   1305
      End
      Begin RichTextLib.RichTextBox RecTxt 
         Height          =   1710
         Left            =   7680
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Mensajes del servidor"
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3016
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         TextRTF         =   $"FrmSeguridad.frx":000C
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ataques"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   6240
         TabIndex        =   15
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hechizos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   4440
         TabIndex        =   13
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventario"
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
         Height          =   345
         Left            =   180
         TabIndex        =   11
         Top             =   2430
         Width           =   3225
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Click"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   1920
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   3210
      TabIndex        =   2
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información del personaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "FrmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClear_Click(Index As Integer)
        lstU.Clear
        lstClick.Clear
        lstSpell.Clear
        lstAttack.Clear

End Sub

Private Sub cmbSeguir_Click()
    
    If cmbSeguir.Caption = "Dejar de seguir" Then
        If MsgBox("¿Estás seguro que deseas dejar de seguir al usuario " & lblName.Caption & "?", vbYesNo) = vbYes Then
            cmbSeguir.Caption = "Seguir"
            Call WritePro_Seguimiento(lblName.Caption, False)
        End If

    Else
        cmbSeguir.Caption = "Dejar de seguir"
        Call WritePro_Seguimiento(lblName.Caption, True)
    End If
    
End Sub

Private Sub cmbClearConsole_Click()
    RecTxt.Text = vbNullString
End Sub

Private Sub cmbUpdateCaptions_Click()
    Me.lstCaptions.Clear
    Call WriteSolicitaSeguridad(lblName.Caption, 2)
End Sub

Private Sub cmbUpdateProcess_Click()
    Me.lstProcess.Clear
    Call WriteSolicitaSeguridad(lblName.Caption, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WritePro_Seguimiento(lblName.Caption, False)
End Sub

