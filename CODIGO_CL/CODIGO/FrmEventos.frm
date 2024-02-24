VERSION 5.00
Begin VB.Form FrmEventos 
   BackColor       =   &H80000008&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Eventos automáticos"
   ClientHeight    =   10875
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   10395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmEventos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H80000007&
      Caption         =   "EVENTOS EN CURSO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4335
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      Begin VB.Frame Frame2 
         BackColor       =   &H80000007&
         Caption         =   "Info del evento seleccionado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   2805
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4065
         Begin VB.ListBox lstUser 
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
            ForeColor       =   &H80000005&
            Height          =   1710
            Left            =   1800
            TabIndex        =   10
            Top             =   360
            Width           =   2145
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REENVIAR REGLAS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   73
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label lblEnabled 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO EMPEZO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   90
            TabIndex        =   14
            Top             =   300
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Label lblKick 
            AutoSize        =   -1  'True
            BackColor       =   &H00000040&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "/KICK"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   3120
            TabIndex        =   13
            Top             =   2400
            Width           =   705
         End
         Begin VB.Label lblSum 
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "/SUM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   2280
            TabIndex        =   12
            Top             =   2400
            Width           =   705
         End
         Begin VB.Label lblInfoEvent 
            BackStyle       =   0  'Transparent
            Caption         =   "Info no disponbile"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   1725
            Left            =   90
            TabIndex        =   11
            Top             =   570
            Visible         =   0   'False
            Width           =   2250
         End
      End
      Begin VB.ListBox lstEventos 
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
         ForeColor       =   &H80000005&
         Height          =   1230
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label lblCloseEvent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CERRAR"
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
         Height          =   240
         Left            =   2760
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACTUALIZAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   2640
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtApu 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   4920
         TabIndex        =   109
         Text            =   "5"
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbQuotas 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   1
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmbMaxArenas 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cmbArenas 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   1
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cmbArenas 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   0
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cmbRounds 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   1
         ItemData        =   "FrmEventos.frx":000C
         Left            =   3120
         List            =   "FrmEventos.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbRounds 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   0
         ItemData        =   "FrmEventos.frx":0034
         Left            =   1200
         List            =   "FrmEventos.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtHP 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1560
         TabIndex        =   86
         Text            =   "0"
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtRojas 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1800
         TabIndex        =   83
         Text            =   "0"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtRequiredGld 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1440
         TabIndex        =   82
         Text            =   "0"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtRequiredEldhir 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   3840
         TabIndex        =   81
         Text            =   "0"
         Top             =   3600
         Width           =   495
      End
      Begin VB.ComboBox cmbMaxElv 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2760
         Width           =   795
      End
      Begin VB.ComboBox cmbMinElv 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2760
         Width           =   795
      End
      Begin VB.ComboBox cmbTeam 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cmbQuotas 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Index           =   0
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cmbTipo 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prob Asesino:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   210
         Index           =   20
         Left            =   3480
         TabIndex        =   108
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   2520
         TabIndex        =   104
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   3960
         TabIndex        =   102
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label lblInfoArenas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   5280
         TabIndex        =   101
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ARENAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   98
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final a:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   2
         Left            =   2280
         TabIndex        =   96
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rounds:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   93
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DSP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   18
         Left            =   4440
         TabIndex        =   89
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORO y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   17
         Left            =   3240
         TabIndex        =   88
         Top             =   3600
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIDA INICIAL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   87
         Top             =   3240
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Rojas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   85
         Top             =   2280
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscripción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   10
         Left            =   240
         TabIndex        =   84
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   13
         Left            =   2280
         TabIndex        =   72
         Top             =   2880
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite Nivel:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   71
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de evento:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MiN Cupos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x Team"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   645
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000005&
      Caption         =   "Extra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   6495
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   10215
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir 'Fuego Amigo'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   23
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   111
         Top             =   5640
         Value           =   1  'Checked
         Width           =   3585
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Uso de 'Cascos y Escudos'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   22
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   110
         Top             =   5400
         Value           =   1  'Checked
         Width           =   3585
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clases"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2985
         Left            =   4560
         TabIndex        =   22
         Top             =   2520
         Width           =   1995
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Trabajador"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   160
            TabIndex        =   94
            Top             =   2520
            Width           =   1425
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Trabajador"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   150
            TabIndex        =   31
            Top             =   2280
            Width           =   1425
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Cazador"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   150
            TabIndex        =   30
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1365
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Paladin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   150
            TabIndex        =   29
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Druida"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   150
            TabIndex        =   28
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1245
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Bardo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   150
            TabIndex        =   27
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1365
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Asesino"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   150
            TabIndex        =   26
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Guerrero"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   150
            TabIndex        =   25
            Top             =   810
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Clerigo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   24
            Top             =   540
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkClass 
            BackColor       =   &H8000000B&
            Caption         =   "Mago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   23
            Top             =   270
            Value           =   1  'Checked
            Width           =   1305
         End
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Teletransportar usuarios constantemente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   21
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   91
         Top             =   1320
         Visible         =   0   'False
         Width           =   4110
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Ajustar Cupos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   90
         Top             =   2280
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000007&
         Caption         =   "Tiempos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1485
         Left            =   2760
         TabIndex        =   74
         Top             =   240
         Width           =   3675
         Begin VB.ComboBox cmbCierran 
            BackColor       =   &H80000007&
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
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   310
            Width           =   855
         End
         Begin VB.ComboBox cmbAbren 
            BackColor       =   &H80000007&
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   310
            Width           =   855
         End
         Begin VB.ComboBox cmbDuration 
            BackColor       =   &H80000007&
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inscripciones:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   15
            Left            =   2400
            TabIndex        =   79
            Top             =   400
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   16
            Left            =   120
            TabIndex        =   78
            Top             =   780
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daño de Zona"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   68
         Top             =   1920
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Hechizos de Curacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Index           =   15
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   67
         Top             =   3960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solo Daga Comun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   14
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   66
         Top             =   3720
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Uso de 'Tormenta de Fuego"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   20
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   65
         Top             =   5160
         Value           =   1  'Checked
         Width           =   3585
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Uso de 'Descarga Eléctrica'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   19
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   64
         Top             =   4920
         Value           =   1  'Checked
         Width           =   3585
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Uso de 'Apocalipsis' y 'Explosion Abismal'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   18
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   63
         Top             =   4680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Paralizar/Inmovilizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   17
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   62
         Top             =   4440
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitir Uso de Pociones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   16
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   61
         Top             =   4200
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mezclar Apariencias de los personajes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   105
         MaskColor       =   &H00000000&
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vale INVOCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   12
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   55
         Top             =   3240
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vale INVISIBILIDAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   11
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   54
         Top             =   3000
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vale OCULTAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Index           =   10
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   53
         Top             =   2760
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vale RESUCITAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Index           =   9
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   52
         Top             =   720
         Width           =   1665
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000007&
         Caption         =   "Cambiar clase, raza y nivel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   3255
         Left            =   6600
         TabIndex        =   46
         Top             =   2160
         Width           =   2865
         Begin VB.ComboBox cmbChangeLevel 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   1680
            Width           =   795
         End
         Begin VB.ComboBox cmbRaza 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1080
            Width           =   1725
         End
         Begin VB.ComboBox cmbClass 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   480
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   19
            Left            =   240
            TabIndex        =   106
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Raza:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   4
            Left            =   150
            TabIndex        =   49
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000007&
         Caption         =   "Premios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1965
         Left            =   6600
         TabIndex        =   39
         Top             =   240
         Width           =   2715
         Begin VB.ComboBox cmbPrizeGld 
            BackColor       =   &H80000007&
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
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cmbPrizeEldhir 
            BackColor       =   &H80000007&
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
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtPrizeObjCant 
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   2280
            TabIndex        =   41
            Text            =   "1"
            Top             =   1110
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtPrizeObj 
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   840
            TabIndex        =   40
            Text            =   "0"
            Top             =   1110
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cant"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   9
            Left            =   1800
            TabIndex        =   45
            Top             =   1200
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DSP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   8
            Left            =   240
            TabIndex        =   44
            Top             =   750
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   7
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Visible         =   0   'False
            Width           =   630
         End
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solo Grupos/Partys"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   38
         Top             =   240
         Width           =   1995
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Inventario vacío"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   37
         Top             =   480
         Width           =   1875
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Items [BRONCE]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   4515
         MaskColor       =   &H00000000&
         TabIndex        =   36
         Top             =   315
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solo Clanes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Items [PLATA]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   4515
         MaskColor       =   &H00000000&
         TabIndex        =   34
         Top             =   550
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Items [ORO]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   4515
         MaskColor       =   &H00000000&
         TabIndex        =   33
         Top             =   800
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox chkConfig 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Items [PREMIUM]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Index           =   3
         Left            =   4515
         TabIndex        =   32
         Top             =   1050
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000007&
         Caption         =   "Facciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1485
         Left            =   9360
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
         Begin VB.CheckBox chkFaction 
            BackColor       =   &H8000000B&
            Caption         =   "Armada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.CheckBox chkFaction 
            BackColor       =   &H8000000B&
            Caption         =   "Legion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox chkFaction 
            BackColor       =   &H8000000B&
            Caption         =   "Ciudadano"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.CheckBox chkFaction 
            BackColor       =   &H8000000B&
            Caption         =   "Criminal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Value           =   1  'Checked
            Width           =   1155
         End
      End
      Begin VB.Label lblUnload 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CERRAR VENTANA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   3240
         TabIndex        =   95
         Top             =   2280
         Width           =   2850
      End
      Begin VB.Label lblNuevo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ENVIAR EVENTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   405
         Left            =   3480
         TabIndex        =   51
         Top             =   1800
         Width           =   2850
      End
   End
End
Attribute VB_Name = "FrmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkChangeUser_Click()

    If chkChangeUser.Value = 1 Then
        chkInven.Value = 1
    End If

End Sub

Private Sub chkInven_Click()
    If cmbTipo.ListIndex = -1 Then Exit Sub
    
    If cmbTipo.Text = "DAGARUSA" Or cmbTipo.Text = "JDH" Then
        chkInven.Value = 1
    End If
    
    If cmbClass.ListIndex > 0 Then
        chkInven.Value = 1
    End If
    
    If cmbRaza.ListIndex > 0 Then
        chkInven.Value = 1
    End If
End Sub

Private Sub chkParejas_Click()
    If cmbTipo.Text = "DUELOS" Then
        If txtTeam.Text = 1 Then
            chkParejas.Value = 0
        End If
    End If
End Sub

Private Sub cmbClass_Click()
    If cmbClass.ListIndex > 0 Then
        chkConfig(eConfigEvent.eInvFree) = 1
    End If
End Sub

Private Sub cmbRaza_Change()
    If cmbRaza.ListIndex > 0 Then
         chkConfig(eConfigEvent.eInvFree) = 1
    End If
End Sub

Private Sub cmbTeam_Click()
    Call Events_SettingQuotas(cmbTipo.Text)
    cmbQuotas(0).ListIndex = 0
    cmbQuotas(1).ListIndex = 2
End Sub

Private Sub cmbTipo_Click()
    If cmbTipo.ListIndex = -1 Then Exit Sub
    
    
    If cmbTipo.Text = "DAGARUSA" Or cmbTipo.Text = "JDH" Or cmbClass.ListIndex > 0 Or cmbRaza.ListIndex > 0 Or cmbChangeLevel.ListIndex > 0 Then
        chkConfig(eConfigEvent.eInvFree) = 1
    Else
        chkConfig(eConfigEvent.eInvFree) = 0
    End If
    
        
    If cmbTipo.Text = "DEATHMATCH" Then
        chkConfig(eConfigEvent.eTeletransportacion) = 1
        chkConfig(eConfigEvent.eMezclarApariencias) = 1
    Else
        chkConfig(eConfigEvent.eTeletransportacion) = 0
        chkConfig(eConfigEvent.eMezclarApariencias) = 0
    End If
    
    
    txtName.Text = cmbTipo.Text
    Events_SettingQuotas cmbTipo.Text
End Sub

Private Sub Events_SettingQuotas(ByVal Text As String)

    On Error GoTo ErrHandler
    
    Dim A    As Long

    Dim Temp As Long

    Dim Mult As Long
    
    
    
    For A = cmbQuotas.LBound To cmbQuotas.UBound
        cmbQuotas(A).Clear
        
        Select Case Text

            Case "DUELOS"
                txtName.Text = cmbTeam.Text & "vs" & cmbTeam.Text
           
                Select Case Val(cmbTeam.Text)

                    Case 1
                        cmbQuotas(A).AddItem "4"
                        cmbQuotas(A).AddItem "8"
                        cmbQuotas(A).AddItem "16"
                        cmbQuotas(A).AddItem "32"
                        cmbQuotas(A).AddItem "64"
                        cmbQuotas(A).AddItem "128"

                    Case 2
                        cmbQuotas(A).AddItem "4"
                        cmbQuotas(A).AddItem "8"
                        cmbQuotas(A).AddItem "16"
                        cmbQuotas(A).AddItem "32"
                        cmbQuotas(A).AddItem "64"
                        cmbQuotas(A).AddItem "128"

                    Case 3
                        cmbQuotas(A).AddItem "6"
                        cmbQuotas(A).AddItem "12"
                        cmbQuotas(A).AddItem "24"
                        cmbQuotas(A).AddItem "48"
                        cmbQuotas(A).AddItem "96"

                    Case 4
                        cmbQuotas(A).AddItem "8"
                        cmbQuotas(A).AddItem "16"
                        cmbQuotas(A).AddItem "32"
                        cmbQuotas(A).AddItem "64"
                        cmbQuotas(A).AddItem "128"

                    Case 5
                        cmbQuotas(A).AddItem "10"
                        cmbQuotas(A).AddItem "20"
                        cmbQuotas(A).AddItem "30"
                        cmbQuotas(A).AddItem "40"
                        cmbQuotas(A).AddItem "50"
                        cmbQuotas(A).AddItem "60"
                        cmbQuotas(A).AddItem "70"
                        cmbQuotas(A).AddItem "80"
                        cmbQuotas(A).AddItem "90"
                        cmbQuotas(A).AddItem "100"
                
                End Select
            
            Case Else
                cmbQuotas(A).AddItem "2"
                cmbQuotas(A).AddItem "4"
                cmbQuotas(A).AddItem "5"
                cmbQuotas(A).AddItem "6"
                cmbQuotas(A).AddItem "7"
                cmbQuotas(A).AddItem "8"
                cmbQuotas(A).AddItem "9"
                cmbQuotas(A).AddItem "10"
                cmbQuotas(A).AddItem "11"
                cmbQuotas(A).AddItem "12"
                cmbQuotas(A).AddItem "13"
                cmbQuotas(A).AddItem "14"
                cmbQuotas(A).AddItem "15"
                cmbQuotas(A).AddItem "20"
                cmbQuotas(A).AddItem "30"
                cmbQuotas(A).AddItem "40"
                cmbQuotas(A).AddItem "50"
                cmbQuotas(A).AddItem "60"
                cmbQuotas(A).AddItem "70"
                cmbQuotas(A).AddItem "80"
                cmbQuotas(A).AddItem "90"
                cmbQuotas(A).AddItem "100"
        
        End Select
        
        
        
        cmbQuotas(A).ListIndex = 0
    Next A
    
    
    
    Exit Sub
ErrHandler:

End Sub

Private Sub Form_Load()
    
    cmbTipo.AddItem "REY"
    cmbTipo.AddItem "DAGARUSA"
    cmbTipo.AddItem "DEATHMATCH"
    cmbTipo.AddItem "DUELOS"
    cmbTipo.AddItem "TELEPORTS"
    cmbTipo.AddItem "GRANBESTIA"
    cmbTipo.AddItem "BUSQUEDA"
    cmbTipo.AddItem "IMPARABLE"
    cmbTipo.AddItem "JDH"
    cmbTipo.AddItem "MANUAL"
    
    
    Dim A As Long
    
    cmbClass.AddItem "Seleccionar"
    cmbRaza.AddItem "Seleccionar"
    
    For A = 1 To NUMCLASES
        cmbClass.AddItem ListaClases(A)
        chkClass(A).Caption = ListaClases(A)
    Next A
    
    For A = 1 To NUMRAZAS
        cmbRaza.AddItem ListaRazas(A)
        
        cmbTeam.AddItem A
    Next A
    
    cmbClass.ListIndex = 0
    cmbRaza.ListIndex = 0
    cmbTeam.ListIndex = 0
    
    For A = 0 To 47
        cmbChangeLevel.AddItem A
    Next A
    
    For A = 1 To 47
        cmbMinElv.AddItem A
        cmbMaxElv.AddItem A
    Next A
    
    
    For A = 1 To 5
        cmbMaxArenas.AddItem A
    Next A
    
    For A = 1 To 16
        cmbArenas(0).AddItem A
        cmbArenas(1).AddItem A
    Next A
    
    cmbPrizeGld.AddItem "0"
    cmbPrizeGld.AddItem "25000"
    cmbPrizeGld.AddItem "75000"
    cmbPrizeGld.AddItem "100000"
    cmbPrizeGld.AddItem "125000"
    cmbPrizeGld.AddItem "150000"
    cmbPrizeGld.AddItem "175000"
    cmbPrizeGld.AddItem "200000"
    cmbPrizeGld.AddItem "275000"
    cmbPrizeGld.AddItem "300000"
    cmbPrizeGld.AddItem "375000"
    cmbPrizeGld.AddItem "400000"
    cmbPrizeGld.AddItem "475000"
    cmbPrizeGld.AddItem "500000"
    cmbPrizeGld.AddItem "575000"
    cmbPrizeGld.AddItem "600000"
    cmbPrizeGld.AddItem "675000"
    cmbPrizeGld.AddItem "700000"
    cmbPrizeGld.AddItem "1000000"
    
    For A = 0 To 10
        cmbPrizeEldhir.AddItem A
    Next A
    
    cmbPrizeEldhir.ListIndex = 0
    cmbPrizeGld.ListIndex = 0
    
    cmbMinElv.ListIndex = 24
    cmbMaxElv.ListIndex = 46
    cmbChangeLevel.ListIndex = 0
    
    For A = 0 To 20
        cmbAbren.AddItem A
        cmbDuration.AddItem A
    Next A
    
    For A = 1 To 10
        cmbCierran.AddItem A
    Next A
    
    cmbDuration.ListIndex = 0
    cmbAbren.ListIndex = 1
    cmbCierran.ListIndex = 0
    cmbRounds(0).ListIndex = 0
    cmbRounds(1).ListIndex = 0
    cmbTeam.ListIndex = 0
    cmbMaxArenas.ListIndex = 3
    cmbArenas(0).ListIndex = 0
    cmbArenas(1).ListIndex = 3
    
End Sub

Private Function CheckAll() As Boolean
    
    If cmbTipo.ListIndex < 0 Then
        MsgBox "Selecciona el tipo de evento"

        Exit Function

    End If
    
    If cmbTipo.Text = "MANUAL" Then
        If txtName.Text = vbNullString Or Len(txtName.Text) <= 4 Then
            MsgBox ("Utiliza un nombre para el evento más largo")
            Exit Function

        End If
    
    End If
        
    If Val(cmbAbren.Text) <= 0 Then
        MsgBox "Tiempo de abrir inscripciones inválido"

        Exit Function

    End If
      
    If Val(cmbCierran.Text) <= 0 Then
        MsgBox "Tiempo de cerrar inscripciones inválido. Pone más de 10"

        Exit Function

    End If
    
    If Val(txtPrizeObj.Text) < 0 Then
        MsgBox ("El objeto que has elegido, es inválido.")

        Exit Function

    End If
        
    If Val(txtPrizeObj.Text) > 0 Then
        If Val(txtPrizeObjCant.Text) <= 0 Then
            MsgBox ("Si deseas entregar objeto como premio del evento, chequea que su cantidad sea al menos 1")

            Exit Function

        End If

    End If

    If Val(txtRojas.Text) < 0 Then
        MsgBox "Debes pedir un número válido de pociones 0-10000"
        Exit Function
    
    End If
    
    If cmbClass.ListIndex = 0 And cmbRaza.ListIndex > 0 Then
        MsgBox "Debes seleccionar una clase además de la raza."
        Exit Function

    End If
    
    If cmbRaza.ListIndex = 0 And cmbClass.ListIndex > 0 Then
        MsgBox "Debes seleccionar una raza además de la clase."
        Exit Function

    End If
    
    If Val(cmbArenas(0).Text) > Val(cmbArenas(1).Text) Then
        MsgBox "El minimo de arena debe ser menor al máximo."
        Exit Function

    End If
    
    If cmbClass.ListIndex > 0 Or cmbRaza.ListIndex Or cmbChangeLevel.ListIndex > 0 Then
        If (txtRojas.Text) <= 0 Then
            MsgBox "Necesitas poner un Limite de Rojas, que funcionará para que les de esa cantidad de rojas a los usuarios."
            Exit Function

        End If

    End If
    
    If Val(cmbRounds(1).Text) < Val(cmbRounds(0).Text) Then
        MsgBox "El máximo de cupos no puede ser menor que el mínimo de cupos."
        Exit Function

    End If
    
    If cmbTipo.Text = "DAGARUSA" Then
        If Val(txtApu.Text) <= 0 Or Val(txtApu.Text) > 100 Then
            MsgBox "Mayor a cero o menor a 100 amigo"
            Exit Function

        End If

    End If
    
    CheckAll = True

End Function


Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label4_Click()
    lstUser.Clear
    lblInfoEvent.Visible = False
    Protocol.WriteRequiredEvents
End Sub

Private Sub lblCloseEvent_Click()

    If lstEventos.ListIndex = -1 Then Exit Sub
    

    
    If MsgBox("¿Estás seguro que deseas cerrar el evento seleccionado?", vbYesNo) = vbYes Then
        lblEnabled.Visible = False
        lstUser.Clear
        lblInfoEvent.Visible = False
        WriteCloseEvent lstEventos.ListIndex + 1
        Protocol.WriteRequiredEvents
    End If
End Sub

Private Sub lblInfoArenas_Click()

    Dim Temp As String
    
    Temp = "Arenas de 1 a 4= Arenas comunes"
    Temp = Temp & vbCrLf & "Arenas de 5 a 16= Arenas con texturas de dungeons variables."
    Temp = Temp & vbCrLf & "Si no saben usarlo pongan (suponiendo A-B, dos puuntos para generar randoms) A-B=1-4 o A-B=5-16"
    Call MsgBox(Temp)
End Sub

Private Sub lblKick_Click()
    If lstUser.ListIndex = -1 Then Exit Sub
    If lstUser.List(lstUser.ListIndex) = vbNullString Or lstUser.List(lstUser.ListIndex) = "(VACIO)" Then Exit Sub
    
    Call WriteEvents_KickUser(lstUser.List(lstUser.ListIndex))
End Sub

Private Sub lblNuevo_Click()

    Dim AllowedClasses(1 To NUMCLASES) As Byte

    Dim AllowedFaction(1 To 4)         As Byte
          
    Dim LoopC                          As Byte
          
    If CheckAll() Then
        Dim Events As tEvents
        
        With Events
            ReDim .AllowedClass(1 To NUMCLASES) As Byte
            ReDim .AllowedFaction(1 To 4) As Byte
            
            For LoopC = 1 To (NUMCLASES)
                .AllowedClass(LoopC) = chkClass(LoopC).Value
            Next LoopC
    
            For LoopC = 1 To 4
                .AllowedFaction(LoopC) = chkFaction(LoopC).Value
            Next LoopC
            
            .Name = txtName.Text
            .Modality = cmbTipo.ListIndex + 1
            .QuotasMin = Val(cmbQuotas(0).Text)
            .QuotasMax = Val(cmbQuotas(1).Text)
            .MinLvl = Val(cmbMinElv.Text)
            .MaxLvl = Val(cmbMaxElv.Text)
            .InscriptionGld = Val(txtRequiredGld.Text)
            .InscriptionEldhir = Val(txtRequiredEldhir.Text)
            .TimeInit = Val(cmbAbren.Text) * 60
            .TimeCancel = Val(cmbCierran.Text) * 60
            .TeamCant = Val(cmbTeam.Text)
            
            For LoopC = 0 To MAX_EVENTS_CONFIG - 1
                .Config(LoopC) = chkConfig(LoopC).Value
            Next LoopC

            .ChangeClass = cmbClass.ListIndex
            .ChangeRaze = cmbRaza.ListIndex

            .LimitRed = Val(txtRojas.Text)
            .PrizeGld = Val(cmbPrizeGld.Text)
            .PrizeEldhir = Val(cmbPrizeEldhir.Text)
            .PrizeObj = Val(txtPrizeObj.Text)
            .PrizeObj_Amount = Val(txtPrizeObjCant.Text)
            .LimitRound = Val(cmbRounds(0).Text)
            .LimitRoundFinal = Val(cmbRounds(1).Text)
            .LimitArenas = Val(cmbMaxArenas.Text)
            .ArenasMin = Val(cmbArenas(0).Text)
            .ArenasMax = Val(cmbArenas(1).Text)
            .ChangeLevel = Val(cmbChangeLevel.Text)
            .ProbApu = Val(txtApu.Text)
        End With
        
        Call WriteNewEvent(Events)
        Unload Me
    End If
        
End Sub

Private Sub lblSum_Click()
    If lstUser.ListIndex = -1 Then Exit Sub
    If lstUser.List(lstUser.ListIndex) = vbNullString Or lstUser.List(lstUser.ListIndex) = "(VACIO)" Then Exit Sub
    
    Call WriteSummonChar(lstUser.List(lstUser.ListIndex), True)
End Sub

Private Sub lblUnload_Click()
    Unload Me
End Sub

Private Sub lstEventos_Click()
    lblCloseEvent.Visible = IIf((lstEventos.List(lstEventos.ListIndex) <> "Vacio"), True, False)
          
    If lstEventos.List(lstEventos.ListIndex) = "Vacio" Then
        lblEnabled.Visible = False
        lstUser.Clear
        lblInfoEvent.Visible = False
        Exit Sub

    End If
          
    Protocol.WriteRequiredDataEvent lstEventos.ListIndex + 1
End Sub

Private Sub txtTeam_Change()
    If Len(txtTeam.Text) <= 0 Then Exit Sub
    
    If cmbTipo.Text = "DUELOS" Then
        If txtTeam.Text = 1 Then
            chkParejas.Value = 0
        End If
    End If
End Sub
