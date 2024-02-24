VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmManual 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmManual.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicMenu 
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
      Height          =   2040
      Left            =   4380
      ScaleHeight     =   136
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   25
      Top             =   3390
      Visible         =   0   'False
      Width           =   3360
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volver"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1200
         TabIndex        =   30
         Top             =   5880
         Width           =   990
      End
      Begin VB.Label lblGemas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gemas con poderes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   630
         TabIndex        =   29
         Top             =   420
         Width           =   2160
      End
      Begin VB.Label lblCajas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cajas Dragons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   840
         TabIndex        =   28
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblEntrenamiento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTRENAMIENTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   750
         TabIndex        =   27
         Top             =   1740
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reliquias del Dragón"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   270
         Left            =   570
         TabIndex        =   26
         Top             =   0
         Width           =   2325
      End
   End
   Begin RichTextLib.RichTextBox Console 
      Height          =   6825
      Left            =   3420
      TabIndex        =   1
      Top             =   1860
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   12039
      _Version        =   393217
      BackColor       =   -2147483641
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmManual.frx":1D212
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rey Inmortal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   25
      Left            =   210
      TabIndex        =   33
      Top             =   7100
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clan vs Clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   24
      Left            =   210
      TabIndex        =   32
      Top             =   6900
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   23
      Left            =   210
      TabIndex        =   31
      Top             =   6700
      Width           =   2745
   End
   Begin VB.Label lblExtras 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXTRAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   390
      TabIndex        =   24
      Top             =   8220
      Width           =   2205
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Poder BONUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   22
      Left            =   210
      TabIndex        =   23
      Top             =   6480
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Poder SUPREMO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   21
      Left            =   210
      TabIndex        =   22
      Top             =   6270
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Angeles y Demonios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   20
      Left            =   210
      TabIndex        =   21
      Top             =   6060
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios <LEGENDARIO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   19
      Left            =   210
      TabIndex        =   20
      Top             =   5850
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios <PLATINO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   18
      Left            =   210
      TabIndex        =   19
      Top             =   5640
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios <ORO>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   17
      Left            =   210
      TabIndex        =   18
      Top             =   5430
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fortaleza del Rey Dragon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   16
      Left            =   210
      TabIndex        =   17
      Top             =   5220
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Castillo de Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   15
      Left            =   210
      TabIndex        =   16
      Top             =   5010
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reliquias del Dragon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   14
      Left            =   210
      TabIndex        =   15
      Top             =   4800
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ranking Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   13
      Left            =   210
      TabIndex        =   14
      Top             =   4590
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ranking"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   12
      Left            =   210
      TabIndex        =   13
      Top             =   4380
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invocaciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   11
      Left            =   210
      TabIndex        =   12
      Top             =   4170
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misiones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   10
      Left            =   210
      TabIndex        =   11
      Top             =   3960
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gemas Caballero Dragons"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   9
      Left            =   210
      TabIndex        =   10
      Top             =   3750
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cajas Dragons"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   8
      Left            =   210
      TabIndex        =   9
      Top             =   3540
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Desafios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   7
      Left            =   210
      TabIndex        =   8
      Top             =   3330
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eventos automáticos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   6
      Left            =   210
      TabIndex        =   7
      Top             =   3120
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   5
      Left            =   210
      TabIndex        =   6
      Top             =   2940
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dinero en el juego"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   4
      Left            =   210
      TabIndex        =   5
      Top             =   2730
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Viaje al mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   2520
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comenzando"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   2
      Left            =   210
      TabIndex        =   3
      Top             =   2310
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Personaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   2100
      Width           =   2745
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   1890
      Width           =   2745
   End
   Begin VB.Image imgUnload 
      Height          =   465
      Left            =   9000
      Top             =   300
      Width           =   465
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub imgUnload_Click()
    Unload Me
End Sub

Private Sub lblSection_Selected(ByVal Selected As Integer)

    Dim A As Long
    
    Console.Visible = True
    Console.Text = vbNullString
    PicMenu.Visible = False
    
    For A = lblSection.LBound To lblSection.UBound
        lblSection(A).ForeColor = vbWhite
    Next A
    
    lblSection(Selected).ForeColor = vbYellow
End Sub

Private Sub Label2_Click()
    FrmReliquias.Show vbModeless, frmMain
End Sub

Private Sub lblCajas_Click()
    FrmCajas.Show vbModeless, frmMain
End Sub

Private Sub lblEntrenamiento_Click()
    FrmEntrenamiento.Show vbModeless, frmMain
End Sub

Private Sub lblExtras_Click()
    Console.Visible = False
    PicMenu.Visible = True
End Sub

Private Sub lblGemas_Click()
    frmGemas.Show vbModeless, frmMain
End Sub

Private Sub lblSection_Click(Index As Integer)
    
    lblSection_Selected (Index)
    
    Select Case Index

        Case 0  ' Crear cuenta
            
            Call AddtoRichTextBox(Console, "Crear una cuenta para guardar tus personajes te será útil para no tener que recordar más de una contraseña.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "No utilices contraseñas de redes sociales, ni de otros juegos (Por tu seguridad)", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El staff no se hace responsable de pérdida de objetos ni baneos de personajes.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Cada usuario se hace responsable de su cuenta y de los datos de la misma.", 255, 255, 255, True)
        
        Case 1 ' Crear Personaje
            Call AddtoRichTextBox(Console, "Clases sociales: Mago, Druida, Bardo, Clérigo, Paladin, Asesino, Cazador, Guerrero.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Razas: Humano, Elfo, Elfo Drow, Gnomo, Enano.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Facción: Podrás elegir entre Legión Oscura y Armada real. No podrás cambiar de facción luego.", 255, 255, 255, True)
        
        Case 2 ' Comenzando
            Call AddtoRichTextBox(Console, "La ciudad principal se considera ULLATHORPE.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Te otorgamos un SET INICIAL, el cual te ayudará durante todo el entrenamiento de tu personaje. Este equipo no se caerá.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Todas las clases podrán realizar trabajos.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Tu personaje posee características para poder luchar contra otros personajes.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Para saber si tu personaje está por encima del promedio podrás consultarlo con el comando /EST", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Algunos niveles recibirás recompensas EXTRAS.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Nivel n°10: Todas las clases reciben la Armadura +10 la cual aporta un beneficio extra en defensa.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Nivel n°11: Hechizo o Espada para causar más daño a las criaturas no hostiles.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Nivel n°15: Desbloqueo de [ANGEL] o [DEMONIO] según facción escogida.", 255, 255, 255, True)
            
        Case 3 'Viaje al mundo
            Call AddtoRichTextBox(Console, "Desde la ciudad Principal podrás acceder a todos los mapas del juego. El viajero está ubicado en el puerto.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Al hacerle doble clic nos mostrará la lista de mapas donde podemos circular, además de una breve descripción de cada uno para ser de ayuda.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Nos dirá las criaturas que hay en los mapas y también los drops de cada una.", 255, 255, 255, True)
    
        Case 4 ' Dinero en el juego
            Call AddtoRichTextBox(Console, "La moneda principal se llama Diamante Rojo. Las criaturas vendedoras de la ciudad te venderán la mayoría de los objetos a cambio de estos diamantes.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "La moneda secundaria y no menos importante es el Diamante Azul. Es considerada la moneda PREMIUM del juego. Con ésta podrás comprar los mejores objetos, que no todos los personajes lográn conseguir.", 255, 255, 255, True)
        
        Case 5 ' Retos
            Call AddtoRichTextBox(Console, "Podrás luchar contra otros usuarios a cambio de Diamantes Rojos y diamantes Azules.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los enfrentamientos son 1vs1 y 2vs2.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "En estos enfrentamientos no caen los objetos.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Con el comando /RETOS se abrirá el PANEL.", 255, 255, 255, True)
    
        Case 6 ' Eventos automáticos
            Call AddtoRichTextBox(Console, "Los Game Master organizarán eventos de forma diaria. Todos los eventos se ingresan con el mismo comando /ENTRAR", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Dependiendo el tipo de evento que sea el comando varia a /ENTRAR 1VS1 (Por ejemplo)", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Con cada evento automático que ganes  recibirás diamantes azules.", 255, 255, 255, True)
        
        Case 7 ' Desafios
            Call AddtoRichTextBox(Console, "¿Te sientes el mejor de las Tierras de Dragones? Podrás demostrarlo en los desafios", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "/DESAFIO ¡Necesitaras 3 diamantes Azules!", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "/SALIRDESAFIO ¡Saldrás del desafio pero perderás los diamantes!", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Cada ciertos combates ganados el personaje recibe recompensas y reconocimientos", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "5 Desafios Ganados: 5 Diamantes Azules", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "10 Desafios Ganados: 10 Diamantes Azules", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "15 Desafios Ganados: 20 Diamantes Azules", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "20 Desafios Ganados: 30 Diamantes Azules", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "25 Desafios Ganados: 50 Diamantes Azules", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "50 Desafios Ganados: 100 Diamantes Azules", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "100 Desafios Ganados: 500 Diamantes Azules", 255, 255, 255, True)
        
        Case 8 ' Cajas Dragons
            Call AddtoRichTextBox(Console, "Con el nuevo sistema de pesca además de conseguir diamantes rojos vendiendo los pescados obtenidos podrás recoger cajas Dragons desde la profundidad del océano.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Hay distintos tipos de Cajas y cada puede arrojar objetos, pero también pueden no darte algo. Cada caja tiene su propia dificultad", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Caja Verde: 10%", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Caja Violeta: 10%", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Caja Roja: 3%", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Caja Celeste: 2%", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Caja Amarilla: 1%", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Para saber los objetos específicos que da cada CAJA, puedes consultarlo desde la sección 'EXTRAS' de este manual.", 255, 255, 255, True)
        
        Case 9 'Gemas del Caballero Dragons
            Call AddtoRichTextBox(Console, "Las gemas te otorgarán distintos poderes que benefician a tu personaje. Actualmente hay 8 gemas disponibles y su poder se activa haciendo doble clic. Al deslogear el efecto de la gema permanecerá activado.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Solo se permite utilizar una gema por personaje.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Para ver el efecto que causa cada gema en tu personaje podrás consultarlo en la sección 'EXTRAS' de este manual.", 255, 255, 255, True)
        
        Case 10 ' Misiones
            Call AddtoRichTextBox(Console, "Podrás realizar distintas misiones para ir desbloqueando logros con tu personaje. Las primeras misiones que tenemos para tí consta del asesinato de varios dragones, para obtener así los objetos necesarios para la fundación del clan.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Las misiones disponibles estarán distribuidas en distintas criaturas por toda la ciudad. Deberás buscarlas y ver en que consiste cada misión.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los comandos que se usan son /QUEST e /INFOQUEST", 255, 255, 255, True)
        
        Case 11 ' Invocaciones
            Call AddtoRichTextBox(Console, "Un nuevo mapa ha sido habilitado... Dungeon de Dragones. ¿En que consta? Podrás dirigirte a este mapa con 3 amigos más e invocar los dragones de estas tierras. Cada dragon te ayudará para recibir distintas bonificaciones.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Podrás invocar 8 dragones distintos.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El sistema viene relacionado con las misiones del juego. ¡Atento!", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los dragones tienen que ser asesinados con el Báculo Mata Dragones o una Espada Mata Dragones (Se compran en la Ciudad Principal)", 255, 255, 255, True)
        
        Case 12 ' Ranking
            Call AddtoRichTextBox(Console, "En Dragons podrás ver quienes son los mejores personajes. El ranking se caracteriza por mostrarte los primeros 100 personajes (TOP100)", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los TOP se basan en: Nivel, Frags, Retos (1vs1), Retos (2vs2) y Torneos ganados. ", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los primeros tres personajes quedaran visualizados al comienzo del Ranking.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Con el comando /RANKING podrás ver quienes son los mejores de estas tierras.", 255, 255, 255, True)
            
        Case 13 ' Ranking Online
            Call AddtoRichTextBox(Console, "De igual forma que el Ranking TOP100, en dragons podrás saber que personaje de los que está ONLINE es el mejor tanto en la facción Legión Oscura como en la Armada Real.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "¿Qué estás esperando para ser el mejor?", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Con el comando /RANKINGONLINE podrás ver quien es el mejor usuario ONLINE", 255, 255, 255, True)
            
        Case 14 ' Reliquias del Dragon
            Call AddtoRichTextBox(Console, "Las reliquias del dragon son anillos y objetos especiales que se equipan y que otorgan al personaje beneficios.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los efectos que causan sobre el personaje son variados. Estos anillos no son acumulativos, podrás equipar solo uno a la vez.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los efectos que pueden darte son: Daño mágico, daño físico, evasión, defensa física, defensa mágica, daño sobre npcs, reducción de parálisis, aumento de experiencia, entre otros.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Se obtienen a través de las Cajas Amarillas y en eventos ESPECIALES.", 255, 255, 255, True)
            
        Case 15 ' Castillo de Clanes
            Call AddtoRichTextBox(Console, "Todos los clanes de las tierras de dragones podrán conquistar el Castillo del Rey", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El rey es protegido por sus guardianes, los cuales tienen una inteligencia artificial propia y te atacarán hasta destruirte.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Para atacar el castillo deberás ser miembro de algun Clan", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El clan conquistador recibirá premios cada WorldSave. Los premios son entregados a los personajes ONLINE.", 255, 255, 255, True)
        
        Case 16 ' Fortaleza del Rey Dragon
            Call AddtoRichTextBox(Console, "Todos los usuarios podrán conquistar la Fortaleza de un Rey Dragon. La ventaja de este Castillo es que no necesitas de ayuda para conquistarlo, podrás hacerlo tu solo. Eso si.. Ten cuidado que otros usuarios pensarán lo mismo que tú!", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Cada WorldSave del juego el personaje conquistador ganará diversos premios y beneficios para su personaje ¿Quieres saber los beneficios? ¡Conquistalo!", 255, 255, 255, True)
        
        Case 17 ' Usuarios <ORO>
            Call AddtoRichTextBox(Console, "Los usuarios <ORO> poseen beneficios que otros personajes no. Las criaturas nos darán más experiencia y así nuestro personaje subirá más rápido de nivel (30% extra por golpe).", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Además trae consigo la ventaja de poder concurrir a nuevos mapas, incluyendo nuevos lugares de entrenamiento y de Dropeo de objetos.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El desbloqueo del <ORO> se realiza mediante entrenamiento. Al alcanzar el nivel 7, este logro será desbloqueado.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Al cliclear el personaje aparecerá <ORO> en color Amarillo.", 255, 255, 255, True)
        
        Case 18 ' Usuarios <PLATINO>
            Call AddtoRichTextBox(Console, "Los usuarios <PLATINO> poseen una característica especial. Su daño físico y daño mágico aumentará un 3%.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Al cliclear el personaje aparecerá <PLATINO> en color Gris.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Daño 25% extra a las criaturas", 255, 255, 255, True)
        
        Case 19 ' Usuarios <LEGENDARIO>
            Call AddtoRichTextBox(Console, "Los usuarios <LEGENDARIO> no podrán ser atacados por Npcs. El efecto de su poder los hace inmune.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Al cliclear el personaje aparecerá <LEGENDARIO> en color Verde.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Daño 25% extra a las criaturas", 255, 255, 255, True)
            
        Case 20 ' Angeles y demonios
            Call AddtoRichTextBox(Console, "Al alcanzar el nivel 15 según la facción de tu personaje podrás convertirte en un Angel o en un Demonio.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "A esto se lo denomina TRANSFORMACIÓN y se realiza mediante el comando /TRANSFORMAR", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Al transformarte tu energía comenzará a bajar lentamente. Al llegar a cero, la transformación desaparece y recuperarás tu apariencia normal.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Durante la transformación no es posible mimetizarte con criaturas/usuarios", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Estar transformado te otorga la ventaja de aumentar tu daño contra criaturas y también contra otros personajes.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Daño 30% (Npcs)", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Daño +5 (Usuarios)", 255, 255, 255, True)
        
        Case 21 ' PODER SUPREMO
            Call AddtoRichTextBox(Console, "Alcanzar le poder de los Dioses no es una tarea fácil. Deberás acabar con la vida de varios personajes de forma consecutiva y sin morir. Si logras esto tu personaje pasará a ser Dios, durante un tiempo determinado.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Se te llenará una barra y esta irá bajando lentamente.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "50% de ataque extra contra criaturas.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Las criaturas no te pegan ni paralizan.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Color de NICK en blanco.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Al cliclearte aparece [DIOS] en Blanco.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Reconocimiento en la consola al ganar el poder", 255, 255, 255, True)
            
        Case 22 ' PODER BONUS
            Call AddtoRichTextBox(Console, "¡MUY PRONTO!", 255, 255, 255, True)
        
        Case 23 ' FUNDAR CLAN
            Call AddtoRichTextBox(Console, "En las tierras de dragones fundar clan viene acompañado de unos requisitos únicos.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Dichas reliquias se logran obtener tras realizar invocaciones junto a tus futuros compañeros de clan.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Es por eso que deberán realizar un trabajo en equipo para cumplir el objetivo final.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Una vez obtenido todos los requisitos deberás tipear /FUNDARCLAN y escoger la alineación de tu clan", 255, 255, 255, True)
            
            Call AddtoRichTextBox(Console, "Requisitos para la fundación del clan:", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Nivel 9", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Libro de sabiduria", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Libro del Liderazgo", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Amuleto Anhk", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "100 Diamantes Azules.", 255, 255, 255, True)
            
        Case 24 ' CLAN VS CLAN
            Call AddtoRichTextBox(Console, "Una vez que poseas un clan, el líder podra realizar combates contra otros clanes. El número de miembros mínimo por clan es de '3'. Y debes tener en cuenta que los usuarios a jugar tienen que estar disponibles, es decir no deberán estar en la cárcel, en otro evento y cuestiones demás cuestiones lógicas.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Para enviar un enfrentamiento el LIDER del clan n°1 deberá tipear el comando /CVC LiderN°2", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El Lider n°2 tipea el comando /SICVC LiderN°1 para aceptar el enfrentamiento.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Si ambos equipos están listos, se los enviará a la arena de combate.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El Ganador del Clan vs Clan se llevará GuildPoints que permitén a su clan, obtener mejoras.", 255, 255, 255, True)
        
        Case 25 ' Rey Inmortal
            Call AddtoRichTextBox(Console, "Un temido Rey presenció las tirras de dragones hace mucho tiempo. Ambas facciones tuvieron que unirse para lograr derrotarlo y así obtener su enorme poder.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "El Rey y el demonio decidieron distribuir cofres donde el poder del Rey Inmortal se conserve y así poder ser utilizado por los personajes.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "Los cofres disponibles son [ORO] [PLATINO] [LEGENDARIO] Y [PREMIUM] considerado el mejor de estas tierras.", 255, 255, 255, True)
            Call AddtoRichTextBox(Console, "En la Ciudad Principal estará el Rey Inmortal petrificado con sus cofres a la venta.", 255, 255, 255, True)
            
    End Select

End Sub
