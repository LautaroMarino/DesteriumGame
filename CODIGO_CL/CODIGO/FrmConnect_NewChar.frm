VERSION 5.00
Begin VB.Form FrmConnect_NewChar 
   BorderStyle     =   0  'None
   Caption         =   "Crear nuevo personaje"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "FrmConnect_NewChar.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   3750
      Width           =   2715
   End
   Begin VB.ComboBox cmbGenero 
      BackColor       =   &H80000001&
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
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6060
      Width           =   2415
   End
   Begin VB.ComboBox cmbRaze 
      BackColor       =   &H80000001&
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
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5280
      Width           =   2415
   End
   Begin VB.ComboBox cmbClass 
      BackColor       =   &H80000001&
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
      Height          =   360
      Left            =   4650
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4530
      Width           =   2415
   End
   Begin VB.Image imgUnload 
      Height          =   465
      Left            =   10680
      Top             =   750
      Width           =   885
   End
   Begin VB.Image imgManual 
      Height          =   495
      Left            =   4590
      Top             =   7020
      Width           =   2535
   End
   Begin VB.Image imgNew 
      Height          =   435
      Left            =   3930
      Top             =   7620
      Width           =   3945
   End
End
Attribute VB_Name = "FrmConnect_NewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit


Private Sub Form_Activate()
    'UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserEmail = vbNullString
    UserKey = vbNullString
    UserPin = vbNullString
    UserHead = 0
    WorkGenero = 0
    WorkRaza = 0
    WorkName = vbNullString

    Call CargarCombos
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        frmConnect.Show
        Unload Me
    End If

End Sub

Private Sub CargarCombos()

    Dim A As Integer
          
    cmbClass.Clear
    cmbGenero.Clear
    cmbRaze.Clear
    
    For A = 1 To NUMCLASES
        cmbClass.AddItem ListaClases(A)
    Next A

          
    For A = 1 To NUMRAZAS
        cmbRaze.AddItem ListaRazas(A)
    Next A

    
    cmbGenero.AddItem "Hombre"
    cmbGenero.AddItem "Mujer"
    

    cmbClass.ListIndex = 0
    cmbRaze.ListIndex = 0
    cmbGenero.ListIndex = 0
    
    
    txtName.Text = vbNullString
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub
