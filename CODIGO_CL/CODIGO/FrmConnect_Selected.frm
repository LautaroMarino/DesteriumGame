VERSION 5.00
Begin VB.Form FrmConnect_Selected 
   BorderStyle     =   0  'None
   Caption         =   "BattleServer"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConnect_Selected.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   11640
      Picture         =   "FrmConnect_Selected.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblOns 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   5
      Left            =   10250
      TabIndex        =   7
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label lblOns 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   4
      Left            =   10250
      TabIndex        =   6
      Top             =   4320
      Width           =   525
   End
   Begin VB.Label lblOns 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   3
      Left            =   6780
      TabIndex        =   5
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label lblOns 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   2
      Left            =   6780
      TabIndex        =   4
      Top             =   4320
      Width           =   525
   End
   Begin VB.Label lblOns 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label lblOns 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   4320
      Width           =   525
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000.000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   10200
      TabIndex        =   1
      Top             =   600
      Width           =   1230
   End
   Begin VB.Image imgExtraccion 
      Height          =   495
      Left            =   6360
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Image imgRoyal 
      Height          =   495
      Left            =   1320
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label lblAlias 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNNAMED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   270
      Left            =   8160
      TabIndex        =   0
      Top             =   720
      Width           =   1230
   End
   Begin VB.Image imgArena 
      Height          =   615
      Index           =   5
      Left            =   8160
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Image imgArena 
      Height          =   615
      Index           =   4
      Left            =   8160
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Image imgArena 
      Height          =   735
      Index           =   3
      Left            =   4680
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Image imgArena 
      Height          =   615
      Index           =   2
      Left            =   4680
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Image imgArena 
      Height          =   735
      Index           =   1
      Left            =   1200
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Image imgArena 
      Height          =   615
      Index           =   0
      Left            =   1320
      Top             =   3960
      Width           =   2655
   End
End
Attribute VB_Name = "FrmConnect_Selected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim A As Long
    
    'Me.Picture = LoadPicture(DirInterface & "connect\selected.jpg")
    

End Sub

Private Sub imgArena_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    If Account.Alias = vbNullString Then
        Call MsgBox("Tu nombre/alias aún no fue seteado. Haz clic sobre 'UNNAMED' para configurarlo y comenzar a jugar los servidores")
        Exit Sub
    End If
    
    ServerSelected = Index + 1
  '  Prepare_And_Connect E_MODO.e_LoginBattle, Me
End Sub

Private Sub imgExtraccion_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call MsgBox("Pronto conocerás los modos llenos de diversión y estrategia. ¿Estás listo? ¡Preparate!")
End Sub

Private Sub imgRoyal_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call MsgBox("Pronto conocerás los modos llenos de diversión y estrategia. ¿Estás listo? ¡Preparate!")
End Sub

Private Sub imgUnload_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
   If MsgBox("¿Estás seguro que deseas cerrar tu cuenta?", vbYesNo) = vbYes Then
            Call Disconnect
            FrmConnect_Selected.visible = False
            'FrmConnect.visible = True
            prgRun = False
        End If
End Sub

Private Sub lblAlias_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Name As String
    
    Name = InputBox("Escribe tu nuevo alias para poder jugar. El mínimo de caracteres es " & ACCOUNT_MIN_CHARACTER_CHAR & " y máximo" & ACCOUNT_MAX_CHARACTER_CHAR)

    If Right$(Name, 1) = " " Then
        Name = RTrim$(Name)
    End If
    
    If Len(Name) < ACCOUNT_MIN_CHARACTER_CHAR Then
        Call MsgBox("El nombre debe contener más de " & ACCOUNT_MIN_CHARACTER_CHAR & " caracteres.")
        Exit Sub
    End If
    
    If Len(Name) > ACCOUNT_MAX_CHARACTER_CHAR Then
        Call MsgBox("El nombre debe contener menos de " & ACCOUNT_MAX_CHARACTER_CHAR & " caracteres.")
        Exit Sub
    End If
    
    
    
    If Not ValidarNombre(Name) Then
        Call MsgBox("El Alias contiene caracteres inválidos")
        Exit Sub
    End If
    
    Account.Alias = Name
    UserName = Name
    lblAlias.Caption = UCase$(Name)
    
    Prepare_And_Connect E_MODO.e_LoginName, Me
End Sub
