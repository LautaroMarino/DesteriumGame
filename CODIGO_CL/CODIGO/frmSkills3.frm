VERSION 5.00
Begin VB.Form frmSkills3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSkills3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "frmSkills3.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   9
      Left            =   4500
      Top             =   4470
      Width           =   270
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999"
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
      Height          =   270
      Left            =   4350
      TabIndex        =   15
      Top             =   1080
      Width           =   465
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   15
      Left            =   3600
      Top             =   2250
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   14
      Left            =   3600
      Top             =   6270
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   13
      Left            =   3600
      Top             =   3750
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   12
      Left            =   3600
      Top             =   6555
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   11
      Left            =   3600
      Top             =   5970
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   10
      Left            =   3600
      Top             =   5070
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   9
      Left            =   3600
      Top             =   4470
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   8
      Left            =   3600
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   7
      Left            =   3600
      Top             =   4770
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   6
      Left            =   3600
      Top             =   5370
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   5
      Left            =   3600
      Top             =   3450
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   4
      Left            =   3600
      Top             =   2550
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   3
      Left            =   3600
      Top             =   3150
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   2
      Left            =   3600
      Top             =   5670
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   1
      Left            =   3600
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   15
      Left            =   4500
      Top             =   2250
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   14
      Left            =   4500
      Top             =   6270
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   13
      Left            =   4500
      Top             =   3750
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   12
      Left            =   4500
      Top             =   6555
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   11
      Left            =   4500
      Top             =   5970
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   10
      Left            =   4500
      Top             =   5070
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   8
      Left            =   4500
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   7
      Left            =   4500
      Top             =   4770
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   6
      Left            =   4500
      Top             =   5370
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   5
      Left            =   4500
      Top             =   3450
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   4
      Left            =   4500
      Top             =   2550
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   3
      Left            =   4500
      Top             =   3150
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   2
      Left            =   4500
      Top             =   5670
      Width           =   270
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   15
      Left            =   4020
      TabIndex        =   14
      Top             =   2265
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   14
      Left            =   4020
      TabIndex        =   13
      Top             =   6285
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   13
      Left            =   4020
      TabIndex        =   12
      Top             =   3765
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   12
      Left            =   4020
      TabIndex        =   11
      Top             =   6570
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   11
      Left            =   4020
      TabIndex        =   10
      Top             =   6000
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   10
      Left            =   4020
      TabIndex        =   9
      Top             =   5100
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   9
      Left            =   4020
      TabIndex        =   8
      Top             =   4500
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   8
      Left            =   4020
      TabIndex        =   7
      Top             =   2865
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   7
      Left            =   4020
      TabIndex        =   6
      Top             =   4800
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   6
      Left            =   4020
      TabIndex        =   5
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   5
      Left            =   4020
      TabIndex        =   4
      Top             =   3465
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   4
      Left            =   4020
      TabIndex        =   3
      Top             =   2565
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   3
      Left            =   4020
      TabIndex        =   2
      Top             =   3165
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   2
      Left            =   4020
      TabIndex        =   1
      Top             =   5685
      Width           =   315
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   4020
      TabIndex        =   0
      Top             =   1965
      Width           =   315
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   1
      Left            =   4500
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgAceptar 
      Height          =   495
      Left            =   1680
      Picture         =   "frmSkills3.frx":10BE
      Top             =   6900
      Width           =   1680
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario          As clsFormMovementManager

Private bPuedeMagia            As Boolean

Private bPuedeMeditar          As Boolean

Private bPuedeEscudo           As Boolean

Private bPuedeCombateDistancia As Boolean

Private vsHelp(1 To NUMSKILLS) As String

Private Const ANCHO_BARRA      As Byte = 73 'pixeles

Private Const BAR_LEFT_POS     As Integer = 361 'pixeles


' Botones Graficos
Private cBotonCerrar       As clsGraphicalButton
Private cBotonAceptar    As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    MirandoEstadisticas = True
    
    Call LoadButtons
    
    Me.Picture = LoadPicture(DirInterface & "menucompacto\skills.jpg")
    
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
          
    'Flags para saber que skills se modificaron
    ReDim Flags(1 To NUMSKILLS)
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserEstadisticas.Skills(i) = UserSkills(i)
        lblSkill(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
    lblPoints.Caption = Alocados
    
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = DirInterface

    Set cBotonAceptar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "menucompacto\buttons\save.jpg", GrhPath & "menucompacto\buttons\save_hover.jpg", GrhPath & "menucompacto\buttons\save.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    LastButtonPressed.ToggleToNormal
    
    Dim A As Long
    
    For A = 1 To NUMSKILLS
        imgMenos(A).Picture = Nothing
        imgMas(A).Picture = Nothing
    Next A
End Sub


Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)

    If Alocados > 0 Then

        If Val(UserEstadisticas.Skills(SkillIndex)) < MAXSKILLPOINTS Then
            UserEstadisticas.Skills(SkillIndex) = Val(UserEstadisticas.Skills(SkillIndex)) + 1
            Flags(SkillIndex) = Flags(SkillIndex) + 1
            Alocados = Alocados - 1
            
        End If
                  
    End If
          
   ' SkillPoints = Alocados
   lblSkill(SkillIndex).Caption = UserEstadisticas.Skills(SkillIndex)
   lblPoints.Caption = Alocados
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)

    If Alocados < SkillPoints Then
              
        If Val(UserEstadisticas.Skills(SkillIndex)) > 0 And Flags(SkillIndex) > 0 Then
            UserEstadisticas.Skills(SkillIndex) = Val(UserEstadisticas.Skills(SkillIndex)) - 1
            Flags(SkillIndex) = Flags(SkillIndex) - 1
            Alocados = Alocados + 1
        End If
    End If
          
          
    'SkillPoints = Alocados
    lblSkill(SkillIndex).Caption = UserEstadisticas.Skills(SkillIndex)
    lblPoints.Caption = Alocados
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoEstadisticas = False
End Sub

Private Sub imgAceptar_Click()
    
    On Error GoTo ErrHandler
    
    

        Dim skillChanges(NUMSKILLS) As Byte
    
        Dim i                       As Long
    
    
        If SkillPoints > 0 Then
            For i = 1 To NUMSKILLS
                skillChanges(i) = CByte(UserEstadisticas.Skills(i)) - UserSkills(i)
                'Actualizamos nuestros datos locales
                UserSkills(i) = Val(UserEstadisticas.Skills(i))
            Next i
            
            Call WriteModifySkills(skillChanges())
              
            SkillPoints = Alocados
        End If
        
    
    Unload Me

    Exit Sub

ErrHandler:
End Sub

Private Sub imgMas_Click(Index As Integer)
     Call Audio.PlayInterface(SND_CLICK)
    Call SumarSkillPoint(Index)
End Sub

Private Sub imgMas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMas(Index).Picture = LoadPicture(DirInterface & "menucompacto\buttons\mas.jpg")
End Sub
Private Sub imgMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMenos(Index).Picture = LoadPicture(DirInterface & "menucompacto\buttons\menos.jpg")
End Sub
Private Sub imgMenos_Click(Index As Integer)
     Call Audio.PlayInterface(SND_CLICK)
    Call RestarSkillPoint(Index)
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub
