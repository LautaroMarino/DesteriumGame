VERSION 5.00
Begin VB.Form FrmMercader_Other 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Más información"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMercader_Other.frx":0000
   ScaleHeight     =   6810
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbChars 
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
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   2940
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   90
      Width           =   1965
   End
   Begin VB.Label lblQuest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Niguna"
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
      Height          =   2175
      Left            =   6960
      TabIndex        =   21
      Top             =   1440
      Width           =   2505
   End
   Begin VB.Label lblPenas 
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
      ForeColor       =   &H8000000B&
      Height          =   1695
      Left            =   150
      TabIndex        =   19
      Top             =   4920
      Width           =   4395
   End
   Begin VB.Label lblSpells 
      Alignment       =   2  'Center
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
      ForeColor       =   &H8000000B&
      Height          =   3105
      Left            =   4920
      TabIndex        =   18
      Top             =   3660
      Width           =   2145
   End
   Begin VB.Label lblNobleza 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   17
      Top             =   2580
      Width           =   2025
   End
   Begin VB.Label lblAsesino 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   16
      Top             =   2370
      Width           =   2025
   End
   Begin VB.Label lblBandido 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   15
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Label lblEvents 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   14
      Top             =   1370
      Width           =   2025
   End
   Begin VB.Label lblRetos3 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   13
      Top             =   1160
      Width           =   2025
   End
   Begin VB.Label lblRetos2 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   12
      Top             =   930
      Width           =   2025
   End
   Begin VB.Label lblRetos1 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   11
      Top             =   720
      Width           =   2025
   End
   Begin VB.Label lblHonour 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   5700
      TabIndex        =   10
      Top             =   510
      Width           =   2025
   End
   Begin VB.Label lblNpcs 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   9
      Top             =   4280
      Width           =   2025
   End
   Begin VB.Label lblFrags 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   8
      Top             =   4070
      Width           =   2025
   End
   Begin VB.Label lblCiu 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   7
      Top             =   3860
      Width           =   2025
   End
   Begin VB.Label lblCri 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   6
      Top             =   3630
      Width           =   2025
   End
   Begin VB.Label lblExFaction 
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
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1500
      TabIndex        =   5
      Top             =   2340
      Width           =   2025
   End
   Begin VB.Label lblFaction 
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
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1500
      TabIndex        =   4
      Top             =   2130
      Width           =   2025
   End
   Begin VB.Label lblEldhir 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   3
      Top             =   1120
      Width           =   2025
   End
   Begin VB.Label lblGld 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   2
      Top             =   910
      Width           =   2025
   End
   Begin VB.Label lblMan 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   1
      Top             =   710
      Width           =   2025
   End
   Begin VB.Label lblHp 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   1500
      TabIndex        =   0
      Top             =   480
      Width           =   2025
   End
End
Attribute VB_Name = "FrmMercader_Other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbChars_Click()
    ChangeInfo (cmbChars.ListIndex)
End Sub

Private Sub Form_Load()
    Dim A As Long
    
    
    For A = 0 To 4
        If MercaderChars(A).Name <> vbNullString Then
            cmbChars.AddItem (MercaderChars(A).Name)
        End If
    Next A
End Sub


Private Sub ChangeInfo(ByVal SelectedChar As Byte)
    
    Dim A As Long
    Dim Temp As String
    
    With MercaderChars(SelectedChar)
        lblHp.Caption = .Hp
        lblMan.Caption = IIf(.Man > 0, Format$(.Man, "#,##"), "0")
        lblGld.Caption = IIf(.Gld > 0, Format$(.Gld, "#,##"), "0")
        lblEldhir.Caption = IIf(.Eldhir > 0, Format$(.Eldhir, "#,##"), "0")
        
        
        If .Faction = 1 Then
            lblFaction.Caption = "Armada Real"
            lblFaction.ForeColor = vbCyan
        ElseIf .Faction = 2 Then
            lblFaction.Caption = "Legión Oscura"
            lblFaction.ForeColor = vbRed
        Else
            lblFaction.Caption = "Ninguna"
            lblFaction.ForeColor = vbWhite
        End If
        
        If .FactionEx = 1 Then
            lblExFaction.Caption = "ex Armada Real"
            lblExFaction.ForeColor = vbCyan
        ElseIf .FactionEx = 2 Then
            lblExFaction.Caption = "ex Legión Oscura"
            lblExFaction.ForeColor = vbRed
        Else
            lblExFaction.Caption = "NO"
            lblExFaction.ForeColor = vbGreen
        End If
        
        lblCri.Caption = IIf(.FragsCri > 0, Format$(.FragsCri, "#,##"), "0")
        lblCiu.Caption = IIf(.FragsCiu > 0, Format$(.FragsCiu, "#,##"), "0")
        lblFrags.Caption = IIf(.FragsOther > 0, Format$(.FragsOther, "#,##"), "0")
        lblNpcs.Caption = IIf(.FragsNpc > 0, Format$(.FragsNpc, "#,##"), "0")
        
        
        If .Penas > 0 Then
            
            For A = 1 To .Penas
                Temp = Temp & A & "° " & .PenasText(A) & vbCrLf
            Next A
            
            lblPenas.Caption = Temp
        Else
            lblPenas.Caption = "No tiene penas."
        End If
        
        lblHonour.Caption = IIf(.Points > 0, Format$(.Points, "#,##"), "0")
        lblRetos1.Caption = .Retos1Ganados & "/" & .Retos1Jugados
        lblRetos2.Caption = .Retos2Ganados & "/" & .Retos2Jugados
        lblRetos3.Caption = .Retos3Ganados & "/" & .Retos3Jugados
        lblEvents.Caption = .EventosGanados & "/" & .EventosJugados
        
        
        Temp = vbNullString
        For A = 1 To 35
            If .Spells(A) <> vbNullString Then
                Temp = Temp & .Spells(A) & vbCrLf
            
            End If
        Next A
        
        lblSpells.Caption = Temp
        
        
        lblBandido.Caption = IIf(.Bandido > 0, Format$(.Bandido, "#,##"), "0")
        lblAsesino.Caption = IIf(.Asesino > 0, Format$(.Asesino, "#,##"), "0")
        lblNobleza.Caption = IIf(.Nobleza > 0, Format$(.Nobleza, "#,##"), "0")
        
        Temp = vbNullString
        
        If (.NumQuests > 0) Then
            For A = LBound(.Quests) To UBound(.Quests)
                Temp = Temp & QuestList(.Quests(A)).Name & vbCrLf
            Next A
        Else
            lblQuest.Caption = "Ninguna"
        End If
        
       
        lblQuest.Caption = Temp
        
    End With
End Sub

