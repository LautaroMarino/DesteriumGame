VERSION 5.00
Begin VB.Form FrmPanelCreator 
   Caption         =   "Panel Creator 3000"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
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
   ScaleHeight     =   3780
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Personaje"
      Height          =   360
      Left            =   3360
      TabIndex        =   24
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtGld 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   22
      Text            =   "5000000"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtFrags 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   21
      Text            =   "50"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtGld 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   20
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtFrags 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   18
      Text            =   "0"
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox cmbGenero 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtDsp 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Publicar en SHOP"
      Height          =   360
      Left            =   840
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtElv 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtUps 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox cmbRaza 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblExito 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   195
      Left            =   3120
      TabIndex        =   23
      Top             =   2400
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frags:"
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
      Left            =   3240
      TabIndex        =   17
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio de Venta DSP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje Generado:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   2265
   End
   Begin VB.Label lblPjs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Descripcion)"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ups:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "FrmPanelCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UserClase As Byte
Private UserRaza As Byte
Private UserLvl As Byte
Private UserUps As Byte
Private UserHead As Integer
Private UserSexo As Byte
Private UserFrags As Integer
Private UserGld As Long

Private Sub btnSave_Click()

    If PersonajeExiste(txtName.Text) Then
        FrmPanelCreator.lblExito.Caption = "El personaje ya existe."
        Exit Sub
    End If
    
    IUser_Editation True
End Sub

Private Sub cmbClass_Click()
    UserClase = cmbClass.ListIndex + 1
    
    UpdateLabel
End Sub
Private Sub cmbGenero_Click()
    UserSexo = cmbGenero.ListIndex + 1
    
    UpdateLabel
End Sub
Private Sub cmbRaza_Click()
    UserRaza = cmbRaza.ListIndex + 1
  
    UpdateLabel
End Sub

Private Sub UpdateLabel()
    
      lblPjs.Caption = vbNullString
      If UserClase > 0 Then lblPjs.Caption = lblPjs.Caption & ListaClases(UserClase)
      If UserRaza > 0 Then lblPjs.Caption = lblPjs.Caption & " " & ListaRazas(UserRaza)
      If UserLvl > 0 Then lblPjs.Caption = lblPjs.Caption & " " & UserLvl
      If UserUps > 0 Then lblPjs.Caption = lblPjs.Caption & " +" & UserUps
End Sub

Private Sub Command1_Click()
       If PersonajeExiste(txtName.Text) Then
        FrmPanelCreator.lblExito.Caption = "El personaje ya existe."
        Exit Sub
    End If
    
    IUser_Editation False
End Sub

Private Sub Form_Load()

    Dim A As Long
    
    For A = 1 To NUMCLASES
        cmbClass.AddItem ListaClases(A)
    
    Next A
    
    For A = 1 To NUMRAZAS
        cmbRaza.AddItem ListaRazas(A)
    Next A
    
    cmbGenero.AddItem "Hombre"
    cmbGenero.AddItem "Mujer"
    
    cmbGenero.ListIndex = 0
    cmbClass.ListIndex = 0
    cmbRaza.ListIndex = 0
    
    UserLvl = 1
    UserUps = 0
    UserGld = 0
    UserFrags = 0
End Sub

Private Sub txtElv_Change()
    If Not IsNumeric(txtElv.Text) Then
        txtElv.Text = "1"
    End If
    
    If val(txtElv) <= 0 Or val(txtElv) > STAT_MAXELV Then
        txtElv.Text = STAT_MAXELV
    End If
    
    UserLvl = val(txtElv.Text)
    
     UpdateLabel
End Sub

Private Sub txtFrags_Change(Index As Integer)
    If Not IsNumeric(txtFrags(Index).Text) Then
        txtFrags(Index).Text = "1"
    End If
    
    If val(txtFrags(Index)) < 0 Or val(txtFrags(Index)) > 200 Then
        txtFrags(Index).Text = 200
    End If
    
    UserFrags = val(txtFrags(Index).Text)

End Sub

Private Sub txtGld_Change(Index As Integer)
 If Not IsNumeric(txtGld(Index).Text) Then
        txtGld(Index).Text = "1"
    End If
    
    If val(txtGld(Index)) <= 0 Or val(txtGld(Index)) > 5000000 Then
        txtGld(Index).Text = 0
    End If
    
    UserGld = val(txtGld(Index).Text)
End Sub

Private Sub txtUps_Change()
    If Not IsNumeric(txtUps.Text) Then
        txtUps.Text = "1"
    End If
    
    If val(txtUps) <= 0 Or val(txtUps) > 40 Then
        txtUps.Text = 40
    End If
    
    UserUps = val(txtUps.Text)
    
     UpdateLabel
End Sub

Private Sub IUser_Editation(ByVal Publicar As Boolean)
        '<EhHeader>
        On Error GoTo IUser_Editation_Err
        '</EhHeader>

        Dim Temp     As User, CharShop As tShopChars
        Dim A As Long
    
        Dim UserName As String
        Dim QuestIndex As Integer
    
100     UserName = txtName.Text
    
102     Call ConnectNewUser(UserName, UserClase, UserRaza, UserSexo, 0, Temp)
104     Temp.Pos = Ullathorpe
106     'QuestIndex = 34
108     'Temp.QuestStats.QuestIndex = QuestIndex
        'HACER
110    ' If QuestList(QuestIndex).RequiredNPCs > 0 Then ReDim Temp.QuestStats.NPCsKilled(1 To QuestList(QuestIndex).RequiredNPCs) As Long
112    ' If QuestList(QuestIndex).RequiredChestOBJs > 0 Then ReDim Temp.QuestStats.ObjsPick(1 To QuestList(QuestIndex).RequiredChestOBJs) As Long
114    ' If QuestList(QuestIndex).RequiredSaleOBJs > 0 Then ReDim Temp.QuestStats.ObjsSale(1 To QuestList(QuestIndex).RequiredSaleOBJs) As Long
    
116     Call InitialUserStats(Temp)
118     Call UserLevelEditation(Temp, UserLvl, UserUps)
120     Call IUser_Editation_Skills(Temp)
122     Call IUser_Editation_Reputacion_Frags(Temp, UserFrags)
          If Temp.Stats.MaxMan > 0 Then Call IUser_Editation_Spells(Temp)
          
124     Temp.Stats.SkillPts = RandomNumber(0, 17)

126     Temp.UpTime = RandomNumber(19124, 354545)
128     Temp.Stats.NPCsMuertos = RandomNumber(777, 2786)
    
130     Temp.flags.Desnudo = 1
132     Temp.Stats.Gld = RandomNumber(val(txtGld(0).Text), val(txtGld(1).Text))
134     Temp.Stats.Exp = RandomNumber(0, (Temp.Stats.Elu / 2.5))

136     Call SaveUser(Temp, CharPath & UCase$(UserName) & ".chr")
    
        
        If Publicar Then
            ' Publicación de SHOP
138         With CharShop
140             .Name = UCase$(UserName)
142             .Dsp = val(txtDsp.Text)
            
144             Call Shop_CharAdd(CharShop)
            End With
        
               FrmPanelCreator.lblExito.Caption = "Personaje publicado: " & UserName
            Else
                FrmPanelCreator.lblExito.Caption = "Personaje creado: " & UserName
            
            End If
            
        '<EhFooter>
        Exit Sub

IUser_Editation_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.FrmPanelCreator.IUser_Editation " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
