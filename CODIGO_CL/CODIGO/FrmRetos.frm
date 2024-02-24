VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmRetos 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   Icon            =   "FrmRetos.frx":0000
   LinkTopic       =   "Retos Privados"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmRetos.frx":000C
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   2835
      MaxLength       =   160
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2715
      Width           =   1905
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   2835
      MaxLength       =   160
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2340
      Width           =   1905
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2835
      MaxLength       =   160
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1995
      Width           =   1905
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   555
      MaxLength       =   160
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2715
      Width           =   1905
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   555
      MaxLength       =   160
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2340
      Width           =   1905
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   555
      MaxLength       =   160
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Width           =   1905
   End
   Begin VB.TextBox txtDsp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   3570
      MaxLength       =   160
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "Chat"
      Top             =   3390
      Width           =   900
   End
   Begin VB.TextBox txtGld 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   1020
      MaxLength       =   160
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "Chat"
      Top             =   3390
      Width           =   1905
   End
   Begin VB.ComboBox cmbTime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5880
      Width           =   855
   End
   Begin VB.ComboBox cmbRounds 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4320
      Width           =   735
   End
   Begin VB.CheckBox chkConfig 
      BackColor       =   &H80000007&
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   6210
      Width           =   195
   End
   Begin VB.CheckBox chkConfig 
      BackColor       =   &H80000007&
      Caption         =   "chkPlantes"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   5625
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chkConfig 
      BackColor       =   &H80000007&
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   3
      Top             =   6240
      Width           =   195
   End
   Begin VB.CheckBox chkConfig 
      BackColor       =   &H80000007&
      Caption         =   "chkPlantes"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   2
      Top             =   5910
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chkConfig 
      BackColor       =   &H80000007&
      Caption         =   "chkPlantes"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   5340
      Width           =   195
   End
   Begin VB.CheckBox chkConfig 
      BackColor       =   &H80000007&
      Caption         =   "chkPlantes"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin RichTextLib.RichTextBox txtRojas 
      Height          =   300
      Left            =   6240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4125
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   529
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmRetos.frx":1DEEE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   7
      Left            =   4245
      Top             =   4860
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   6
      Left            =   3750
      Top             =   4860
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   5
      Left            =   3255
      Top             =   4860
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   4
      Left            =   2760
      Top             =   4860
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   3
      Left            =   4245
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   2
      Left            =   3750
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   1
      Left            =   3255
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image ButtonTex 
      Height          =   480
      Index           =   0
      Left            =   2760
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image ButtonTipo 
      Height          =   330
      Index           =   3
      Left            =   3855
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Image ButtonTipo 
      Height          =   330
      Index           =   2
      Left            =   2685
      Stretch         =   -1  'True
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Image ButtonTipo 
      Height          =   330
      Index           =   1
      Left            =   1530
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Image ButtonTipo 
      Height          =   330
      Index           =   0
      Left            =   375
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmRetos.frx":1DF6B
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgTop 
      Height          =   375
      Left            =   1920
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image lblActionCheck 
      Height          =   195
      Index           =   5
      Left            =   600
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Image lblActionCheck 
      Height          =   195
      Index           =   4
      Left            =   3240
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Image lblActionCheck 
      Height          =   195
      Index           =   3
      Left            =   600
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image lblActionCheck 
      Height          =   195
      Index           =   2
      Left            =   600
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Image lblActionCheck 
      Height          =   195
      Index           =   1
      Left            =   600
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image lblActionCheck 
      Height          =   195
      Index           =   0
      Left            =   600
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image imgSend 
      Height          =   375
      Left            =   1920
      MousePointer    =   2  'Cross
      Top             =   6840
      Width           =   1245
   End
End
Attribute VB_Name = "FrmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private MouseBoton As Integer
Private MouseShift As Integer

Private Const MAX_ZONA As Byte = 8
Private Const MAX_RETOS_PERSONAJES As Byte = 6

Dim Fight As tFight

Public Enum eTypeFight
    eAccept = 1
    eSend = 2
End Enum

Public SelectedFight As Byte ' 1vs1,2vs2,3vs3,4vs4

Public TypeFight As eTypeFight

Public TexFight As Integer
Public TextUsers_Temp As Integer


Public Sub ClicUser(ByVal Name As String)
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim txtCopy() As String
    Dim A As Long
    
    ReDim txtCopy(txtUser.LBound To txtUser.UBound) As String
    
    For A = txtUser.LBound To txtUser.UBound
        txtCopy(A) = txtUser(A).Text
    Next A
    
   ' If HayRepetidos(txtCopy) Then Exit Sub
    If ExistUser(Name) Then Exit Sub
    
    txtUser(TextUsers_Temp).Text = Name
End Sub


Private Sub EffectTexture(Index As Integer)

    TexFight = Index + 1
    
    Dim A As Long
    
    For A = ButtonTex.LBound To ButtonTex.UBound
        Set ButtonTex(A).Picture = Nothing
    Next A
    
    Set ButtonTex(Index).Picture = LoadPicture(DirInterface & "fight\Tex" & TexFight & "_Clic.jpg")
End Sub
Private Sub ButtonTex_Click(Index As Integer)
    If SelectedFight = 4 Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)

    EffectTexture (Index)
    
    
End Sub

Private Sub ButtonTipo_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    SelectedFight = Index + 1
    
    Dim A As Long
    
    For A = ButtonTipo.LBound To ButtonTipo.UBound
        Set ButtonTipo(A).Picture = Nothing
    Next A
    
    Set ButtonTipo(Index).Picture = LoadPicture(DirInterface & "fight\Button" & SelectedFight & ".jpg")
    
    
    ' · Recorrer los txt y poner Enabled=True-False
    For A = 1 To txtUser.UBound
        txtUser(A).Enabled = False
        
        txtUser(A).Text = vbNullString
    Next A
    
    Select Case SelectedFight
        Case 1
            txtUser(3).Enabled = True
            

        Case 2
            txtUser(1).Enabled = True
            
            txtUser(3).Enabled = True
            txtUser(4).Enabled = True
        Case 3
            txtUser(1).Enabled = True
            txtUser(2).Enabled = True
            
            txtUser(3).Enabled = True
            txtUser(4).Enabled = True
            txtUser(5).Enabled = True
            
        Case 4 ' Plantes
            txtUser(3).Enabled = True
            EffectTexture (0)  ' Solo puede elegir esta textura de plante.
    End Select
    

    
End Sub

Private Sub chkConfig_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
End Sub

Private Sub Form_Load()
    
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    Me.Picture = LoadPicture(DirInterface & "fight\VentanaRetos.jpg")
    
    cmbRounds.AddItem "1"
    cmbRounds.AddItem "3"
    cmbRounds.AddItem "5"
    cmbRounds.AddItem "10"
    cmbRounds.AddItem "20"
    
    cmbTime.AddItem "5"
    cmbTime.AddItem "10"
    cmbTime.AddItem "20"
    cmbTime.AddItem "30"
    cmbTime.AddItem "60"
    
    
    cmbRounds.ListIndex = 1
    cmbTime.ListIndex = 1
    
    txtGld.Text = "30000"

    lblActionCheck(0).ToolTipText = "Los personajes no podrán utilizar ningún hechizo que deje inmobil o paralizado a la víctima."
    lblActionCheck(1).ToolTipText = "Tu pareja podrá devolverte a la vida. Seguirás luchando y dando todo por el Honor."
    lblActionCheck(2).ToolTipText = "Los personajes no pueden defenderse con Escudos."
    lblActionCheck(3).ToolTipText = "Los personajes no pueden cubrir su cabeza con ningun objeto."

    TexFight = 1
    
    ButtonTipo_Click (0)
    EffectTexture (0)
    
    txtUser(0).Text = UserName
    txtUser(0).Enabled = False
    
    MirandoRetos = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoRetos = False
End Sub

Private Sub Image1_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If TypeFight = eTypeFight.eAccept Then
        Call WriteFight_CancelInvitation
    End If

    Unload Me
End Sub

' Ir al TOP WEB
Private Sub imgTop_Click()
    Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/retos/", vbNullString, vbNullString, 1)
End Sub

' Simulación del checkbox a través del label
Private Sub lblActionCheck_Click(Index As Integer)

    If chkConfig(Index).Enabled Then
        Call Audio.PlayInterface(SND_CLICK)
        chkConfig(Index) = IIf(chkConfig(Index).Value = 0, 1, 0)
    End If
End Sub




' Efecto de Formulario (Cerrar el Formulario)
Private Sub Form_MouseDown(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
End Sub

Public Function ExistUser(ByVal Name As String)
    Dim A As Long
    
    For A = txtUser.LBound To txtUser.UBound
        If txtUser(A).Text = Name Then
            ExistUser = True
            Exit Function
        End If
    Next A
End Function
Public Function HayRepetidos(Lista() As String) As Boolean

    On Error GoTo ErrHandler
    
    Dim I As Integer
    Dim j As Integer
    
    ' Recorrer la lista
    For I = 0 To UBound(Lista) - 1
        ' Verificar si el elemento actual está repetido en la lista
        For j = I + 1 To UBound(Lista)
            If StrComp(Lista(I), Lista(j), vbTextCompare) = 0 Then
                ' Elemento repetido encontrado, retornar False
                HayRepetidos = True
                Exit Function
            End If
        Next j
    Next I
    
    ' No se encontraron elementos repetidos, retornar True
    HayRepetidos = False
    
    Exit Function
ErrHandler:
    
End Function

Private Function Retos_CheckData(ByRef Users() As String) As Boolean
    
    Dim A As Long
    Dim FoundChar As Boolean

    
    If UBound(Users) = -1 Then
        Call MsgBox("Utiliza el Formato Player1-Player2 vs Enemigo1-Enemigo2.")
        Exit Function
    End If
    
    If HayRepetidos(Users) Then
        Call MsgBox("¡Has puesto un nombre repetido!")
        Exit Function
    End If
    
    For A = LBound(Users) To UBound(Users)
        If Users(A) = vbNullString Then
            Call MsgBox("Has introducido un nombre incorrecto. Fijate de elegir bien la cantidad de usuarios.")
            Exit Function
        End If
        
        If UCase$(Users(A)) = UCase$(FrmMain.Label8(0).Caption) Then
            FoundChar = True
        End If
    Next A
    
    If Not FoundChar Then
        Call MsgBox("¡Tú tambien debes participar del evento! Y ahora que ha hecho eso, pues mas le vale ganar.")
        Exit Function
    End If
    
    If UBound(Users) > (MAX_RETOS_PERSONAJES - 1) Then
        Call MsgBox("Se han encontrado más usuarios de los que tiene permitido el servidor.")
        Exit Function
    End If
    
    If Not (UBound(Users) + 1) Mod 2 = 0 Then
        Call MsgBox("La cantidad de miembros para enfrentarse es inválida.")
        Exit Function
    End If
    
    Retos_CheckData = True
End Function

' # Chequea que no este nada vacio.
Public Function CheckValidUsers_Txt() As Boolean
    
    
    Select Case SelectedFight
        Case 1, 4
                If Len(txtUser(0).Text) = 0 Or Len(txtUser(3).Text) = 0 Then Exit Function
        Case 2
                If Len(txtUser(0).Text) = 0 Or Len(txtUser(3).Text) = 0 Or _
                    Len(txtUser(1).Text) = 0 Or Len(txtUser(4).Text) = 0 Then Exit Function
        Case 3
                If Len(txtUser(0).Text) = 0 Or Len(txtUser(3).Text) = 0 Or _
                    Len(txtUser(1).Text) = 0 Or Len(txtUser(4).Text) = 0 Or _
                    Len(txtUser(2).Text) = 0 Or Len(txtUser(5).Text) = 0 Then Exit Function
        Case 4
        
    End Select
    
    CheckValidUsers_Txt = True
    
End Function
Private Sub imgSend_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If TypeFight = eTypeFight.eSend Then
        Dim Users() As String
        Dim Temp As String
        Dim A As Long

        If Not CheckValidUsers_Txt Then Exit Sub
        
        ' Team n°1
        For A = txtUser.LBound To txtUser.UBound
            If Len(txtUser(A).Text) > 0 Then
                Temp = Temp & txtUser(A).Text & "-"
            End If
        Next A
        
        Temp = Left$(Temp, Len(Temp) - 1)
        Fight.Users = Temp
        Temp = vbNullString
        
        Fight.Gld = Val(txtGld.Text)
        Fight.Tipo = SelectedFight
        
        Users = Split(UCase$(Fight.Users), "-")
        
        If Not Retos_CheckData(Users) Then Exit Sub
    
        For A = LBound(Fight.Config) To UBound(Fight.Config)
            Fight.Config(A) = chkConfig(A).Value
        Next A
        
        Call WriteSendFight(Fight)
                            
        Unload Me
    ElseIf TypeFight = eTypeFight.eAccept Then
        Call WriteAcceptFight(Fight_UserName)
        Unload Me
    
    End If
    
End Sub

Private Sub imgUnload_Click()

End Sub


Private Sub txtGld_Change()
    If Not IsNumeric(txtGld.Text) Then
        txtGld.Text = "30000"
        txtGld.SelStart = Len(txtGld.Text)
    End If
        
    If Val(txtGld.Text) < 30000 Or Val(txtGld.Text) > 200000000 Then
        txtGld.Text = "30000"
        txtGld.SelStart = Len(txtGld.Text)
    End If

    Fight.Gld = Val(txtGld.Text)
End Sub


Private Sub txtUser_Click(Index As Integer)
    TextUsers_Temp = Index
    
End Sub
