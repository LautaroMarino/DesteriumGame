VERSION 5.00
Begin VB.Form FrmConnect 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Conectando Mundo"
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
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MenuExterno 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   4080
      Picture         =   "frmConnect.frx":000C
      ScaleHeight     =   4050
      ScaleWidth      =   3780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   3780
      Begin VB.TextBox txtNameExtern 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   380
         TabIndex        =   0
         Top             =   1000
         Width           =   3000
      End
      Begin VB.Image imgConfirm 
         Height          =   465
         Left            =   580
         Top             =   2380
         Width           =   2655
      End
      Begin VB.Image imgVolver 
         Height          =   465
         Left            =   850
         Top             =   2950
         Width           =   2145
      End
   End
   Begin VB.Timer tSearching 
      Interval        =   5000
      Left            =   1800
      Top             =   2760
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4515
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6300
      Width           =   3000
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   4515
      TabIndex        =   2
      Top             =   5595
      Width           =   3000
   End
   Begin VB.Timer tEnabled 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   525
      Top             =   1995
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   11670
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgRecover 
      Height          =   225
      Left            =   4935
      Top             =   8130
      Width           =   2115
   End
   Begin VB.Image imgConnect 
      Height          =   465
      Left            =   4995
      Top             =   7530
      Width           =   2145
   End
   Begin VB.Image imgNewAccount 
      Height          =   465
      Left            =   4725
      Top             =   6960
      Width           =   2655
   End
End
Attribute VB_Name = "FrmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TimeButton       As Long
Private cBotonCerrar     As clsGraphicalButton
Private cBotonConfirmar  As clsGraphicalButton
Private cBotonRegister   As clsGraphicalButton
Private cBotonConectar   As clsGraphicalButton
Private cBotonVolver     As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    Me.Picture = LoadPicture(DirInterface & "connect\initial.jpg")
    
    Call LoadButtons

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyEscape Then
        imgUnload_Click
        Exit Sub
    End If
    
End Sub
Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = DirInterface

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonRegister = New clsGraphicalButton
    Set cBotonConectar = New clsGraphicalButton
    Set cBotonVolver = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCerrar.Initialize(imgUnload, vbNullString, GrhPath & "generic\BotonCerrarActivo.jpg", vbNullString, Me)
    Call cBotonRegister.Initialize(imgNewAccount, vbNullString, GrhPath & "connect\BotonRegistroActivo.jpg", vbNullString, Me)
    Call cBotonConectar.Initialize(imgConnect, vbNullString, GrhPath & "connect\BotonConectarseActivo.jpg", vbNullString, Me)
    
    Call cBotonVolver.Initialize(imgVolver, vbNullString, GrhPath & "connect\BotonVolverActivo.jpg", vbNullString, Me)
    Call cBotonConfirmar.Initialize(imgConfirm, vbNullString, GrhPath & "connect\BotonConfirmarActivo.jpg", vbNullString, Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgConnect_Click()
    If tEnabled.Enabled = True Then
         Call MsgBox("Espera unos instantes para volver a ingresar. ¡Haz realizado más de un clic al mismo tiempo!")
        Exit Sub
    End If
    
    Call Audio.PlayInterface(SND_CLICK)
    Account.Email = txtName.Text
    Account.Passwd = txtPasswd.Text
    
    
    If Account.Email = vbNullString Then
        Call MsgBox("Debes ingresar una cuenta")
        Exit Sub
    End If
    
    If Not CheckMailString(Account.Email) Then
        Call MsgBox("Corrobora que el Email ingresado sea válido.")
        Exit Sub
    End If
        
    If Len(Account.Passwd) < 5 Then
        Call MsgBox("La contraseña parece ser inválida.")
        Exit Sub
    End If
    
    TimeButton = 0
    Prepare_And_Connect E_MODO.e_LoginAccount
    
    tEnabled.Enabled = True
End Sub

Private Sub imgInfo_Click()

End Sub

Private Sub imgNewAccount_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    MenuExterno.Picture = LoadPicture(DirInterface & "connect\NuevaCuenta.jpg")
    Account_PanelSelected = ePanelAccountRegister
    MenuExterno.visible = True
End Sub

Private Sub imgRecover_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    MenuExterno.Picture = LoadPicture(DirInterface & "connect\Recuperar.jpg")
    
    Account_PanelSelected = ePanelAccountRecover
    MenuExterno.visible = True
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If MsgBox("¿Desea cerrar el juego?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        prgRun = False
    Else

        Exit Sub

    End If
End Sub

Private Sub imgVolver_Click()
    Account_PanelSelected = ePrincipal
    MenuExterno.visible = False
End Sub

Private Sub tEnabled_Timer()
    TimeButton = TimeButton + 1
    
    If TimeButton >= 3 Then
        If Not IsConnected Then
            Call MsgBox("Servidor aparentemente Offline.", vbExclamation)
        End If
        
        imgConnect.Enabled = True
        TimeButton = 0
        tEnabled.Enabled = False
    End If
End Sub

Private Sub txtName_Change()
    ' Chequeos básicos de Seguridad a la hora de escribir textos
    If Right$(txtName, 2) = "  " Then
        txtName.Text = RTrim$(txtName.Text)
    End If
      
    If Left$(txtName, 1) = " " Then
        txtName.Text = LTrim$(txtName)
    End If
    
    If txtName.Text <> vbNullString Then
        Call AutoCompletar(txtName)
        txtPasswd.Text = SearchPasswd(txtName.Text)
    End If

End Sub


Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        imgUnload_Click
        Exit Sub
    End If
End Sub

Private Sub txtPasswd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        imgUnload_Click
        Exit Sub
    End If
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then imgConnect_Click
    
End Sub


Private Sub AutoCompletar(textBox As textBox)

    Dim i         As Integer

    Dim posSelect As Integer
  
    With textBox

        For i = 1 To NUMPASSWD

            ' Buscamos coincidencias
            If InStr(1, ListPasswd(i).Account, .Text, vbTextCompare) = 1 Then
                posSelect = .SelStart
                .Text = ListPasswd(i).Account
              
                ' seleccionar el texto
                .SelStart = posSelect
                .SelLength = Len(.Text) - posSelect

                Exit For ' salimos del bucle

            End If

        Next i
  
    End With

End Sub

Private Sub AutoCompletar_Passwd(textBox As textBox)

    Dim i         As Integer

    Dim posSelect As Integer
  
    With textBox

        For i = 1 To NUMPASSWD

            ' Buscamos coincidencias
            If InStr(1, ListPasswd(i).Passwd, .Text, vbTextCompare) = 1 Then
                posSelect = .SelStart
                .Text = ListPasswd(i).Passwd
              
                ' seleccionar el texto
                .SelStart = posSelect
                .SelLength = Len(.Text) - posSelect

                Exit For ' salimos del bucle

            End If

        Next i
  
    End With

End Sub
