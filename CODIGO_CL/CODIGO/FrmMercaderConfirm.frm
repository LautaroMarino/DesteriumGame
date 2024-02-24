VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMercaderConfirm 
   BorderStyle     =   0  'None
   Caption         =   "Mercado"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercaderConfirm.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPin 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      IMEMode         =   3  'DISABLE
      Left            =   1530
      MaxLength       =   160
      PasswordChar    =   "*"
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   4995
      Width           =   2175
   End
   Begin VB.TextBox txtPasswd 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      IMEMode         =   3  'DISABLE
      Left            =   1530
      MaxLength       =   160
      PasswordChar    =   "*"
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   4410
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox Console 
      Height          =   1830
      Left            =   600
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes de eventos"
      Top             =   1800
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   3228
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmMercaderConfirm.frx":1797C
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo de la publicacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   825
      TabIndex        =   3
      Top             =   1140
      Width           =   3570
   End
   Begin VB.Image chkView2 
      Height          =   225
      Left            =   3810
      Picture         =   "FrmMercaderConfirm.frx":179FA
      Top             =   4995
      Width           =   210
   End
   Begin VB.Image chkView1 
      Height          =   225
      Left            =   3810
      Picture         =   "FrmMercaderConfirm.frx":18853
      Top             =   4410
      Width           =   210
   End
   Begin VB.Image chkConfirm 
      Height          =   225
      Left            =   555
      Picture         =   "FrmMercaderConfirm.frx":196AC
      Top             =   3750
      Width           =   210
   End
   Begin VB.Image imgView 
      Height          =   255
      Left            =   840
      MousePointer    =   14  'Arrow and Question
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Image ImgSecure 
      Height          =   735
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Image ButtonNext 
      Height          =   375
      Left            =   1680
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMercaderConfirm.frx":1A505
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMercaderConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario          As clsFormMovementManager
Private picCheckBox          As Picture
Private picCheckBoxNulo      As Picture

Private Confirm As Boolean

Public Enum eTypeConfirm
    Venta = 1
    Oferta = 2
End Enum

Public TypeConfirm As eTypeConfirm

Private Sub ButtonNext_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not Confirm Then
        Call MsgBox("Debes aceptar los términos y condiciones.")
        Exit Sub
    End If
    
    If Len(txtPasswd.Text) < 5 Then
        Call MsgBox("Contraseña incorrecta.")
        Exit Sub
    End If
    
    If Len(txtPin.Text) < 5 Then
        Call MsgBox("Clave Pin incorrecta.")
        Exit Sub
    End If
    

    Account.Passwd = txtPasswd.Text
    Account.Key = txtPin.Text
    
    If TypeConfirm = eTypeConfirm.Venta Then
        Call WriteMercader_New(0, MercaderUser)
    Else
        Call WriteMercader_New(MercaderSelectedOffer, MercaderUserOffer)
    End If
    
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub chkConfirm_Click()
    
    Call Audio.PlayInterface(SND_CLICK)
     
    If Confirm Then
        Confirm = False
        Set chkConfirm.Picture = picCheckBoxNulo
    Else
        Confirm = True
        chkConfirm.Picture = picCheckBox
        
    End If
    
End Sub

Private Sub chkView1_Click()
    Call Audio.PlayInterface(SND_CLICK)
     
    If txtPasswd.PasswordChar = "*" Then
        txtPasswd.PasswordChar = vbNullString
        chkView1.Picture = picCheckBox
        
    Else
        txtPasswd.PasswordChar = "*"
        Set chkView1.Picture = picCheckBoxNulo
    End If
    
End Sub
Private Sub chkView2_Click()
    Call Audio.PlayInterface(SND_CLICK)
     
    If txtPin.PasswordChar = "*" Then
        txtPin.PasswordChar = vbNullString
        chkView2.Picture = picCheckBox
        
    Else
        txtPin.PasswordChar = "*"
        Set chkView2.Picture = picCheckBoxNulo
        
    End If
    
End Sub
Private Sub Form_Load()

    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "Mercader_Confirm.jpg")
    
    Call LoadButtons
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    ImgSecure.ToolTipText = "Si ofreces DSP el MERCADO se encargar  de transferirle el dinero correspondiente al personaje. Adem s podr  optar por quedarse los DSP y usarlos en la compra de otros personajes del MERCADO."
    
    Confirm = False
    
    UpdateInfo
    
End Sub

Public Sub UpdateInfo()

    Dim A As Long
    
    Dim TempA As Long
    
    With MercaderList(MercaderSelected)
        If TypeConfirm = eTypeConfirm.Oferta Then
            lblTitle.Caption = "OFERTA a» " + .Chars(1).Desc & IIf(.Char > 1, " +" & .Char - 1 & " pjs", vbNullString)
            
            
            With MercaderUserOffer
             
                ' # Personajes que vas a vender
                Call AddtoRichTextBox(Console, "Personajes que vas a ofrecer: ", 255, 255, 255, True, False)
                
                For A = 1 To ACCOUNT_MAX_CHARS
                    If .bChars(A) > 0 Then
                        TempA = TempA + 1
                        Call AddtoRichTextBox(Console, TempA & "°: " & Account.Chars(A).Name, 172, 246, 180, True, False)
                    End If
                Next A
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                
                ' # Monedas de Oro que vas a pedir
                Call AddtoRichTextBox(Console, "Monedas de Oro que ofreces: ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, PonerPuntos(.Gld), 255, 255, 0, True, False, False)
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                 
                ' # DSP que vas a pedir
                Call AddtoRichTextBox(Console, "DSP que ofreces: ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, PonerPuntos(.Dsp), 255, 138, 0, True, False, False)
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                
                ' # Descripción que tendrá tu publicación
                Call AddtoRichTextBox(Console, "La persona podrá retirar los DSP en DINERO AR$ permitiendo así una COMPRA-VENTA segura.", 189, 243, 243, False, False)
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, "La persona debe reclamar dentro de las 48hs hábiles. El servidor recauda el 20% de los DSP en caso de querer retirarlos por DINERO AR$.", 189, 243, 243, False, False)
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                
                Call AddtoRichTextBox(Console, "Descripción elegida (Cuentale en pocas palabras que le das): ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, .Desc, 191, 197, 249, True, False, False)
                
                
                
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, "FIN de la información. Scrollea el texto para leerlo bien.", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
            End With
            
        Else
            lblTitle.Caption = "Estás realizando una nueva publicación"
            
             With MercaderUser
             
                ' # Personajes que vas a vender
                Call AddtoRichTextBox(Console, "Personajes que vas a publicar: ", 255, 255, 255, True, False)
                
                For A = 1 To ACCOUNT_MAX_CHARS
                    If .bChars(A) > 0 Then
                        Call AddtoRichTextBox(Console, A & "°: " & Account.Chars(A).Name, 172, 246, 180, True, False)
                    End If
                Next A
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                
                ' # Monedas de Oro que vas a pedir
                Call AddtoRichTextBox(Console, "Monedas de Oro mínimas a recibir: ", 255, 255, 255, True, False, False)
                Call AddtoRichTextBox(Console, PonerPuntos(.Gld), 255, 255, 0, True, False, False)
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                 
                ' # DSP que vas a pedir
                Call AddtoRichTextBox(Console, "DSP mínimos a recibir: ", 255, 255, 255, True, False, False)
                Call AddtoRichTextBox(Console, PonerPuntos(.Dsp), 255, 138, 0, True, False, False)
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                
                ' # Descripción que tendrá tu publicación
                Call AddtoRichTextBox(Console, "Esto es lo que leeran los demás usuarios. Aprovecha este espacio para buscar una clase específica o solicitar una cierta cantidad de monedas arriba. ¡Usalo de manera respetuosa!", 189, 243, 243, False, False)
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                
                Call AddtoRichTextBox(Console, "Descripción elegida: ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, .Desc, 191, 197, 249, True, False, False)
                
                
                
                
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, "FIN de la información. Scrollea el texto para leerlo bien.", 255, 255, 255, True, False)
                Call AddtoRichTextBox(Console, " ", 255, 255, 255, True, False)
            End With
        End If
        
    End With

   
    
    
    
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String
    
    GrhPath = DirInterface
    Set picCheckBox = LoadPicture(DirInterface & "options\CheckBoxOpciones.jpg")
    Set picCheckBoxNulo = LoadPicture(DirInterface & "options\CheckBoxOpcionesNulo.jpg")
    
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Form_KeyDown vbKeyEscape, 0
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

