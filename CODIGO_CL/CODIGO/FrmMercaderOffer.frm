VERSION 5.00
Begin VB.Form FrmMercaderOffer 
   BorderStyle     =   0  'None
   Caption         =   "Mercado"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmMercaderOffer.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicChars1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1700
      Left            =   2870
      MousePointer    =   99  'Custom
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1860
   End
   Begin VB.PictureBox PicChars 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1700
      Left            =   600
      MousePointer    =   99  'Custom
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1860
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   860
      MaxLength       =   160
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   5850
      Width           =   3645
   End
   Begin VB.TextBox txtDsp 
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
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   2880
      MaxLength       =   160
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "Chat"
      Top             =   5300
      Width           =   1850
   End
   Begin VB.TextBox txtGld 
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
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   600
      MaxLength       =   160
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "Chat"
      Top             =   5300
      Width           =   1850
   End
   Begin VB.Image ImgSecure 
      Height          =   255
      Left            =   4800
      MousePointer    =   14  'Arrow and Question
      Top             =   5280
      Width           =   255
   End
   Begin VB.Image ImgInfo 
      Height          =   255
      Left            =   2280
      MousePointer    =   14  'Arrow and Question
      Top             =   5640
      Width           =   735
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
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   3570
   End
   Begin VB.Image ButtonRemove 
      Height          =   375
      Left            =   3000
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Image ButtonAdd 
      Height          =   375
      Left            =   720
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Image ButtonNext 
      Height          =   375
      Left            =   1680
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMercaderOffer.frx":1AB43
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMercaderOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario          As clsFormMovementManager

Private ListChars As clsGraphicalList
Private ListChars1 As clsGraphicalList


Private Sub ButtonAdd_Click()

  
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Char As String
    
    Char = ListChars.List(ListChars.ListIndex)
    
    If Char = "<Vacio>" Then Exit Sub
    
    If ListCharsExist(Char) Then
        Exit Sub
    End If
    
    If Account.Chars(ListChars.ListIndex + 1).Elv < MERCADER_MIN_LVL Then
        Call MsgBox("¡Solo puedes ofrecer personajes superiores a Nivel " & MERCADER_MIN_LVL & "!")
        Exit Sub
    End If
    
    If MercaderUserOffer.Char > 0 Then
        If Account.Premium < 3 Then
            Call MsgBox("Debes ser al menos Tier 3 para poder ofrecer mas de 1 personaje. ¡Mientras tanto podrás vender u ofrecer de a uno!")
            Exit Sub
    
        End If
    End If
    
    ListChars1.List(ListChars.ListIndex) = Char
    
    MercaderUserOffer.IDCHARS(Account.Chars(ListChars.ListIndex + 1).ID) = 1
    MercaderUserOffer.bChars(Account.Chars(ListChars.ListIndex + 1).ID) = 1
    MercaderUserOffer.Char = MercaderUserOffer.Char + 1
End Sub

Private Function ListCharsExist(ByVal Char As String) As Boolean
    Dim A As Long
    
    
    If ListChars1.ListIndex = -1 Then Exit Function
    
    For A = 0 To ListChars1.ListCount - 1
        If Char = UCase$(ListChars1.List(A)) Then
            ListCharsExist = True
            Exit Function
        End If
    Next A
End Function

Private Sub ButtonNext_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Len(txtDesc.Text) <= 4 Then
        Call MsgBox("Elige una descripción un poco más larga.")
        Exit Sub
    End If
    
    If Len(txtDesc.Text) >= 30 Then
        Call MsgBox("Elige una descripción un poco más corta.")
        Exit Sub
    End If
    
    If Val(txtGld.Text) < 0 Then txtGld.Text = "0"
    If Val(txtDsp.Text) < 0 Then txtDsp.Text = "0"
    
     If MercaderList_Copy(MercaderSelected).Gld > Account.Gld Then
        Call MsgBox("El usuario requiere como mínimo " & Format(MercaderList_Copy(MercaderSelected).Gld, "##,##") & " Monedas de Oro.")
        Exit Sub

    End If
        
    MercaderUserOffer.Gld = Val(txtGld.Text)
    MercaderUserOffer.Dsp = Val(txtDsp.Text)
    MercaderUserOffer.Desc = txtDesc.Text
    
    
    FrmMercaderConfirm.TypeConfirm = eTypeConfirm.Oferta
    
    FrmMercaderConfirm.Show , FrmMain
End Sub

Private Sub ButtonRemove_Click()

  
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Char As String
    
    Char = ListChars1.List(ListChars1.ListIndex)
    
    If Char = "<Vacio>" Then Exit Sub

    ListChars1.List(ListChars1.ListIndex) = "<Vacio>"
    
    MercaderUserOffer.IDCHARS(Account.Chars(ListChars1.ListIndex + 1).ID) = 0
    MercaderUserOffer.bChars(Account.Chars(ListChars1.ListIndex + 1).ID) = 0
    MercaderUserOffer.Char = MercaderUserOffer.Char - 1
    

End Sub

Private Sub ResetMercaderOffer()
    Dim A As Long
    
    For A = 1 To ACCOUNT_MAX_CHARS
        MercaderUserOffer.bChars(A) = 0
    Next A
    
    MercaderUserOffer.Gld = 0
    MercaderUserOffer.Dsp = 0
    MercaderUserOffer.Desc = ""
    MercaderUserOffer.Char = 0
    MercaderUserOffer.Blocked = 0
End Sub
Public Sub UpdateInfo()
    
    Dim A As Long
    
    Call ResetMercaderOffer
    
    With MercaderList_Copy(MercaderSelected)
        lblTitle.Caption = .Chars(1).Desc & IIf(.Char > 1, " +" & .Char - 1 & " pjs", vbNullString)
    End With
    
End Sub
Private Sub Form_Load()

    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "Mercader_Offer.jpg")
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Set ListChars = New clsGraphicalList
    Set ListChars1 = New clsGraphicalList
    
    Call ListChars.Initialize(PicChars, RGB(200, 190, 190))
    Call ListChars1.Initialize(PicChars1, RGB(200, 190, 190))
    
    Dim A As Long
    
    For A = 1 To ACCOUNT_MAX_CHARS
        If Account.Chars(A).Class > 0 Then
            ListChars.AddItem UCase$(Account.Chars(A).Name)
        Else
            ListChars.AddItem "<Vacio>"
        End If
        
        ListChars1.AddItem "<Vacio>"
    Next A
    
    UpdateInfo
    
    ImgInfo.ToolTipText = "Pon algo como 'Te ofrezco mi mago y 250.000 Monedas de Oro'. Cualquier palabra considerada ofensiva será castigada."
    ImgSecure.ToolTipText = "Si ofreces DSP el MERCADO se encargará de transferirle el dinero correspondiente al personaje. Además podrá optar por quedarse los DSP y usarlos en la compra de otros personajes del MERCADO."
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



' Lista Gráfica de Hechizos
Private Sub PicChars_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 0 Then Y = 0
    
    If Y > Int(PicChars.ScaleHeight / ListChars.Pixel_Alto) * ListChars.Pixel_Alto - 1 Then Y = Int(PicChars.ScaleHeight / ListChars.Pixel_Alto) * ListChars.Pixel_Alto - 1
    
    If X < PicChars.ScaleWidth - 10 Then
        ListChars.ListIndex = Int(Y / ListChars.Pixel_Alto) + ListChars.Scroll
        ListChars.DownBarrita = 0
    
    Else
        ListChars.DownBarrita = Y - ListChars.Scroll * (PicChars.ScaleHeight - ListChars.BarraHeight) / (ListChars.ListCount - ListChars.VisibleCount)
    End If
    
End Sub

Private Sub PicChars_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    
    If yy < 0 Then yy = 0
    
    If yy > Int(PicChars.ScaleHeight / ListChars.Pixel_Alto) * ListChars.Pixel_Alto - 1 Then yy = Int(PicChars.ScaleHeight / ListChars.Pixel_Alto) * ListChars.Pixel_Alto - 1
    
    If ListChars.DownBarrita > 0 Then
        ListChars.Scroll = (Y - ListChars.DownBarrita) * (ListChars.ListCount - ListChars.VisibleCount) / (PicChars.ScaleHeight - ListChars.BarraHeight)
    Else
        ListChars.ListIndex = Int(yy / ListChars.Pixel_Alto) + ListChars.Scroll
    End If
ElseIf Button = 0 Then
    ListChars.ShowBarrita = X > PicChars.ScaleWidth - ListChars.BarraWidth * 2
End If
End Sub

Private Sub PicChars_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListChars.DownBarrita = 0
End Sub


' # Chars 2
' Lista Gráfica de Hechizos
Private Sub PicChars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 0 Then Y = 0
    
    If Y > Int(PicChars1.ScaleHeight / ListChars1.Pixel_Alto) * ListChars1.Pixel_Alto - 1 Then Y = Int(PicChars1.ScaleHeight / ListChars1.Pixel_Alto) * ListChars1.Pixel_Alto - 1
    
    If X < PicChars1.ScaleWidth - 10 Then
        ListChars1.ListIndex = Int(Y / ListChars1.Pixel_Alto) + ListChars1.Scroll
        ListChars1.DownBarrita = 0
    
    Else
        ListChars1.DownBarrita = Y - ListChars1.Scroll * (PicChars1.ScaleHeight - ListChars1.BarraHeight) / (ListChars1.ListCount - ListChars1.VisibleCount)
    End If
    
End Sub

Private Sub PicChars1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    
    If yy < 0 Then yy = 0
    
    If yy > Int(PicChars1.ScaleHeight / ListChars1.Pixel_Alto) * ListChars1.Pixel_Alto - 1 Then yy = Int(PicChars1.ScaleHeight / ListChars1.Pixel_Alto) * ListChars1.Pixel_Alto - 1
    
    If ListChars1.DownBarrita > 0 Then
        ListChars1.Scroll = (Y - ListChars1.DownBarrita) * (ListChars1.ListCount - ListChars1.VisibleCount) / (PicChars1.ScaleHeight - ListChars1.BarraHeight)
    Else
        ListChars1.ListIndex = Int(yy / ListChars1.Pixel_Alto) + ListChars1.Scroll
    End If
ElseIf Button = 0 Then
    ListChars1.ShowBarrita = X > PicChars1.ScaleWidth - ListChars1.BarraWidth * 2
End If
End Sub

Private Sub PicChars1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListChars1.DownBarrita = 0
End Sub


Private Sub txtGld_Change()
    
    If Not IsNumeric(txtGld.Text) Then
        txtGld.Text = "0"
    End If
    
    If Val(txtGld.Text) < 0 Then
        txtGld.Text = "0"
    End If
    
    If Val(txtGld.Text) > MERCADER_MAX_GLD Then
        txtGld.Text = MERCADER_MAX_GLD
        txtGld.SelStart = Len(txtGld.Text)
    End If
End Sub
Private Sub txtDsp_Change()
    
    If Not IsNumeric(txtDsp.Text) Then
        txtDsp.Text = "0"
    End If
    
    If Val(txtDsp.Text) < 0 Then
        txtDsp.Text = "0"
    End If
    
    If Val(txtDsp.Text) > MERCADER_MAX_DSP Then
        txtDsp.Text = MERCADER_MAX_DSP
        txtDsp.SelStart = Len(txtDsp.Text)
    End If
End Sub
