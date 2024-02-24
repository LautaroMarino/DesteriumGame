VERSION 5.00
Begin VB.Form FrmMercaderInfoOffer 
   BorderStyle     =   0  'None
   Caption         =   "Mercado"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercaderInfoOffer.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicMercader 
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
      Height          =   1680
      Left            =   795
      MousePointer    =   99  'Custom
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   242
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3630
   End
   Begin VB.Image ButtonReturn 
      Height          =   375
      Left            =   600
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblChars 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   195
      Left            =   3990
      TabIndex        =   5
      Top             =   1605
      Width           =   450
   End
   Begin VB.Image ImgInfo 
      Height          =   375
      Left            =   720
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblDsp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   5085
      Width           =   1890
   End
   Begin VB.Label lblGld 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   5085
      Width           =   1890
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion de la publicacion"
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
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   5760
      Width           =   3570
   End
   Begin VB.Image ButtonAccept 
      Height          =   510
      Left            =   3390
      Picture         =   "FrmMercaderInfoOffer.frx":15EF3
      Top             =   6480
      Width           =   1410
   End
   Begin VB.Image ImgSecure 
      Height          =   255
      Left            =   3600
      MousePointer    =   14  'Arrow and Question
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo de la Oferta"
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
      TabIndex        =   0
      Top             =   1185
      Width           =   3570
   End
   Begin VB.Image ButtonRechace 
      Height          =   510
      Left            =   1875
      Picture         =   "FrmMercaderInfoOffer.frx":1646F
      Top             =   6480
      Width           =   1410
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMercaderInfoOffer.frx":169EB
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMercaderInfoOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario          As clsFormMovementManager

Private ListMercader As clsGraphicalList

Dim SelectedOffer As Integer

Private Sub ButtonAdd_Click()

  
    Call Audio.PlayInterface(SND_CLICK)
    
End Sub

Public Sub UpdateInfo()
    
    Dim A As Long

    Buttons_Update
    
    With MercaderListOffer(MercaderSelectedOffer)
        lblTitle.Caption = .Chars(1).Desc & IIf(.Char > 1, " +" & .Char - 1 & " pjs", vbNullString)
        
        lblGld.Caption = PonerPuntos(.Gld)
        lblDsp.Caption = PonerPuntos(.Dsp)
        
        lblInfo.Caption = .Desc
        
        lblChars.Caption = .Char
        
        For A = 1 To ACCOUNT_MAX_CHARS
            If .Chars(A).Class > 0 Then
                ListMercader.AddItem .Chars(A).Name & .Chars(A).Desc
            End If
        Next A
        
    End With
    
End Sub

Private Sub ButtonAccept_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
    If SelectedOffer = 0 Then
        Call MsgBox("Selecciona una oferta.")
        Exit Sub
    End If
    
    If MsgBox(" Estas seguro que deseas aceptar la oferta?  Una vez aceptada no hay vuelta atr s!", vbYesNo) = vbYes Then
        Call WriteMercader_Required(5, SelectedOffer, 0)
    End If
End Sub
Private Function Mercader_Prepare_List(ByVal A As Long) As String
    
    Mercader_Prepare_List = MercaderListOffer(A).Chars(1).Desc & IIf(MercaderListOffer(A).Char > 1, " +" & MercaderListOffer(A).Char - 1 & " pjs", vbNullString)
    
End Function

Private Sub ButtonReturn_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
    MercaderOff = 3
    Call WriteMercader_Required(MercaderOff, MercaderSelected, 0)
    
    Unload Me
End Sub

Private Sub Form_Load()

    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "Mercader_Info_Offer.jpg")
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Set ListMercader = New clsGraphicalList
    
    Call ListMercader.Initialize(PicMercader, RGB(200, 190, 190))

    UpdateInfo
    
    ImgSecure.ToolTipText = "Tienes 48hs para reclamar el cambio de DSP a DINERO REAL."
End Sub

Public Sub Buttons_Update()
    
    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    
    If MercaderSelected = MercaderUserSlot Then
        Set ButtonRechace.Picture = Nothing
        ButtonRechace.Enabled = True
        
        Set ButtonAccept.Picture = Nothing
        ButtonAccept.Enabled = True
        
    Else
        ButtonRechace.Picture = LoadPicture(filePath & "NoButton2.jpg")
        ButtonRechace.Enabled = False
        
        ButtonAccept.Picture = LoadPicture(filePath & "NoButton2.jpg")
        ButtonAccept.Enabled = False
    End If
    
    
    MercaderSelected1 = MercaderSelected
    MercaderSelectedOffer1 = MercaderSelectedOffer
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



' Lista Gr fica de Hechizos
Private Sub PicMercader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 0 Then Y = 0
    
    If Y > Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1 Then Y = Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1
    
    If X < PicMercader.ScaleWidth - 10 Then
        ListMercader.ListIndex = Int(Y / ListMercader.Pixel_Alto) + ListMercader.Scroll
        ListMercader.DownBarrita = 0
    
    Else
        ListMercader.DownBarrita = Y - ListMercader.Scroll * (PicMercader.ScaleHeight - ListMercader.BarraHeight) / (ListMercader.ListCount - ListMercader.VisibleCount)
    End If
    
End Sub

Private Sub PicMercader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    
    If yy < 0 Then yy = 0
    
    If yy > Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1 Then yy = Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1
    
    If ListMercader.DownBarrita > 0 Then
        ListMercader.Scroll = (Y - ListMercader.DownBarrita) * (ListMercader.ListCount - ListMercader.VisibleCount) / (PicMercader.ScaleHeight - ListMercader.BarraHeight)
    Else
        ListMercader.ListIndex = Int(yy / ListMercader.Pixel_Alto) + ListMercader.Scroll
    End If
ElseIf Button = 0 Then
    ListMercader.ShowBarrita = X > PicMercader.ScaleWidth - ListMercader.BarraWidth * 2
End If
End Sub

Private Sub PicMercader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListMercader.DownBarrita = 0
End Sub

Private Sub PicMercader_Click()
    If ListMercader.ListIndex = -1 Then Exit Sub
    
    
End Sub
