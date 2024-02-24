VERSION 5.00
Begin VB.Form FrmMercaderOffers 
   BorderStyle     =   0  'None
   Caption         =   "Mercado"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercaderOffers.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      HasDC           =   0   'False
      Height          =   3075
      Left            =   540
      MousePointer    =   99  'Custom
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   282
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1650
      Width           =   4230
   End
   Begin VB.Image ButtonView 
      Height          =   495
      Left            =   480
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Image ButtonAccept 
      Height          =   510
      Left            =   3480
      Picture         =   "FrmMercaderOffers.frx":1A0D3
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Image ImgSecure 
      Height          =   1455
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      Top             =   5760
      Width           =   4695
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
      TabIndex        =   1
      Top             =   1230
      Width           =   3570
   End
   Begin VB.Image ButtonRechace 
      Height          =   510
      Left            =   2070
      Picture         =   "FrmMercaderOffers.frx":1A64F
      Top             =   4830
      Width           =   1410
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMercaderOffers.frx":1ABCB
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMercaderOffers"
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

   
    
    With MercaderList_Copy(MercaderSelected)
        lblTitle.Caption = .Chars(1).Desc & IIf(.Char > 1, " +" & .Char - 1 & " pjs", vbNullString)
    End With
    
    Reelistar_Ofertas
    
    MirandoOffer = True
End Sub

Private Sub ButtonAccept_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
    If SelectedOffer = 0 Then
        Call MsgBox("Selecciona una oferta.")
        Exit Sub
    End If
    
    If MsgBox("¿Estás seguro que deseas aceptar la oferta? ¡Una vez aceptada no hay vuelta atrás!", vbYesNo) = vbYes Then
        Call WriteMercader_Required(5, SelectedOffer, 0)
    End If
End Sub
Private Function Mercader_Prepare_List(ByVal A As Long) As String
    
    Mercader_Prepare_List = MercaderListOffer(A).Chars(1).Desc & IIf(MercaderListOffer(A).Char > 1, " +" & MercaderListOffer(A).Char - 1 & " pjs", vbNullString)
    
End Function

Private Sub Reelistar_Ofertas()

    ListMercader.Clear
    
    Dim A As Long
    
    For A = 1 To MERCADER_MAX_LIST
        With MercaderListOffer(A)
            If .Desc <> vbNullString Then
                ListMercader.AddItem Mercader_Prepare_List(A) & " [" & .Desc & "]"
            End If
        End With

    Next A
    
    
    If ListMercader.ListCount > 0 Then
        MercaderSelectedOffer = 1
    Else
        MercaderSelectedOffer = -1
    End If
    
End Sub

Private Sub ButtonView_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If ListMercader.ListIndex = -1 Then
        Exit Sub
    End If
    
    
    If FrmMercaderInfoOffer.visible Then
        FrmMercaderInfoOffer.UpdateInfo
    Else
        FrmMercaderInfoOffer.Show , FrmMain
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "Mercader_Offers.jpg")
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Set ListMercader = New clsGraphicalList
    
    Call ListMercader.Initialize(PicChars, RGB(200, 190, 190))
    
    UpdateInfo
    
    MirandoOffer = True
    ImgSecure.ToolTipText = "Tienes 48hs para reclamar el cambio de DSP a DINERO REAL."
End Sub



Private Sub Form_Unload(Cancel As Integer)
    MirandoOffer = False
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



Private Sub PicChars_Click()

    If ListMercader.ListIndex = -1 Then Exit Sub
    MercaderSelectedOffer = ListMercader.ListIndex + 1
End Sub

' Lista Gr fica de Hechizos
Private Sub PicChars_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 0 Then Y = 0
    
    If Y > Int(PicChars.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1 Then Y = Int(PicChars.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1
    
    If X < PicChars.ScaleWidth - 10 Then
        ListMercader.ListIndex = Int(Y / ListMercader.Pixel_Alto) + ListMercader.Scroll
        ListMercader.DownBarrita = 0
    
    Else
        ListMercader.DownBarrita = Y - ListMercader.Scroll * (PicChars.ScaleHeight - ListMercader.BarraHeight) / (ListMercader.ListCount - ListMercader.VisibleCount)
    End If
    
End Sub

Private Sub PicChars_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    
    If yy < 0 Then yy = 0
    
    If yy > Int(PicChars.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1 Then yy = Int(PicChars.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1
    
    If ListMercader.DownBarrita > 0 Then
        ListMercader.Scroll = (Y - ListMercader.DownBarrita) * (ListMercader.ListCount - ListMercader.VisibleCount) / (PicChars.ScaleHeight - ListMercader.BarraHeight)
    Else
        ListMercader.ListIndex = Int(yy / ListMercader.Pixel_Alto) + ListMercader.Scroll
    End If
ElseIf Button = 0 Then
    ListMercader.ShowBarrita = X > PicChars.ScaleWidth - ListMercader.BarraWidth * 2
End If
End Sub

Private Sub PicChars_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListMercader.DownBarrita = 0
End Sub
