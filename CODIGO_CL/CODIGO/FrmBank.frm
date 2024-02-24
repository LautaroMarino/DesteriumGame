VERSION 5.00
Begin VB.Form FrmBank 
   BorderStyle     =   0  'None
   Caption         =   "Finanzas Goliath"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgDeposito 
      Height          =   495
      Left            =   525
      Top             =   3600
      Width           =   5595
   End
   Begin VB.Image imgMercado 
      Height          =   495
      Left            =   525
      Top             =   3030
      Width           =   5595
   End
   Begin VB.Image imgSubasta 
      Height          =   495
      Left            =   525
      Top             =   2475
      Width           =   5595
   End
   Begin VB.Image imgBancoCompartido 
      Height          =   495
      Left            =   525
      Top             =   1905
      Width           =   5595
   End
   Begin VB.Image imgBanco 
      Height          =   495
      Left            =   525
      Top             =   1350
      Width           =   5595
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   6315
      MouseIcon       =   "FrmBank.frx":000C
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario         As clsFormMovementManager
    
Private cBotonCerrar          As clsGraphicalButton
Private cBotonBanco           As clsGraphicalButton
Private cBotonBancoCompartido As clsGraphicalButton
Private cBotonSubastas        As clsGraphicalButton
Private cBotonMercado         As clsGraphicalButton
Private cBotonDeposito        As clsGraphicalButton

Public LastButtonPressed      As clsGraphicalButton

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\resource\interface\bank\bank.jpg")

    Call LoadButtons
End Sub

Private Sub LoadButtons()
 Dim GrhPath As String
    GrhPath = DirInterface

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonSubastas = New clsGraphicalButton
    Set cBotonBanco = New clsGraphicalButton
    Set cBotonBancoCompartido = New clsGraphicalButton
    Set cBotonMercado = New clsGraphicalButton
    Set cBotonDeposito = New clsGraphicalButton
  
    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCerrar.Initialize(imgUnload, vbNullString, GrhPath & "generic\BotonCerrarActivo.jpg", vbNullString, Me)
    Call cBotonSubastas.Initialize(imgSubasta, vbNullString, GrhPath & "bank\SubastaActivo.jpg", vbNullString, Me)
    Call cBotonBanco.Initialize(imgBanco, vbNullString, GrhPath & "bank\BancoPersonalActivo.jpg", vbNullString, Me)
    Call cBotonBancoCompartido.Initialize(imgBancoCompartido, vbNullString, GrhPath & "bank\BancoCompartidoActivo.jpg", vbNullString, Me)
    Call cBotonMercado.Initialize(imgMercado, vbNullString, GrhPath & "bank\MercadoActivo.jpg", vbNullString, Me)
    Call cBotonDeposito.Initialize(imgDeposito, vbNullString, GrhPath & "bank\DepositoActivo.jpg", vbNullString, Me)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgBanco_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteBankStart(E_BANK.e_User)
End Sub

Private Sub imgBancoCompartido_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteBankStart(E_BANK.e_Account)
End Sub

Private Sub imgMercado_Click()
    Call Audio.PlayInterface(SND_CLICK)
    MercaderSelected = ePanelInitial
    Call frmMercader.Show(, FrmMain)
End Sub

Private Sub imgSubasta_Click()
    Call Audio.PlayInterface(SND_CLICK)
    FrmSubasta.Show , FrmMain
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    FrmMain.SetFocus
    Unload Me
End Sub

