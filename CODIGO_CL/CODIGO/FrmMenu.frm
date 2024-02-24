VERSION 5.00
Begin VB.Form FrmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Menu Principal"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgMercado 
      Height          =   660
      Left            =   1935
      Top             =   2385
      Width           =   690
   End
   Begin VB.Image imgManual 
      Height          =   660
      Left            =   1230
      Top             =   2385
      Width           =   690
   End
   Begin VB.Image imgShop 
      Height          =   660
      Left            =   525
      Top             =   2385
      Width           =   690
   End
   Begin VB.Label lblMenuDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   555
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   4290
   End
   Begin VB.Image imgSkills 
      Height          =   660
      Left            =   1230
      Top             =   1695
      Width           =   690
   End
   Begin VB.Image imgGrupo 
      Height          =   660
      Left            =   525
      Top             =   1695
      Width           =   690
   End
   Begin VB.Image imgStats 
      Height          =   660
      Left            =   3345
      Top             =   2385
      Width           =   705
   End
   Begin VB.Image imgQuests 
      Height          =   660
      Left            =   1935
      Top             =   1695
      Width           =   690
   End
   Begin VB.Image imgRetos 
      Height          =   660
      Left            =   2640
      Top             =   1695
      Width           =   690
   End
   Begin VB.Image imgEventos 
      Height          =   660
      Left            =   3345
      Top             =   1695
      Width           =   690
   End
   Begin VB.Image imgRanking 
      Height          =   660
      Left            =   4050
      Top             =   1695
      Width           =   690
   End
   Begin VB.Image imgGuild 
      Height          =   660
      Left            =   2640
      Top             =   2385
      Width           =   705
   End
   Begin VB.Image imgOption 
      Height          =   660
      Left            =   4050
      Top             =   2385
      Width           =   690
   End
   Begin VB.Image imgChangeChar 
      Height          =   495
      Left            =   855
      Top             =   3735
      Width           =   3540
   End
   Begin VB.Image imgChangeAccount 
      Height          =   495
      Left            =   855
      Top             =   4335
      Width           =   3555
   End
   Begin VB.Image imgMinimice 
      Height          =   495
      Left            =   855
      Top             =   4935
      Width           =   3555
   End
   Begin VB.Image imgCerrarJuego 
      Height          =   495
      Left            =   850
      Top             =   5530
      Width           =   3555
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ButtonClose As clsGraphicalButton
Private ButtonMinimice As clsGraphicalButton
Private ButtonChangeAccount As clsGraphicalButton
Private ButtonChangeChar As clsGraphicalButton

Private ButtonParty As clsGraphicalButton
Private ButtonSkills As clsGraphicalButton
Private ButtonQuest As clsGraphicalButton
Private ButtonRetos As clsGraphicalButton
Private ButtonEventos As clsGraphicalButton
Private ButtonRanking As clsGraphicalButton
Private ButtonOpciones As clsGraphicalButton
Private ButtonStats As clsGraphicalButton
Private ButtonClanes As clsGraphicalButton
Private ButtonManual As clsGraphicalButton
Private ButtonMercado As clsGraphicalButton
Private ButtonShop As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton
Private clsFormulario          As clsFormMovementManager


Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = DirInterface & "menucompacto\Buttons\"

    Set ButtonClose = New clsGraphicalButton
    Set ButtonMinimice = New clsGraphicalButton
    Set ButtonChangeAccount = New clsGraphicalButton
    Set ButtonChangeChar = New clsGraphicalButton
   
    Set LastButtonPressed = New clsGraphicalButton

    Call ButtonClose.Initialize(imgCerrarJuego, vbNullString, GrhPath & "CerrarJuegoHover.jpg", vbNullString, Me)
    Call ButtonMinimice.Initialize(imgMinimice, vbNullString, GrhPath & "MinimizarHover.jpg", vbNullString, Me)
    Call ButtonChangeAccount.Initialize(imgChangeAccount, vbNullString, GrhPath & "CambiarCuentaHover.jpg", vbNullString, Me)
    Call ButtonChangeChar.Initialize(imgChangeChar, vbNullString, GrhPath & "CambiarPersonajeHover.jpg", vbNullString, Me)
    
    Set ButtonParty = New clsGraphicalButton
    Set ButtonSkills = New clsGraphicalButton
    Set ButtonQuest = New clsGraphicalButton
    Set ButtonRetos = New clsGraphicalButton
    Set ButtonEventos = New clsGraphicalButton
    Set ButtonRanking = New clsGraphicalButton
    Set ButtonOpciones = New clsGraphicalButton
    Set ButtonStats = New clsGraphicalButton
    Set ButtonClanes = New clsGraphicalButton
    Set ButtonManual = New clsGraphicalButton
    Set ButtonMercado = New clsGraphicalButton
    Set ButtonShop = New clsGraphicalButton
    
    Call ButtonParty.Initialize(imgGrupo, vbNullString, GrhPath & "PartyActivo.jpg", vbNullString, Me)
    Call ButtonSkills.Initialize(imgSkills, vbNullString, GrhPath & "SkillsActivo.jpg", vbNullString, Me)
    Call ButtonQuest.Initialize(imgQuests, vbNullString, GrhPath & "QuestActivo.jpg", vbNullString, Me)
    Call ButtonRetos.Initialize(imgRetos, vbNullString, GrhPath & "RetosActivo.jpg", vbNullString, Me)
    Call ButtonEventos.Initialize(imgEventos, vbNullString, GrhPath & "EventosActivo.jpg", vbNullString, Me)
    Call ButtonRanking.Initialize(imgRanking, vbNullString, GrhPath & "RankingActivo.jpg", vbNullString, Me)
    Call ButtonOpciones.Initialize(imgOption, vbNullString, GrhPath & "SettingsActivo.jpg", vbNullString, Me)
    Call ButtonStats.Initialize(imgStats, vbNullString, GrhPath & "StatsActivo.jpg", vbNullString, Me)
    Call ButtonClanes.Initialize(imgGuild, vbNullString, GrhPath & "ClanActivo.jpg", vbNullString, Me)
    Call ButtonManual.Initialize(imgManual, vbNullString, GrhPath & "ManualActivo.jpg", vbNullString, Me)
    Call ButtonMercado.Initialize(imgMercado, vbNullString, GrhPath & "MercadoActivo.jpg", vbNullString, Me)
    Call ButtonShop.Initialize(imgShop, vbNullString, GrhPath & "ShopActivo.jpg", vbNullString, Me)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
    
    lblMenuDesc.visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    #If ModoBig = 0 Then
    If KeyCode = vbKeyEscape Then
        FrmMain.SetFocus
        Unload Me
        
    End If
    #End If
End Sub


Private Sub Form_Load()
    
    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "menu.jpg")
    
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    LoadButtons
End Sub

Private Sub imgCerrarJuego_Click()
    Call Audio.PlayInterface(SND_CLICK)
     
    If MsgBox("¿Estás seguro que deseas cerrar el Juego por Completo?", vbYesNo) = vbYes Then
        prgRun = False
    End If
    
    #If ModoBig = 0 Then
        Unload Me
    #End If
    
End Sub

Private Sub imgChangeAccount_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    #If ModoBig = 0 Then
        Call FrmMenuAccount.Show(, FrmMain)
         Unload Me
    #Else
        dockForm FrmMenuAccount.hWnd, FrmMain.PicMenu, True
    #End If


End Sub

Private Sub imgChangeChar_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ParseUserCommand("/SALIR")
    
    #If ModoBig = 0 Then
        Unload Me
    #End If
    
End Sub







Private Sub imgMinimice_Click()
    Call Audio.PlayInterface(SND_CLICK)
    FrmMain.WindowState = 1
    
    #If ModoBig = 0 Then
        Unload Me
    #End If
End Sub


Private Sub imgGrupo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Forma grupos con los usuarios"
End Sub







Private Sub imgSkills_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Habilidades y Destrezas de tu Personaje"
End Sub
Private Sub imgQuests_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Misiones y Objetivos a completar"
End Sub
Private Sub imgRetos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Juega Luchas con los demás."
End Sub
Private Sub imgEventos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Eventos y Torneos Oficiales del Game"
End Sub
Private Sub imgRanking_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Premios Mensuales"
End Sub
Private Sub imgGuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "<Alianzas> entre personajes"
End Sub
Private Sub imgStats_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Estadisticas de tu Personaje"
End Sub
Private Sub imgOption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Configuración del Game"
End Sub
Private Sub imgManual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Manual del Juego para aprender a Jugar"
End Sub
Private Sub imgMercado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "¡Compra-Venta de Personajes!"
End Sub
Private Sub imgShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenuDesc.visible = True
    lblMenuDesc.Caption = "Compra de Items/Personajes/Avatares Y Mas.. Shop Oficial."
End Sub
Private Sub imgQuests_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call ParseUserCommand("/MISIONES")
End Sub

Private Sub imgSkills_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call WriteRequestSkills
End Sub
Private Sub imgRetos_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call ParseUserCommand("/RETOS")
End Sub
Private Sub imgEventos_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call ParseUserCommand("/TORNEOS")
End Sub
Private Sub imgRanking_Click()
    Call Audio.PlayInterface(SND_CLICK)

    #If ModoBig = 1 Then
        dockForm FrmPremios.hWnd, FrmMain.PicMenu, True
    #Else
        Call FrmPremios.Show(, FrmMain)
    #End If

End Sub
Private Sub imgGuild_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteGuilds_Required(0)
End Sub

Private Sub imgStats_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ParseUserCommand("/EST")
End Sub

Private Sub imgOption_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    #If ModoBig = 1 Then
        dockForm frmOpciones.hWnd, FrmMain.PicMenu, True
    #Else
        Call frmOpciones.Show(, FrmMain)
    #End If

End Sub

Private Sub imgShop_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ParseUserCommand("/SHOP")
End Sub
Private Sub imgGrupo_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ShowConsoleMsg("Ayuda» Es hora de enviar solicitudes para que usuarios formen un grupo contigo.. Haz clic sobre aquel que desees invitar y luego teclea F3.", 150, 200, 148, True)
    Call WritePartyClient(1)
End Sub
Private Sub imgManual_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    
End Sub

Private Sub imgMercado_Click()
    Call Audio.PlayInterface(SND_CLICK)

    Call WriteMercader_Required(1, 1, 255)
    
End Sub
