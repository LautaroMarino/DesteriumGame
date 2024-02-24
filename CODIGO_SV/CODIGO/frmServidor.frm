VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Configuración del Servidor"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Backup"
      Height          =   1815
      Left            =   120
      TabIndex        =   33
      Top             =   5520
      Width           =   4215
      Begin VB.CommandButton cmdLoadWorldBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cargar Mapas"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdCharBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar Chars"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdWorldBackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar Mapas"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Guardar Archivos"
      Height          =   2280
      Left            =   6960
      TabIndex        =   30
      Top             =   2520
      Width           =   2040
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guilds"
         Height          =   375
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.Frame fGenerals 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Chequeos Generales ¡¡CUIDADO!!"
      Height          =   2925
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   4200
      Begin VB.CheckBox chkSkins 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Skins para comprar-usar"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2640
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkEvents 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Eventos Automáticos JARVIS"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkSubastas 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Subastas de Objetos"
         Height          =   255
         Left            =   210
         TabIndex        =   28
         Top             =   2160
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkCastillo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Conquista de Castillos"
         Height          =   255
         Left            =   210
         TabIndex        =   27
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkCrafting 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Crafting/Herrero"
         Height          =   255
         Left            =   210
         TabIndex        =   26
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkInvocaciones 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Invocaciones mágicas"
         Height          =   255
         Left            =   210
         TabIndex        =   25
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkRetosFast 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enfrentamientos Random (Retos Rapidos)"
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   1200
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkRetos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enfrentamientos Privados (Retos)"
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   960
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkValidation 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Validacion de Personajes"
         Height          =   255
         Left            =   210
         TabIndex        =   21
         Top             =   720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkServerHabilitado 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Server Habilitado Solo Gms"
         Height          =   255
         Left            =   210
         TabIndex        =   20
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox chkMao 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mercado de Personjes"
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   480
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdReiniciar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Administración"
      Height          =   3720
      Left            =   4410
      TabIndex        =   10
      Top             =   2415
      Width           =   2415
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Panel Creator"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdResetListen 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reset Listen"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdDebugUserlist 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Debug UserList"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton frmAdministracion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Administración"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdPausarServidor 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pausar el servidor"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdConfigIntervalos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Config. Intervalos"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir (Esc)"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recargar"
      Height          =   2265
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdMessages 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MENSAJES"
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmbEvents 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EVENTOS"
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "INVASIONES"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DROPS"
         Height          =   375
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Quests"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmbCharsShop 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chars Shop"
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmbShop 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Shop"
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guilds"
         Height          =   375
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1785
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EVENTS_LIST.ini"
         Height          =   375
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1260
         Width           =   1935
      End
      Begin VB.CommandButton cmbCofres 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cofres.dat"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarAdministradores 
         BackColor       =   &H0080C0FF&
         Caption         =   "Administradores"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarGuardiasPosOrig 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardias en pos originales"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1785
         Width           =   3015
      End
      Begin VB.CommandButton cmdRecargarMOTD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOTD"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdRecargarServerIni 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Server.ini"
         Height          =   375
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarNombresInvalidos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NombresInvalidos.txt"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   1965
      End
      Begin VB.CommandButton cmdRecargarNPCs 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Npcs.dat"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarBalance 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Balance.dat"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarHechizos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hechizos.dat"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdRecargarObjetos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Obj.dat"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Sub chkCastillo_Click()
    ConfigServer.ModoCastillo = chkCastillo.Value
End Sub

Private Sub chkCrafting_Click()
    ConfigServer.ModoCrafting = chkCrafting.Value
End Sub

Private Sub chkEvents_Click()
    If Events_Automatic.Events_Automatic_Active = 0 Then
        Events_Automatic.Events_Automatic_Active = 1
    Else
        Events_Automatic.Events_Automatic_Active = 0
    End If
    
 '   If Events_Automatic.Events_Automatic_Active Then
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Jarvis ha comenzado a hacer eventos automáticos. ¡Cada 15 MINUTOS te sorprenderá!", FontTypeNames.FONTTYPE_CRITICO))
   ' End If
    
End Sub

Private Sub chkInvocaciones_Click()
    ConfigServer.ModoInvocaciones = chkInvocaciones.Value
End Sub

Private Sub chkMao_Click()
    MercaderActivate = Not MercaderActivate
End Sub

Private Sub chkRetos_Click()
    ConfigServer.ModoRetos = chkRetos.Value
End Sub

Private Sub chkRetosFast_Click()
    ConfigServer.ModoRetosFast = chkRetosFast.Value
End Sub

Private Sub chkServerHabilitado_Click()
    ServerSoloGMs = chkServerHabilitado.Value
End Sub

Private Sub chkSubastas_Click()
    ConfigServer.ModoSubastas = chkSubastas.Value
End Sub
Private Sub chkSkins_Click()
    ConfigServer.ModoSkins = chkSkins.Value
End Sub
Private Sub chkValidation_Click()
    ValidacionDePjs = chkValidation.Value
End Sub

Private Sub cmbCharsShop_Click()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Shop» ¡Hay Nuevos Personajes a la Venta! Utiliza el Comando /SHOP para verlos.", FontTypeNames.FONTTYPE_INFOGREEN))
    Shop_Load_Chars
End Sub

Private Sub cmbCofres_Click()
    Call mDrop.Drops_Load
End Sub


Private Sub cmdResetSockets_Click()

End Sub

Private Sub cmbShop_Click()
    Call Shop_Load
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Shop» La Tienda ha sido actualizada. Utiliza el Comando /SHOP para verla.", FontTypeNames.FONTTYPE_INFOGREEN))
End Sub

Private Sub cmdMessages_Click()
   mMessages.MessageSpam_Load
    
    ' Testing
    'MessageSpam_SelectedRandom
End Sub

Private Sub Command1_Click()
    Call Events_Read
    
End Sub

Private Sub Command2_Click()
    Call mGuilds.Guilds_Save_All
End Sub

Private Sub Command3_Click()
    Call mGuilds.Guilds_Load
End Sub


Private Sub Command5_Click()
    Call LoadQuests
End Sub

Private Sub Command6_Click()
    Call mDrop.Drops_Load
End Sub

Private Sub Command7_Click()
    Invations_Load
End Sub

Private Sub Command8_Click()
    FrmPanelCreator.Show
End Sub

Private Sub cmbEvents_Click()
EventsAI.Events_Load_PreConfig
End Sub

Private Sub Form_Load()
    cmdResetListen.Visible = True
End Sub

Private Sub cmdApagarServer_Click()
    
    If MsgBox("¿Está seguro que desea hacer WorldSave, guardar pjs y cerrar?", vbYesNo, "Apagar Magicamente") = vbNo Then Exit Sub
    
    Me.MousePointer = 11
    
    FrmStat.Show
    
    ' Cancelamos retos
    Call Retos_Reset_All
    
    ' Cancelamos eventos
    Call Eventos_Reset_All
    
    ' Cancelamos Retos Fast
    Call Fast_Reset_All
    
    ' Cancelamos la Subasta
    Call Auction_Cancel
    
    ' Eventos automáticos
    Call Events_Data_Predetermined
    
    'commit experiencia
    Call DistributeExpAndGldGroups
    
    'WorldSave
    Call ES.DoBackUp
    
    Dim A As Long
    
    For A = 1 To LastUser
        Call Protocol.Kick(A)
    Next
    
    'Chauuu
    Unload frmMain

End Sub

Private Sub cmdCerrar_Click()
    frmServidor.Visible = False
End Sub

Private Sub cmdCharBackup_Click()
    Me.MousePointer = 11
    Call DistributeExpAndGldGroups
    Call GuardarUsuarios(False)
    Me.MousePointer = 0
    MsgBox "Grabado de personajes OK!"
End Sub

Private Sub cmdConfigIntervalos_Click()
    FrmInterv.Show
End Sub

Private Sub cmdDebugUserlist_Click()
    frmUserList.Show
End Sub

Private Sub cmdLoadWorldBackup_Click()
    Call CargarBackUp
End Sub

Private Sub cmdPausarServidor_Click()

    If EnPausa = False Then
        EnPausa = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        cmdPausarServidor.Caption = "Reanudar el servidor"
    Else
        EnPausa = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        cmdPausarServidor.Caption = "Pausar el servidor"
    End If

End Sub

Private Sub cmdRecargarBalance_Click()
    Call LoadBalance
End Sub

Private Sub cmdRecargarGuardiasPosOrig_Click()
    Call ReSpawnOrigPosNpcs
End Sub

Private Sub cmdRecargarHechizos_Click()
    Call CargarHechizos
End Sub

Private Sub cmdRecargarMOTD_Click()
    Call LoadMotd
End Sub

Private Sub cmdRecargarNombresInvalidos_Click()
    Call CargarForbidenWords
End Sub

Private Sub cmdRecargarNPCs_Click()
    Call CargaNpcsDat
End Sub

Private Sub cmdRecargarObjetos_Click()
    Call Crafting_Reset
    Call LoadOBJData
End Sub

Private Sub cmdRecargarServerIni_Click()
    Call LoadSini
End Sub

Private Sub cmdReiniciar_Click()

    If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la pérdida de datos de los usarios. " & "¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbNo Then Exit Sub
    
    Me.Visible = False
    Call General.Restart

End Sub

Private Sub cmdResetListen_Click()

    'Cierra el socket de escucha
    Call Server.Close
    
    'Inicia el socket de escucha
    Call SocketConfig
End Sub

Private Sub cmdWorldBackup_Click()

    On Error GoTo ErrHandler

    Me.MousePointer = 11
    FrmStat.Show
    Call ES.DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
    
    Exit Sub

ErrHandler:
    Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Form_Deactivate()
    frmServidor.Visible = False
End Sub

Private Sub frmAdministracion_Click()
    Me.Visible = False
    frmAdmin.Show
End Sub

Private Sub cmdRecargarAdministradores_Click()
    loadAdministrativeUsers
End Sub

