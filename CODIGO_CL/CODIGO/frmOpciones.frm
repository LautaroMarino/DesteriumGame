VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wiki/Manual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   3
      Left            =   3600
      TabIndex        =   15
      Top             =   6360
      Width           =   1365
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "frmOpciones.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   3
      Left            =   4665
      Top             =   2175
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   2
      Left            =   4665
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   0
      Left            =   4665
      Top             =   1545
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   3
      Left            =   3885
      Top             =   2175
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   2
      Left            =   3885
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   0
      Left            =   3885
      Top             =   1545
      Width           =   270
   End
   Begin VB.Label lblSound 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   3
      Left            =   4245
      TabIndex        =   14
      Top             =   2190
      Width           =   315
   End
   Begin VB.Label lblSound 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   2
      Left            =   4245
      TabIndex        =   13
      Top             =   1890
      Width           =   315
   End
   Begin VB.Label lblSound 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   0
      Left            =   4245
      TabIndex        =   12
      Top             =   1575
      Width           =   315
   End
   Begin VB.Image imgMas 
      Height          =   255
      Index           =   1
      Left            =   4665
      Top             =   1230
      Width           =   270
   End
   Begin VB.Label lblSound 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   1
      Left            =   4245
      TabIndex        =   11
      Top             =   1260
      Width           =   315
   End
   Begin VB.Image imgMenos 
      Height          =   255
      Index           =   1
      Left            =   3885
      Top             =   1230
      Width           =   270
   End
   Begin VB.Image imgAlphaInfo 
      Height          =   375
      Left            =   1440
      Top             =   1080
      Width           =   975
   End
   Begin VB.Image imgFpsInfo 
      Height          =   375
      Left            =   480
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image PicAlpha 
      Height          =   225
      Index           =   3
      Left            =   1560
      Picture         =   "frmOpciones.frx":10BE
      Top             =   2445
      Width           =   210
   End
   Begin VB.Image PicAlpha 
      Height          =   225
      Index           =   2
      Left            =   1560
      Picture         =   "frmOpciones.frx":1F17
      Top             =   2130
      Width           =   210
   End
   Begin VB.Image PicAlpha 
      Height          =   225
      Index           =   1
      Left            =   1560
      Picture         =   "frmOpciones.frx":2D70
      Top             =   1785
      Width           =   210
   End
   Begin VB.Image PicAlpha 
      Height          =   225
      Index           =   0
      Left            =   1560
      Picture         =   "frmOpciones.frx":3BC9
      Top             =   1440
      Width           =   210
   End
   Begin VB.Image PicFPS 
      Height          =   225
      Index           =   3
      Left            =   360
      Picture         =   "frmOpciones.frx":4A22
      Top             =   2445
      Width           =   210
   End
   Begin VB.Image PicFPS 
      Height          =   225
      Index           =   2
      Left            =   360
      Picture         =   "frmOpciones.frx":587B
      Top             =   2130
      Width           =   210
   End
   Begin VB.Image PicFPS 
      Height          =   225
      Index           =   1
      Left            =   360
      Picture         =   "frmOpciones.frx":66D4
      Top             =   1785
      Width           =   210
   End
   Begin VB.Image PicFPS 
      Height          =   225
      Index           =   0
      Left            =   360
      Picture         =   "frmOpciones.frx":752D
      Top             =   1440
      Width           =   210
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colores Dialogos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   6975
      Width           =   2085
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensajes Personalizados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   6675
      Width           =   2745
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configurar Teclas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   6360
      Width           =   1890
   End
   Begin VB.Label lblInterfaz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interfaz TDS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   11
      Left            =   2640
      Picture         =   "frmOpciones.frx":8386
      Top             =   5880
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   10
      Left            =   360
      Picture         =   "frmOpciones.frx":91DF
      Top             =   5880
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   0
      Left            =   5880
      Picture         =   "frmOpciones.frx":A038
      Top             =   1320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblInterfaz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interfaz Moderna"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Index           =   0
      Left            =   2910
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblCursor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cursor\cursores7.ico"
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
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Top             =   3870
      Width           =   3120
   End
   Begin VB.Image imgAlpha 
      Height          =   225
      Index           =   1
      Left            =   6195
      Picture         =   "frmOpciones.frx":AE91
      Top             =   2520
      Width           =   240
   End
   Begin VB.Image imgAlpha 
      Height          =   225
      Index           =   0
      Left            =   5460
      Picture         =   "frmOpciones.frx":BE34
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label lblAlpha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
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
      Left            =   5460
      TabIndex        =   4
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblFps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Libres"
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
      Left            =   5565
      TabIndex        =   3
      Top             =   2205
      Width           =   720
   End
   Begin VB.Image imgFps 
      Height          =   225
      Index           =   1
      Left            =   6195
      Picture         =   "frmOpciones.frx":CDB7
      Top             =   2205
      Width           =   240
   End
   Begin VB.Image imgFps 
      Height          =   225
      Index           =   0
      Left            =   5460
      Picture         =   "frmOpciones.frx":DD5A
      Top             =   2205
      Width           =   240
   End
   Begin VB.Label lblCursor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cursor\cursores7.ico"
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
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   4170
      Width           =   3120
   End
   Begin VB.Label lblCursor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cursor\cursores7.ico"
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
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   3540
      Width           =   3120
   End
   Begin VB.Label lblCursor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "cursor\cursores7.ico"
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
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   3210
      Width           =   3120
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   9
      Left            =   5880
      Picture         =   "frmOpciones.frx":ECDD
      Top             =   360
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   8
      Left            =   2700
      Picture         =   "frmOpciones.frx":FB36
      Top             =   1200
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   7
      Left            =   2640
      Picture         =   "frmOpciones.frx":10A2E
      Top             =   5550
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   6
      Left            =   360
      Picture         =   "frmOpciones.frx":11887
      Top             =   5550
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   5
      Left            =   360
      Picture         =   "frmOpciones.frx":126E0
      Top             =   5220
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   4
      Left            =   2640
      Picture         =   "frmOpciones.frx":13539
      Top             =   5220
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   3
      Left            =   2700
      Picture         =   "frmOpciones.frx":14392
      Top             =   2520
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   2
      Left            =   510
      Picture         =   "frmOpciones.frx":151EB
      Top             =   4440
      Width           =   210
   End
   Begin VB.Image imgChkConfig 
      Height          =   225
      Index           =   1
      Left            =   5880
      Picture         =   "frmOpciones.frx":16044
      Top             =   840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgChkInterfaces 
      Height          =   225
      Left            =   2700
      Picture         =   "frmOpciones.frx":16E9D
      Top             =   2175
      Width           =   210
   End
   Begin VB.Image imgChkSonidos 
      Height          =   225
      Left            =   2700
      Picture         =   "frmOpciones.frx":17CF6
      Top             =   1530
      Width           =   210
   End
   Begin VB.Image imgChkMusica 
      Height          =   225
      Left            =   2700
      Picture         =   "frmOpciones.frx":18B4F
      Top             =   1860
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   4200
      Top             =   10080
      Width           =   210
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tierras Nórdicas 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Tierras Nórdicas is based on Baronsoft's VB6 Online RPG
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

Private clsFormulario        As clsFormMovementManager

Private picCheckBox          As Picture
Private picCheckBoxNulo      As Picture
Public LastButtonPressed           As clsGraphicalButton
Private cBotonCerrar         As clsGraphicalButton
Private cBotonPersonalizados As clsGraphicalButton
Private cBotonColores        As clsGraphicalButton
Private cBotonConfigTeclas   As clsGraphicalButton
Private Loading              As Boolean


Private Sub Form_Click()
    Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ISaveClientSetup
End Sub

Private Sub imgConfigTeclas_Click()
    Call Audio.PlayInterface(SND_CLICK)
     
    #If ModoBig = 1 Then
        dockForm frmCustomKeys.hWnd, FrmMain.PicMenu, True
    #Else
        Call frmCustomKeys.Show(, Me)
    #End If
     
End Sub
Private Sub imgMas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMas(Index).Picture = LoadPicture(DirInterface & "menucompacto\buttons\mas.jpg")
End Sub
Private Sub imgMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMenos(Index).Picture = LoadPicture(DirInterface & "menucompacto\buttons\menos.jpg")
End Sub
Private Sub imgMas_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Value As Byte
    
    Select Case Index
    
        Case 0 ' Master
            ClientSetup.bValueSoundMaster = ClientSetup.bValueSoundMaster + 10
            If ClientSetup.bValueSoundMaster > 100 Then ClientSetup.bValueSoundMaster = 100

            Audio.MasterVolume = ClientSetup.bValueSoundMaster
            lblSound(Index).Caption = Audio.MasterVolume
        Case 1 ' Efecto
            ClientSetup.bValueSoundEffect = ClientSetup.bValueSoundEffect + 10
            If ClientSetup.bValueSoundEffect > 100 Then ClientSetup.bValueSoundEffect = 100

            Audio.EffectVolume = ClientSetup.bValueSoundEffect
            lblSound(Index).Caption = Audio.EffectVolume
        Case 2 ' Musica
            ClientSetup.bValueSoundMusic = ClientSetup.bValueSoundMusic + 10
            If ClientSetup.bValueSoundMusic > 100 Then ClientSetup.bValueSoundMusic = 100

            Audio.MusicVolume = ClientSetup.bValueSoundMusic
            lblSound(Index).Caption = Audio.MusicVolume
        Case 3 ' Interface
            ClientSetup.bValueSoundInterface = ClientSetup.bValueSoundInterface + 10
            If ClientSetup.bValueSoundInterface > 100 Then ClientSetup.bValueSoundInterface = 100

            Audio.InterfaceVolume = ClientSetup.bValueSoundInterface
            lblSound(Index).Caption = Audio.InterfaceVolume
    End Select
    
    FrmMain.SetFocus
End Sub

Private Sub imgMenos_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0 ' Master
            If ClientSetup.bValueSoundMaster < 10 Then
                ClientSetup.bValueSoundMaster = 0
            Else
                ClientSetup.bValueSoundMaster = ClientSetup.bValueSoundMaster - 10
            End If
            
            Audio.MasterVolume = ClientSetup.bValueSoundMaster
            lblSound(Index).Caption = Audio.MasterVolume
            
        Case 1 ' Efecto
            If ClientSetup.bValueSoundEffect < 10 Then
                ClientSetup.bValueSoundEffect = 0
            Else
                ClientSetup.bValueSoundEffect = ClientSetup.bValueSoundEffect - 10
            End If

            Audio.EffectVolume = ClientSetup.bValueSoundEffect
            lblSound(Index).Caption = Audio.EffectVolume
            
        Case 2 ' Musica
            If ClientSetup.bValueSoundMusic < 10 Then
                ClientSetup.bValueSoundMusic = 0
            Else
                ClientSetup.bValueSoundMusic = ClientSetup.bValueSoundMusic - 10
            End If

            Audio.MusicVolume = ClientSetup.bValueSoundMusic
            lblSound(Index).Caption = Audio.MusicVolume
        Case 3 ' Interface
            If ClientSetup.bValueSoundInterface < 10 Then
                ClientSetup.bValueSoundInterface = 0
            Else
                ClientSetup.bValueSoundInterface = ClientSetup.bValueSoundInterface - 10
            End If
            
            Audio.InterfaceVolume = ClientSetup.bValueSoundInterface
            lblSound(Index).Caption = Audio.InterfaceVolume
    End Select
    
    FrmMain.SetFocus
End Sub

Private Sub imgAlpha_Click(Index As Integer)
    Select Case Index
    
        Case 0
            If ClientSetup.bAlpha = 100 Then
                ClientSetup.bAlpha = 255
            ElseIf ClientSetup.bAlpha = 150 Then
                ClientSetup.bAlpha = 100
            ElseIf ClientSetup.bAlpha = 200 Then
                ClientSetup.bAlpha = 150
            ElseIf ClientSetup.bAlpha = 255 Then
                ClientSetup.bAlpha = 200
            End If
        Case 1
            If ClientSetup.bAlpha = 100 Then
                ClientSetup.bAlpha = 150
            ElseIf ClientSetup.bAlpha = 150 Then
                ClientSetup.bAlpha = 200
            ElseIf ClientSetup.bAlpha = 200 Then
                ClientSetup.bAlpha = 255
            ElseIf ClientSetup.bAlpha = 255 Then
                ClientSetup.bAlpha = 100
            End If
    End Select
    
    lblAlpha.Caption = CStr(ClientSetup.bAlpha)
    
    FrmMain.SetFocus
End Sub

Private Sub imgChkConfig_Click(Index As Integer)
    
    If Loading Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
        
    If ClientSetup.bConfig(Index) = 0 Then
        ClientSetup.bConfig(Index) = 1
        imgChkConfig(Index).Picture = picCheckBox
    Else
        ClientSetup.bConfig(Index) = 0
        Set imgChkConfig(Index).Picture = picCheckBoxNulo

    End If
    
    ' Cambio de Interfaz
    
    
    #If ModoBig = 0 Then
            If Index = eSetupMods.SETUP_INTERFAZMODERNA Then
                If ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZMODERNA) = 1 Then
                    ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZTDS) = 0
                    Set imgChkConfig(eSetupMods.SETUP_INTERFAZTDS).Picture = picCheckBoxNulo
                        
                    ' LOAD MODERNA
                    
                    FrmMain.Picture = LoadPicture(DirInterface & "main\VentanaClassic2.JPG")
                    'AdaptateControlsToInterface (0)
                    FrmMain.Label4_Click
                    
                Else
                    ClientSetup.bConfig(Index) = 1
                     imgChkConfig(Index).Picture = picCheckBox
                End If
        
            ElseIf Index = eSetupMods.SETUP_INTERFAZTDS Then
        
                If ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZTDS) = 1 Then
                    ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZMODERNA) = 0
                    Set imgChkConfig(eSetupMods.SETUP_INTERFAZMODERNA).Picture = picCheckBoxNulo
                    
                    ' LOAD TDS
                    FrmMain.Picture = LoadPicture(DirInterface & "main\VentanaClassic.JPG")
                    'AdaptateControlsToInterface (0)
                    FrmMain.Label4_Click
                Else
                    ClientSetup.bConfig(Index) = 1
                    imgChkConfig(Index).Picture = picCheckBox
                End If
        
            End If
            
    
    #End If

    If Index = eSetupMods.SETUP_CURSORES Then
        If ClientSetup.bConfig(eSetupMods.SETUP_CURSORES) = 1 Then
            Call mCursor.Cursores_ResotreDefault
        Else

            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorGeneral, IDC_ARROW)
            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorSpell, IDC_CROSS)
            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorHand, IDC_HAND)

        End If
        
    ElseIf Index = eSetupMods.SETUP_PANTALLACOMPLETA Then

        If ClientSetup.bConfig(Index) = 0 Then
            Call Resolution.ResetResolution
        Else
            Call Resolution.SetResolution

        End If

    ElseIf Index = eSetupMods.SETUP_SOUND3D Then
        Audio.Effect3D = ClientSetup.bConfig(Index)
        
        If Not Audio.Effect3D Then
            Call ShowConsoleMsg("¡Desactivaste el sonido 3D y producto de esto no escucharás la ambientación de los mapas! Es una lástima que prefieras no adentrarte a la fantasía desterium.")
        End If
    ElseIf Index = eSetupMods.SETUP_MASTERSOUND Then

        If ClientSetup.bConfig(eSetupMods.SETUP_MASTERSOUND) = 0 Then
            Audio.MasterVolume = 0
        Else
            Audio.MasterVolume = ClientSetup.bValueSoundMaster
            lblSound(0).Caption = Audio.MasterVolume
        End If

    End If
    
    FrmMain.SetFocus

End Sub


Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub lblCursor_Click(Index As Integer)

    CommonDialog1.InitDir = App.path & "\resource\cursor"
    CommonDialog1.Filter = "Imagenes (*.cur|"
    '"Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
    CommonDialog1.ShowOpen
   
    '   imgCursor(Index).Picture = LoadPicture(CommonDialog1.FileName)
   
    Dim Temp As String
    Temp = CommonDialog1.FileName
    Temp = Replace(Temp, CommonDialog1.InitDir, vbNullString)
    lblCursor(Index).Caption = Temp

    Select Case Index
    
        Case 0 ' Cursor General
            ClientSetup.CursorGeneral = Temp
            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorGeneral, IDC_ARROW)
        Case 1 ' Cursor de Lanzar Flechas/Hechizos
            ClientSetup.CursorSpell = Temp
            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorSpell, IDC_CROSS)
        Case 2 ' Cursor del Inventario
            ClientSetup.CursorInv = Temp
        Case 3 ' Cursor de Hand
            ClientSetup.CursorHand = Temp
            Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorHand, IDC_HAND)
    End Select
    
    
    If lblCursor(Index).Caption = vbNullString Then
        lblCursor(Index).Caption = "Seleccionar"
    End If
    
    FrmMain.SetFocus
End Sub


Private Sub imgChkMusica_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Loading Then Exit Sub

    Audio.MusicActivated = Not Audio.MusicActivated
                  
    If Not Audio.MusicActivated Then
        Set imgChkMusica.Picture = picCheckBoxNulo
    Else
        imgChkMusica.Picture = picCheckBox
    End If
    
    lblSound(2).Caption = Audio.MusicVolume

    ClientSetup.bSoundMusic = IIf(Audio.MusicActivated = True, 1, 0)
    
    FrmMain.SetFocus
End Sub

Private Sub imgChkSonidos_Click()
    
    If Loading Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
  
    Audio.EffectActivated = Not Audio.EffectActivated

    If Not Audio.EffectActivated Then
        RainBufferIndex = 0
        FrmMain.IsPlaying = PlayLoop.plNone
              
        Set imgChkSonidos.Picture = picCheckBoxNulo
    Else
              
        imgChkSonidos.Picture = picCheckBox
    End If
    
    lblSound(1).Caption = Audio.EffectVolume

    ClientSetup.bSoundEffect = IIf(Audio.EffectActivated = True, 1, 0)
    
    FrmMain.SetFocus
End Sub

Private Sub imgChkInterfaces_Click()
    If Loading Then Exit Sub
          
    Call Audio.PlayInterface(SND_CLICK)
          
    Audio.InterfaceActivated = Not Audio.InterfaceActivated

    If Not Audio.InterfaceActivated Then
        Set imgChkInterfaces.Picture = picCheckBoxNulo
    Else
        imgChkInterfaces.Picture = picCheckBox
    End If


    lblSound(3).Caption = Audio.InterfaceVolume
    ClientSetup.bSoundInterface = IIf(Audio.InterfaceActivated = True, 1, 0)
    
    FrmMain.SetFocus
End Sub

Private Sub Form_Load()


    #If ModoBig = 0 Then
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    #End If
    
          
    Me.Picture = LoadPicture(App.path & "\resource\interface\options\options.jpg")
    
    Hover_Disabled
    
    LoadButtons
    Loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    Loading = False     'Enable sounds when setting check's values
    
    
    ' Tooltips
    imgFpsInfo.ToolTipText = "Frames por Segundo."
    imgAlphaInfo.ToolTipText = "Transparencia de los efectos como Apocalipsis & Meditaciones."
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub LoadButtons()

    Dim GrhPath As String
    
    GrhPath = DirInterface

    Set LastButtonPressed = New clsGraphicalButton
                        
    Set picCheckBox = LoadPicture(DirInterface & "options\CheckBoxOpciones.jpg")
    Set picCheckBoxNulo = LoadPicture(DirInterface & "options\CheckBoxOpcionesNulo.jpg")
    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonPersonalizados = New clsGraphicalButton
    Set cBotonColores = New clsGraphicalButton
    Set cBotonConfigTeclas = New clsGraphicalButton
    
    'Call cBotonCerrar.Initialize(imgSalir, vbNullString, GrhPath & "generic\BotonCerrarActivo.jpg", vbNullString, Me)
    'Call cBotonPersonalizados.Initialize(imgPersonalizados, vbNullString, GrhPath & "options\BotonPersonalizadosActivo.jpg", vbNullString, Me)
    'Call cBotonColores.Initialize(imgColores, vbNullString, GrhPath & "options\BotonColoresDialogoActivo.jpg", vbNullString, Me)
    'Call cBotonConfigTeclas.Initialize(imgConfigTeclas, vbNullString, GrhPath & "options\BotonConfigTeclasActivo.jpg", vbNullString, Me)

End Sub

Private Sub LoadUserConfig()

    ' Load music config
    If Audio.MusicActivated Then
        imgChkMusica.Picture = picCheckBox
        
        
    End If
    
    lblSound(0).Caption = Audio.MasterVolume
    lblSound(1).Caption = Audio.EffectVolume
    lblSound(2).Caption = Audio.MusicVolume
    lblSound(3).Caption = Audio.InterfaceVolume
    
    ' Load Sound config
    If Audio.EffectActivated Then
        imgChkSonidos.Picture = picCheckBox
    End If
    
    If Audio.InterfaceActivated Then
        imgChkInterfaces.Picture = picCheckBox
    End If
    
    Dim A As Long
        
    For A = 1 To MAX_SETUP_MODS
        If ClientSetup.bConfig(A) = 1 Then
            imgChkConfig(A).Picture = picCheckBox
        Else
            Set imgChkConfig(A).Picture = picCheckBoxNulo
        End If
    Next A
    
    lblAlpha.Caption = CStr(ClientSetup.bAlpha)
    
    lblFPS.Caption = ClientSetup.bAlpha

    Fps_List_Disabled Fps_Detect_Index
    Alpha_List_Disabled Alpha_Detect_Index

    ' Cursores Predeterminados
    Dim filePath As String
    filePath = App.path & "\resource\cursor"

    lblCursor(0).Caption = Trim$(ClientSetup.CursorGeneral)
    lblCursor(1).Caption = Trim$(ClientSetup.CursorSpell)
    lblCursor(2).Caption = Trim$(ClientSetup.CursorInv)
    lblCursor(3).Caption = Trim$(ClientSetup.CursorHand)
    
    
    For A = lblCursor.LBound To lblCursor.UBound
        If lblCursor(A).Caption = vbNullString Then
             lblCursor(A).Caption = "Seleccionar"
        End If
    Next A
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
    
        
    Dim A As Long
    
    For A = 0 To 3
        imgMenos(A).Picture = Nothing
        imgMas(A).Picture = Nothing
    Next A
    
    Hover_Disabled
    
End Sub

Private Sub Hover_Disabled()

    Dim A As Long
    
    For A = lblMenu.LBound To lblMenu.UBound
        lblMenu(A).ForeColor = RGB(230, 230, 230)
    Next A
    
End Sub

Private Sub Hover_Activate(ByVal Index As Integer)
    Hover_Disabled
    lblMenu(Index).ForeColor = RGB(255, 116, 0)
End Sub
Private Sub lblMenu_Click(Index As Integer)

    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0 ' Configurar Teclas
            #If ModoBig = 1 Then
                dockForm frmCustomKeys.hWnd, FrmMain.PicMenu, True
            #Else
                Call frmCustomKeys.Show(, Me)
            #End If
            
        Case 1
            Call frmMessageTxt.Show(, Me)
        
        Case 2
            Call frmDialogos.Show(, Me)
            
        Case 3
            ' Manual del juego
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/wiki/doku.php?id=Bienvenida", vbNullString, vbNullString, 1)
    End Select
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Hover_Activate Index
End Sub

Private Sub Fps_List_Disabled(ByVal Index As Integer)

    Dim A As Long
    
    For A = PicFPS.LBound To PicFPS.UBound
         Set PicFPS(A).Picture = picCheckBoxNulo
    Next A
    
    PicFPS(Index).Picture = picCheckBox
End Sub
Private Function Fps_Detect_Index() As Integer
    Select Case ClientSetup.bFps
        Case 144 ' 144
            Fps_Detect_Index = 0
        Case 244 ' 244
            Fps_Detect_Index = 1
        Case 2 ' Vsync
            Fps_Detect_Index = 2
        Case 1 ' Libres
            Fps_Detect_Index = 3
    End Select

End Function
Private Sub Alpha_List_Disabled(ByVal Index As Integer)

    Dim A As Long
    
    For A = PicAlpha.LBound To PicAlpha.UBound
         Set PicAlpha(A).Picture = picCheckBoxNulo
    Next A
    
    PicAlpha(Index).Picture = picCheckBox
End Sub

Private Sub PicAlpha_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0 '
            ClientSetup.bAlpha = 100
        Case 1 '
            ClientSetup.bAlpha = 150
        Case 2 '
            ClientSetup.bAlpha = 200
        Case 3 '
            ClientSetup.bAlpha = 255
    End Select
    
    Alpha_List_Disabled Index

End Sub
Private Function Alpha_Detect_Index() As Integer
    Select Case ClientSetup.bAlpha
        Case 100 '
            Alpha_Detect_Index = 0
        Case 150 '
            Alpha_Detect_Index = 1
        Case 200 '
            Alpha_Detect_Index = 2
        Case 255 '
            Alpha_Detect_Index = 3
    End Select

End Function


Private Sub PicFPS_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0 ' 144
            ClientSetup.bFps = 144
        Case 1 ' 244
            ClientSetup.bFps = 244
        Case 2 ' Vsync
            ClientSetup.bFps = 2
        Case 3 ' Libres
            ClientSetup.bFps = 1
    End Select
    
    
    Fps_List_Disabled Index
    
    
    If Index = 2 Then
        Call MsgBox("Debes reiniciar para poder apreciar los cambios. Te recordamos que esto depende de tu monitor y de los Ghz que tenga")
    End If
    
    
End Sub
