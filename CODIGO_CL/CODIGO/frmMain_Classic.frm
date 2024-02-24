VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain_Classic 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain_Classic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox PicMenu 
      BorderStyle     =   0  'None
      Height          =   6240
      Left            =   105
      Picture         =   "frmMain_Classic.frx":000C
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   82
      Top             =   2250
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   14
         Left            =   120
         Top             =   5670
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   13
         Left            =   120
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   12
         Left            =   120
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   11
         Left            =   120
         Top             =   4500
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   10
         Left            =   120
         Top             =   4155
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   9
         Left            =   120
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   8
         Left            =   120
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   7
         Left            =   120
         Top             =   2970
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   6
         Left            =   120
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   5
         Left            =   120
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   4
         Left            =   120
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   3
         Left            =   120
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   2
         Left            =   120
         Top             =   1040
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   1
         Left            =   120
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image imgMiniMenu 
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox PicPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   12120
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1770
   End
   Begin VB.PictureBox PicStats 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   3045
      Left            =   4320
      Picture         =   "frmMain_Classic.frx":C1B6
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   182
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   2730
      Begin VB.Image imgPMSG 
         Height          =   300
         Left            =   1005
         Top             =   1890
         Width           =   315
      End
      Begin VB.Image CMSG 
         Height          =   300
         Left            =   675
         Top             =   1890
         Width           =   315
      End
      Begin VB.Image imgEvents 
         Height          =   375
         Left            =   210
         Top             =   1290
         Width           =   1110
      End
      Begin VB.Image imgClanes 
         Height          =   375
         Left            =   210
         Top             =   915
         Width           =   1110
      End
      Begin VB.Image imgOpciones 
         Height          =   375
         Left            =   1410
         Top             =   1290
         Width           =   1110
      End
      Begin VB.Image imgGoStats 
         Height          =   330
         Left            =   105
         Top             =   60
         Width           =   1275
      End
      Begin VB.Image imgSeg 
         Height          =   300
         Left            =   2295
         Top             =   1890
         Width           =   315
      End
      Begin VB.Image imgDrag 
         Height          =   300
         Left            =   2265
         Top             =   2280
         Width           =   315
      End
      Begin VB.Image imgResu 
         Height          =   300
         Left            =   1050
         Top             =   2265
         Width           =   315
      End
      Begin VB.Image imgFight 
         Height          =   375
         Left            =   1380
         Top             =   540
         Width           =   1155
      End
      Begin VB.Image imgParty 
         Height          =   375
         Left            =   1380
         Top             =   915
         Width           =   1155
      End
      Begin VB.Image imgSocial 
         Height          =   225
         Index           =   0
         Left            =   840
         Top             =   2625
         Width           =   225
      End
      Begin VB.Image imgSocial 
         Height          =   225
         Index           =   1
         Left            =   1125
         Top             =   2625
         Width           =   225
      End
      Begin VB.Image imgSocial 
         Height          =   225
         Index           =   2
         Left            =   1395
         Top             =   2625
         Width           =   225
      End
      Begin VB.Image imgSocial 
         Height          =   225
         Index           =   3
         Left            =   1680
         Top             =   2580
         Width           =   225
      End
   End
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H8000000D&
      Height          =   6240
      Left            =   105
      MousePointer    =   99  'Custom
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   8
      Top             =   2250
      Width           =   8160
      Begin VB.Timer tEngine 
         Interval        =   1000
         Left            =   1800
         Top             =   3120
      End
      Begin VB.Timer dobleclick 
         Left            =   960
         Top             =   2400
      End
      Begin VB.Timer tAnuncios 
         Interval        =   1
         Left            =   5520
         Top             =   960
      End
      Begin VB.Timer tUpdateMS 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   1320
      End
      Begin VB.Timer MacroTrabajo 
         Enabled         =   0   'False
         Left            =   3990
         Top             =   1785
      End
      Begin VB.Timer UpdateMapa 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   3255
         Top             =   1680
      End
      Begin VB.Timer tMapData 
         Enabled         =   0   'False
         Left            =   2205
         Top             =   1560
      End
      Begin VB.Timer tMapName 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2205
         Top             =   315
      End
      Begin VB.Timer tMessage 
         Interval        =   60000
         Left            =   1560
         Top             =   720
      End
      Begin VB.Timer tmrBlink 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   840
         Top             =   240
      End
      Begin VB.Timer tUpdateInactive 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   7440
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   40
         Left            =   4560
         Top             =   840
      End
   End
   Begin VB.PictureBox picHechiz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2880
      Left            =   8640
      MousePointer    =   99  'Custom
      Picture         =   "frmMain_Classic.frx":16C92
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2250
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.PictureBox picHabla 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   6960
      ScaleHeight     =   1125
      ScaleWidth      =   1335
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Label lblHabla 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emojis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   32
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblHabla 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Susurro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   210
         Index           =   3
         Left            =   270
         TabIndex        =   30
         Top             =   570
         Width           =   720
      End
      Begin VB.Label lblHabla 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   210
         Index           =   2
         Left            =   330
         TabIndex        =   29
         Top             =   330
         Width           =   630
      End
      Begin VB.Label lblHabla 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   1
         Left            =   330
         TabIndex        =   28
         Top             =   90
         Width           =   630
      End
   End
   Begin VB.PictureBox MiniMapa 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1470
      Left            =   6780
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   420
      Width           =   1470
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8130
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      Height          =   195
      Left            =   8640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   180
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   7800
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   7320
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   4
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   6840
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   3
      Top             =   9240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   5880
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   2
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   5190
      Top             =   2520
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3990
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1380
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2434
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain_Classic.frx":1BE7D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   8760
      MousePointer    =   4  'Icon
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   24
      Top             =   2250
      Width           =   2625
   End
   Begin VB.Image imgDropGold 
      Height          =   255
      Left            =   10560
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Left            =   240
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label lblHome 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5040
      TabIndex        =   80
      Top             =   8520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgHome 
      Height          =   480
      Left            =   4920
      Picture         =   "frmMain_Classic.frx":1BEFB
      Top             =   8520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMagic 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1500
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblAnillo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50/50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image lblRedes 
      Height          =   495
      Index           =   3
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image lblRedes 
      Height          =   495
      Index           =   2
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image lblRedes 
      Height          =   495
      Index           =   1
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image lblRedes 
      Height          =   495
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDopa 
      Height          =   330
      Left            =   5760
      Picture         =   "frmMain_Classic.frx":1C730
      Top             =   8580
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblDopa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6150
      TabIndex        =   77
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   345
      Index           =   1
      Left            =   14040
      TabIndex        =   76
      Top             =   1440
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label lblFuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   345
      Index           =   1
      Left            =   13560
      TabIndex        =   75
      Top             =   4110
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAgilidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   15840
      TabIndex        =   74
      Top             =   4080
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblOns 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   14880
      TabIndex        =   73
      Top             =   2040
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   14880
      TabIndex        =   72
      Top             =   2520
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblFPS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   14880
      TabIndex        =   71
      Top             =   2280
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   13680
      TabIndex        =   70
      Top             =   3000
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "47%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   14040
      TabIndex        =   69
      Top             =   3480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Index           =   5
      Left            =   14865
      TabIndex        =   68
      Top             =   4800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   5
      Left            =   14760
      TabIndex        =   67
      Top             =   5160
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   5
      Left            =   14910
      TabIndex        =   66
      Top             =   5640
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Index           =   5
      Left            =   15015
      TabIndex        =   65
      Top             =   6240
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   5
      Left            =   15135
      TabIndex        =   64
      Top             =   6720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   1
      Left            =   13200
      TabIndex        =   63
      Top             =   840
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.Label lblParalisis 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   3240
      TabIndex        =   62
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblInvi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   4320
      TabIndex        =   61
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInvisible 
      Height          =   330
      Left            =   3960
      Picture         =   "frmMain_Classic.frx":1FA8B
      Top             =   8595
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgParalisis 
      Height          =   330
      Left            =   3480
      Picture         =   "frmMain_Classic.frx":20C6E
      Top             =   8580
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblAgilidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   0
      Left            =   7770
      TabIndex        =   60
      Top             =   8610
      Width           =   315
   End
   Begin VB.Label lblFuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   6915
      TabIndex        =   59
      Top             =   8610
      Width           =   315
   End
   Begin VB.Image imgExp 
      Height          =   180
      Left            =   8895
      Picture         =   "frmMain_Classic.frx":21F40
      Top             =   975
      Width           =   1830
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   8580
      TabIndex        =   14
      Top             =   8055
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   8595
      TabIndex        =   58
      Top             =   8055
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   8565
      TabIndex        =   57
      Top             =   8055
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   8580
      TabIndex        =   56
      Top             =   8070
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   8580
      TabIndex        =   55
      Top             =   8040
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8580
      TabIndex        =   13
      Top             =   7695
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8565
      TabIndex        =   54
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8595
      TabIndex        =   53
      Top             =   7695
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8580
      TabIndex        =   52
      Top             =   7710
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8580
      TabIndex        =   51
      Top             =   7680
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8580
      TabIndex        =   12
      Top             =   7335
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8565
      TabIndex        =   50
      Top             =   7350
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8595
      TabIndex        =   49
      Top             =   7335
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8580
      TabIndex        =   48
      Top             =   7350
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8580
      TabIndex        =   47
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8580
      TabIndex        =   11
      Top             =   6960
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8565
      TabIndex        =   46
      Top             =   6960
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8565
      TabIndex        =   45
      Top             =   6960
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8580
      TabIndex        =   44
      Top             =   6945
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8580
      TabIndex        =   43
      Top             =   6975
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   8580
      TabIndex        =   10
      Top             =   6600
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   4
      Left            =   8580
      TabIndex        =   42
      Top             =   6615
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   3
      Left            =   8580
      TabIndex        =   41
      Top             =   6600
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   8580
      TabIndex        =   40
      Top             =   6585
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   1
      Left            =   8565
      TabIndex        =   39
      Top             =   6600
      Width           =   1350
   End
   Begin VB.Image imgButton 
      Height          =   405
      Index           =   8
      Left            =   10800
      Top             =   5490
      Width           =   405
   End
   Begin VB.Image imgButton 
      Height          =   405
      Index           =   7
      Left            =   9915
      Top             =   5490
      Width           =   405
   End
   Begin VB.Image imgButton 
      Height          =   405
      Index           =   6
      Left            =   8985
      Top             =   5490
      Width           =   405
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   5
      Left            =   12480
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   4
      Left            =   13200
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image imgButton 
      Height          =   210
      Index           =   3
      Left            =   10620
      Top             =   6975
      Width           =   645
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   0
      Left            =   10320
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   1
      Left            =   10320
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   210
      Index           =   2
      Left            =   10620
      Top             =   6645
      Width           =   645
   End
   Begin VB.Image imgInfo 
      Height          =   375
      Left            =   10440
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   10440
      TabIndex        =   37
      Top             =   990
      Width           =   1215
   End
   Begin VB.Image imgMoveSpell 
      Height          =   375
      Index           =   1
      Left            =   11160
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgMoveSpell 
      Height          =   255
      Index           =   0
      Left            =   11160
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad de Ullathorpe"
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Index           =   0
      Left            =   9000
      TabIndex        =   35
      Top             =   8655
      Width           =   2355
   End
   Begin VB.Image CmdLanzar 
      Height          =   645
      Left            =   8520
      MouseIcon       =   "frmMain_Classic.frx":234AA
      MousePointer    =   99  'Custom
      Top             =   5160
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Left            =   8640
      Top             =   120
      Width           =   1290
   End
   Begin VB.Image imgCerrar 
      Height          =   315
      Left            =   10320
      Top             =   120
      Width           =   1170
   End
   Begin VB.Image imgStats 
      Height          =   255
      Left            =   10320
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lblEldhir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   11640
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblHabla 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   390
      Index           =   0
      Left            =   10710
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[BRONCE]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   0
      Left            =   8625
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   90
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   11925
      TabIndex        =   22
      Top             =   210
      Width           =   135
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblhelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   " "
      Top             =   8760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anstirion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   0
      Left            =   8700
      TabIndex        =   16
      Top             =   675
      Width           =   2175
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Index           =   0
      Left            =   10800
      TabIndex        =   15
      Top             =   6420
      Width           =   765
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13200
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   180
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13470
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   180
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10080
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   13320
      Top             =   120
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   13365
      Top             =   120
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8400
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   1560
      Width           =   1605
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   6240
      Left            =   120
      Top             =   2235
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Label lblarmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   8520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image STAShp 
      Height          =   165
      Left            =   8550
      Picture         =   "frmMain_Classic.frx":235FC
      Top             =   6615
      Width           =   1455
   End
   Begin VB.Image Hpshp 
      Height          =   180
      Left            =   8550
      MouseIcon       =   "frmMain_Classic.frx":2459D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain_Classic.frx":246EF
      Top             =   7350
      Width           =   1455
   End
   Begin VB.Image COMIDAsp 
      Height          =   195
      Left            =   8550
      Picture         =   "frmMain_Classic.frx":25683
      Top             =   7695
      Width           =   1455
   End
   Begin VB.Image AGUAsp 
      Height          =   180
      Left            =   8550
      Picture         =   "frmMain_Classic.frx":26613
      Top             =   8070
      Width           =   1455
   End
   Begin VB.Image InvEqu 
      Height          =   4890
      Left            =   8325
      Top             =   1425
      Width           =   3495
   End
   Begin VB.Image MANShp 
      Height          =   165
      Left            =   8550
      Picture         =   "frmMain_Classic.frx":27598
      Top             =   6975
      Width           =   1455
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain_Classic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MouseData As clsMouse
Attribute MouseData.VB_VarHelpID = -1

Private ZoomIn As Boolean

Private SpellSelected As Byte
Public CoordBloqued As Boolean
Public PorcBloqued As Boolean
Public CursorSelected As Byte



' Detectar posicion del cursor.
Private Declare Function GetCursorPos Lib "user32.dll" (Pt As Point) As Long

Private totalclicks As Integer

Public ModoTab As Boolean

Private Type Point

    X As Long
    Y As Long

End Type

'End Security
        
' x Auto Pots
Private Enum eVentanas

    vHechizos = 1
    vInventario = 2

End Enum

Public Panel                   As Byte

Private LastPanel               As Byte

Private Const InvalidSlot       As Byte = 255

' x Auto Pots

' x button
Private mouse_Down              As Boolean

Private mouse_UP                As Boolean
' x button

Public N                        As Byte

Private MouseInvBoton           As Long

Public Attack                   As Boolean

Private Last_I                  As Long

Public WithEvents dragInventory As clsGrapchicalInventory
Attribute dragInventory.VB_VarHelpID = -1

Dim Ancho                       As Integer

Dim alto                        As Integer

Public tX                       As Byte

Public tY                       As Byte

Public MouseX                   As Long



Public MouseY                   As Long

Public MouseBoton               As Long

Public MouseShift               As Long

Private clicX                   As Long

Private clicY                   As Long

Public IsPlaying                As Byte

Private clsFormulario           As clsFormMovementManager

Public picSkillStar             As Picture

Private bCMSG                   As Boolean

Private PMSGimg                 As Boolean

Private btmpCMSG                As Boolean

Private sPartyChat              As String

Private bLastBrightBlink        As Boolean

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&


' Botones Gráficos
Private cBotonOpciones     As clsGraphicalButton
Private cBotonParty        As clsGraphicalButton
Private cBotonRetos        As clsGraphicalButton
Private cBotonEventos      As clsGraphicalButton
Private cBotonClanes       As clsGraphicalButton
Private cBotonObjetive     As clsGraphicalButton
Private cBotonCerrar       As clsGraphicalButton
Private cBotonMinimizar    As clsGraphicalButton
Private cBotonLanzar       As clsGraphicalButton
Private cBotonSkills       As clsGraphicalButton
Public LastButtonPressed   As clsGraphicalButton

Private Sub CMSG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
    SetHand
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableURLDetect
    
    'MouseData.Hook
   ' hSwapCursor = SetClassLong(frmMain.hWnd, GLC_HCURSOR, hSwapCursor)
End Sub

Private Sub dobleclick_Timer()

    Static segundo As Long

    segundo = segundo + 1

    If segundo = 2 And totalclicks > 20 Then
        Call WriteDenounce("[SEGURIDAD]: Posible uso de Doble-Clic: " & totalclicks)
        totalclicks = 0
        segundo = 0
        dobleclick.Interval = 0
    End If

    If segundo = 2 And totalclicks <= 20 Then
        totalclicks = 0
        segundo = 0
        dobleclick.Interval = 0

    End If

End Sub

'Private Sub hlst_Click()
   ' If (MouseShift And 1) = 1 Then
      '  If hlst.ListIndex <> -1 Then
      '      Call WriteSpellInfo(hlst.ListIndex + 1)
      '  End If
   ' End If
'End Sub

'Private Sub hlst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' MouseShift = Shift
'End Sub

Private Sub Hpshp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
    SetHand
End Sub

Private Sub imgButton_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0
            Call frmOpciones.Show(, FrmMain)

        Case 1
            Call WriteGuilds_Required(0)

        Case 2
           CMSG_Click
        Case 3
            imgPMSG_Click
        Case 6
            Call WriteSafeToggle
        Case 7
            Call WriteDragToggle
        Case 8
            Call WriteResuscitationToggle
    End Select

End Sub

Private Sub imgCerrar_Click()
    Call Audio.PlayInterface(SND_CLICK)

    If MsgBox("¿Estás seguro que deseas salir del personaje?", vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
        'prgRun = False
        Call ParseUserCommand("/SALIR")
    End If
End Sub

Private Sub imgClanes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHand
End Sub

Private Sub imgDrag_Click()
    Call WriteDragToggle
End Sub

Private Sub imgDropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then FrmCantidad.Show , FrmMain
    End If
End Sub

Private Sub imgEvents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetHand
End Sub

Private Sub imgMapa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetHand
End Sub

Private Sub imgFight_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ShowConsoleMsg("Ayuda» Los comandos /RETOSON y /RETOSOFF activan un Panel que te ayudará a ver la invitación en una nueva Ventana.", 150, 200, 148, True)
    Call ParseUserCommand("/RETOS")
End Sub

Private Sub imgGoMenu_Click()
     Call Audio.PlayInterface(SND_CLICK)
     PicStats.visible = True
End Sub

Private Sub imgGoStats_Click()
     Call Audio.PlayInterface(SND_CLICK)
    PicStats.visible = False
End Sub

Private Sub ImgInfo_Click()

    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)

    End If

End Sub

Private Sub imgMenu_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If PicMenu.visible Then
        PicMenu.visible = False
    Else
        PicMenu.visible = True
    End If
End Sub

Private Sub imgMiniMenu_Click(Index As Integer)

    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
        Case 0 ' SHOP
            Call ParseUserCommand("/SHOP")
        Case 1 ' Skins
            Call ParseUserCommand("/SKINS")
        Case 2 ' Mercado
            Call WriteMercader_Required(1, 1, 255)
        Case 3 ' Misiones
            Call ParseUserCommand("/MISIONES")
        Case 4 ' TOP Nivel
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/level/", vbNullString, vbNullString, 1)
        Case 5 ' TOP Retos
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/retos/", vbNullString, vbNullString, 1)
        Case 6 ' TOP Frags
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/aniquiladores/", vbNullString, vbNullString, 1)
        Case 7 ' TOP Oro
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/millonarios/", vbNullString, vbNullString, 1)
        Case 8 ' Partidas
            Call ParseUserCommand("/TORNEOS")
        Case 9 ' Discord
            Call ShellExecute(hWnd, "open", "https://discord.com/servers/desterium-game-700452646326239252", vbNullString, vbNullString, 1)
        Case 10 ' Web
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/", vbNullString, vbNullString, 1)
        Case 11 ' Menu ESC
             #If ModoBig = 1 Then
                dockForm FrmMenu.hWnd, FrmMain.PicMenu, True
            #Else
                FrmMain.SetFocus
            #End If
        Case 12 ' Manual
            Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/wiki", vbNullString, vbNullString, 1)
        Case 13 ' Streaming
            Call ShellExecute(hWnd, "open", "https://forms.gle/PWiYqfWdGbGN6vY89", vbNullString, vbNullString, 1)
        Case 14 ' Staff
            Call ShellExecute(hWnd, "open", "https://forms.gle/rS9jrkXfMYjmdaiy9", vbNullString, vbNullString, 1)
    End Select
    
    PicMenu.visible = False
End Sub

Private Sub imgMinimize_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Me.WindowState = 1
End Sub

Private Sub imgMoveSpell_Click(Index As Integer)
      
    Call Audio.PlayInterface(SND_CLICK)
        
    If hlst.ListIndex = -1 Then Exit Sub
        
    Dim sTemp As String

    Dim Temp  As Integer
        
    Select Case Index

        Case 0 'subir

            If hlst.ListIndex = 0 Then Exit Sub

        Case 1 'bajar

            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

    End Select
    
    Temp = UserHechizos(hlst.ListIndex + 1)
        
    Select Case Index

        Case 0 'subir
            Call WriteMoveSpell(hlst.ListIndex, hlst.ListIndex + 1)
            sTemp = hlst.List(hlst.ListIndex - 1)

             
            'UserHechizos(hlst.ListIndex) = UserHechizos(hlst.ListIndex + 1)
            UserHechizos(hlst.ListIndex) = Temp
                
            
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1
            hlst.Scroll = hlst.Scroll - 1
                
        Case 1 'bajar
            Call WriteMoveSpell(hlst.ListIndex + 1, hlst.ListIndex + 2)
            sTemp = hlst.List(hlst.ListIndex + 1)
           
            
            'UserHechizos(hlst.ListIndex) = UserHechizos(hlst.ListIndex + 1)
            UserHechizos(hlst.ListIndex + 2) = Temp
                
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1
            hlst.Scroll = hlst.Scroll + 1
            
    End Select
        
    hlst.DownBarrita = 0

End Sub

Private Sub imgMoveSpell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgObjetive_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call WriteQuestRequired(0)
End Sub


Private Sub imgRank_Click()
    Call Audio.PlayInterface(SND_CLICK)
     Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/level/", vbNullString, vbNullString, 1)
End Sub

Private Sub imgWeb_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/", vbNullString, vbNullString, 1)
End Sub

Private Sub imgParty_Click()
    Call Audio.PlayInterface(SND_CLICK)
Call ShowConsoleMsg("Ayuda» Es hora de enviar solicitudes para que usuarios formen un grupo contigo.. Haz clic sobre aquel que desees invitar y luego teclea F3.", 150, 200, 148, True)
Call WritePartyClient(1)
End Sub

Private Sub imgResu_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteResuscitationToggle
End Sub

Private Sub imgSeg_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteSafeToggle
End Sub

Private Sub imgSocial_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Url As String
    
    Select Case Index
    
        Case 0 ' Instagram
            Url = "https://www.instagram.com/ArgentumGame"
        Case 1 ' Youtube
             Url = "https://www.instagram.com/ArgentumGame"
        Case 2 ' Facebook
             Url = "https://www.facebook.com/ArgentumGame"
        Case 3 ' Discord
             Url = "https://www.discord.argentumgame.com/"
    End Select
    
    Call ShellExecute(hWnd, "open", Url, vbNullString, vbNullString, 1)
End Sub

Private Sub imgStats_Click()

  If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)

    Call WriteRequestSkills
End Sub



Private Sub SetHand()
    If Not CursorSelected = 3 Then
    Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorHand, IDC_ARROW)
    CursorSelected = 3
    End If
End Sub

Private Sub Label1_Click()
    frmOpciones.Show , FrmMain

End Sub

Private Sub lblMana_Click(Index As Integer)
    MANShp_Click
End Sub

Private Sub lblMap_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)

    If CoordBloqued Then
        CoordBloqued = False
    Else
        CoordBloqued = True

    End If

End Sub

Private Sub lblMap_MouseMove(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    If Not CoordBloqued Then
        Map_UpdateLabel (True)
    End If
End Sub

Private Sub lblporclvl_Click(Index As Integer)

    
        PorcBloqued = Not PorcBloqued
    
    If UserPasarNivel > 0 Then
        Call ShowConsoleMsg(Format$(UserExp, "#,###") & "/" & Format$(UserPasarNivel, "#,###") & " " & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%")
        
    End If
    
    Call Render_Exp(False)
End Sub

Private Sub lblporclvl_MouseMove(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim A As Long
    
    
    #If ModoBig > 0 Then
        If Index <> 2 Then Exit Sub
        
    #End If
    
    SetHand
    
    Call Render_Exp(False)

End Sub

Private Sub lblVida_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetHand
End Sub

Private Sub MANShp_Click()
    If Not MainTimer.Check(TimersIndex.Packet250) Then Exit Sub
    Call ParseUserCommand("/MEDITAR")
    
End Sub

Private Sub lblVida_Click(Index As Integer)
    Hpshp_Click
End Sub

Private Sub Hpshp_Click()
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call ParseUserCommand("/EST")
End Sub



Private Sub MANShp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetHand
End Sub

Private Sub MiniMapa_Click()


    #If ModoScreen = 1 Then
        frmScreenShot.Show , FrmMain
    #Else
        Call Audio.PlayInterface(SND_CLICK)
        Call FrmMapa.Show(, FrmMain)
    #End If
    
End Sub


Private Sub picHechiz_DblClick()
  Dim Temp As String
    
    If (MouseShift And 1) = 1 Then
        If SpellSelected = 0 Then
            SpellSelected = hlst.ListIndex + 1
            hlst.SetForeColor = vbGreen
            'hlst.List(SpellSelected).ForeColor = vbGreen
        Else

            Call WriteMoveSpell(SpellSelected, hlst.ListIndex + 1)
            Temp = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = hlst.List(SpellSelected - 1)
            hlst.List(SpellSelected - 1) = Temp
            
            hlst.SetForeColor = vbWhite
            'hlst.ForeColor = vbWhite
            SpellSelected = 0
        
        End If
    
    End If
End Sub


Private Sub PicStats_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    StartCheckingLinks
End Sub

Private Sub Form_Load()
    
    ' Lista Gráfica
    Set hlst = New clsGraphicalList

    Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))
 
    
    'Drag And Drop
    Set dragInventory = Inventario

    ' Handles Form movement (drag and dr|op).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, 120
          
    Call LoadButtons
    
      If ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZTDS) Then
            FrmMain.Picture = LoadPicture(DirInterface & "main\VentanaClassic.JPG")
        Else
            FrmMain.Picture = LoadPicture(DirInterface & "main\VentanaClassic2.JPG")

        End If
        
    imgExp.Picture = LoadPicture(DirInterface & "main\ExpBar.jpg")
    
    If ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZTDS) = 1 Then
         imgExp.visible = False
    Else
        imgExp.visible = True
    End If
    
    picHechiz.Picture = Nothing

      'Set MouseData = New clsMouse
     ' MouseData.Hook Me
    
    Me.Left = 0
    Me.Top = 0
    Me.Width = 12000
    Me.Height = 9000
    
    EnableURLDetect RecTxt.hWnd, Me.hWnd
    
    Call SetWindowLong(RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = DirInterface

    Set cBotonOpciones = New clsGraphicalButton
    Set cBotonParty = New clsGraphicalButton
    Set cBotonRetos = New clsGraphicalButton
    Set cBotonEventos = New clsGraphicalButton
    Set cBotonClanes = New clsGraphicalButton
    Set cBotonObjetive = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMinimizar = New clsGraphicalButton
    Set cBotonLanzar = New clsGraphicalButton
    Set cBotonSkills = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '***************************************************
  
  '  If KeyCode = 16 Then
          '  If MainTimer.Check(TimersIndex.Shift) Then
             '   Call WriteAcelerationChar
          '  End If
    '    End If
        
    If (Not SendTxt.visible) Then
    
        If esGM(UserCharIndex) Then
            If KeyCode = vbKeyI Then
                Call ParseUserCommand("/INVISIBLE")

            End If

        End If
        
      
                    
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyPanelParty)
                        Call ShowConsoleMsg("Ayuda» Es hora de enviar solicitudes para que usuarios formen un grupo contigo.. Haz clic sobre aquel que desees invitar y luego teclea F3.", 150, 200, 148, True)
                        Call WritePartyClient(1)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTabPanel)
                    FrmMain.ModoTab = True
                    
                    If FrmMain.Panel = 1 Then
                       ' FrmMain.Label4_Click
                    Else
                      '  FrmMain.Label7_Click
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyPanelFight)
                    Call ShowConsoleMsg("Recuerda que podrás utilizar el comando /RETOSON y /RETOSOFF para mostrar la invitación recibida de una forma más segura.", 150, 200, 148, True)
                    Call ParseUserCommand("/RETOS")
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyWork)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                        Exit Sub

                    End If
                  
                    If MacroTrabajo.Enabled Then
                        Call DesactivarMacroTrabajo
                    Else
                        Call ActivarMacroTrabajo

                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMaster)

                    If Audio.MasterVolume = 0 Then
                        Audio.MasterVolume = 100
                    Else
                        Audio.MasterVolume = 0

                    End If
                          
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.EffectActivated = Not Audio.EffectActivated

                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                      
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeModeResu)
                    Call WriteResuscitationToggle
                      
                Case vbKeyZ:

                    If DialogosClanes.Activo = False Then
                        Call ShowConsoleMsg("Consola flotante de clanes activada.", 255, 200, 200)
                        DialogosClanes.Activo = True
                    Else
                        Call ShowConsoleMsg("Consola flotante de clanes desactivada.", 255, 200, 200)
                        DialogosClanes.Activo = False

                    End If
                      
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                      
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDenounce)

                    'Intervalo permite usar este sistema?
                    If Not FotoD_CanSend Then
                        Call AddtoRichTextBox(FrmMain.RecTxt, "Haz alcanzado el máximo de envio de 1 FotoDenuncia por minuto. Esperá unos instantes y volve a intentar.", 0, 200, 200, False, False, True)
        
                        Exit Sub
        
                    End If

                    'Aca guardamos el string que nos devuelve FotoD_Capturar.
                    Dim nString As String
        
                    FotoD_Capturar nString
        
                    'Si el string da nullo, es por que nadie esta insultando.
                    If nString = vbNullString Then
                        Call AddtoRichTextBox(FrmMain.RecTxt, "Nadie te esta insultando. Las FotoDenuncias solo sirven para denunciar agravios.", 0, 200, 200, False, False, True)
                    Else 'Si no, enviamos.
                        Call AddtoRichTextBox(FrmMain.RecTxt, "La FotoDenuncia fue sacada correctamente.", 0, 200, 200, False, False, True)
                        WriteDenounce "[FOTODENUNCIAS] : " & nString

                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyThief)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                    Else
                        Call WriteWork(eSkill.Robar)

                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyHelpGuild)
                    Call WriteGuilds_Talk("Solicito ayuda en " & UserMapName & " (Coord: " & UserMap & " " & UserPos.X & " " & UserPos.Y & ")", True)

                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                    Else
                        Call WriteWork(eSkill.Ocultarse)

                    End If
                                          
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                    Call UsarItem(0)
                      
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep

                    End If
                
            End Select

        Else

            Select Case KeyCode

                    'Custom messages!
                Case vbKey0 To vbKey9

                    Dim CustomMessage As String
                          
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)

                    If LenB(CustomMessage) <> 0 Then

                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase(Left(CustomMessage, 5)) <> "/CMSG" And Left(CustomMessage, 1) <> "\" Then
                                  
                            Call ParseUserCommand(CustomMessage)

                        End If

                    End If

            End Select

        End If
            
    End If
         
    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

            Exit Sub
                  
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture(, True)
              
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            'If UserMinMAN = UserMaxMAN Then Exit Sub
                  
            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If
                      
            Call RequestMeditate
            
        Case CustomKeys.BindedKey(eKeyType.mKeyPartySend)
            Call WritePartyClient(5)
            
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)

            If FrmMain.MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
           
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

            If FrmMain.MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            If Shift <> 0 Then Exit Sub
             
            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                If MainTimer.Check(TimersIndex.Attack) Then
           
                    If TrainingMacro.Enabled Then DesactivarMacroHechizos
                    Call MainTimer.Restart(TimersIndex.AttackSpell)
                    Call MainTimer.Restart(TimersIndex.AttackUse)
                    Call WriteAttack

                End If
            
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            
            If (Not Comerciando) And (Not MirandoForo) And (Not MirandoEstadisticas) And (Not MirandoCantidad) And (Not MirandoRank) And (Not MirandoGuildPanel) And (Not MirandoTravel) And (Not MirandoComerciarUsu) And (Not MirandoBanco) And (Not MirandoComerciar) And (Not MirandoConcentracion) And (Not MirandoCuenta) Then
        
                SendTxt.visible = True
                
                SendTxt.SetFocus

            End If
                  
    End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
    
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
         
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewPic.Width
    End If
       
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
          
    Dim A As Long
    
    Call Render_Exp(True)
          
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
    
    If Not CoordBloqued Then
        Map_UpdateLabel
    End If
    
    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
          
    If SendTxt.visible Then
        SendTxt.SetFocus
    End If
    
    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    
    If Not CursorSelected = 1 Then
        Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorGeneral, IDC_ARROW)
        CursorSelected = 1
    End If
    
    
    
   ' If MirandoObjetos Then
    '    FrmObject_Info.Close_Form
   ' End If
    
End Sub

Private Sub CMSG_Click()
    Call Audio.PlayInterface(SND_CLICK)

    If Not CharTieneClan And Not bCMSG Then
        Call AddtoRichTextBox(FrmMain.RecTxt, "¡No perteneces a ningún clan!", 0, 200, 200, False, False, True)

    Else
        If PMSGimg Then Call imgPMSG_Click
        
        bCMSG = Not bCMSG

        If bCMSG Then
            

            imgButton(2).Picture = LoadPicture(App.path & "\resource\interface\main\CMSG.jpg")
            Call AddtoRichTextBox(FrmMain.RecTxt, "Todo lo que digas sera escuchado por tu clan.", 0, 200, 200, False, False)
            HablaTemp = "/CMSG "
            
        Else
            Call AddtoRichTextBox(FrmMain.RecTxt, "Dejas de ser escuchado por tu clan. ", 0, 200, 200, False, False)
            CMSG.Picture = Nothing
            imgButton(2).Picture = Nothing
            HablaTemp = vbNullString
        End If
    End If

End Sub


Private Sub imageparty_click()
    
    If MsgBox("¿Estás seguro que deseas crear un grupo?", vbYesNo) = vbYes Then
        WritePartyClient 1
    End If

End Sub

Public Function LeerJPG(ByRef file_path As String) As Byte()

    If Len(Dir$(file_path)) <> 0 Then

        Dim fFile  As Integer

        Dim Temp() As Byte
    
        fFile = FreeFile()
        
        ReDim Temp(FileLen(file_path)) As Byte

        Open file_path For Binary As #fFile

        Get #fFile, , Temp()

        Close #fFile
 
        LeerJPG = Temp()
 
    End If

End Function


Private Sub imgPMSG_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    '----Boton partys Style TDS by IRuleDK----
    'PMSG = False 'Nos fijamos que no este activado con la tecla suprimir

    If Not PMSGimg Then
        If bCMSG Then Call CMSG_Click
        
        PMSGimg = True
        
       ' #If Classic = 0 Then
      '  imgPMSG.Picture = LoadPicture(DirInterface & "main\ChatPartyActivo.jpg") 'Grafico del botón estilo tds
      '  #Else
             imgButton(3).Picture = LoadPicture(DirInterface & "main\PMSG.jpg") 'Grafico del botón estilo tds
       ' #End If
        Call AddtoRichTextBox(FrmMain.RecTxt, "Todo lo que digas sera escuchado por tu party. ", 255, 200, 200, False, False)
        HablaTemp = "/PMSG "
    Else 'si ya estaba apretado lo desactivamos
        PMSGimg = False 'desactivamos el boton
        imgPMSG.Picture = Nothing
        imgButton(3).Picture = Nothing
        Call AddtoRichTextBox(FrmMain.RecTxt, "Dejas de ser escuchado por tu party. ", 255, 200, 200, False, False)
        HablaTemp = vbNullString
    End If

End Sub

Private Sub Labelgm1_Click()
    Call ParseUserCommand("/telep yo 1 50 50")
End Sub

Private Sub Labelgm2_Click()

    If MsgBox("Esta todo listo para empezar la daga rusa?", vbYesNo, "Daga rusa") = vbYes Then
        Call ParseUserCommand("/RMSG Luego de la cuenta envien los interesados en la Daga Rusa")
        Call ParseUserCommand("/cr 5")
    End If

End Sub

Private Sub Labelgm3_Click()
    Call ParseUserCommand("/cr 5")
End Sub

Private Sub Labelgm4_Click()
    frmPanelGm.Show , FrmMain
End Sub

Private Sub Labelgm5_Click()
    Call ParseUserCommand("/online")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If

End Sub

Private Sub imgClanes_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call WriteGuilds_Required(0)
End Sub

Private Sub imgGrupo_Click()
    Call Audio.PlayInterface(SND_CLICK)
End Sub

Private Sub imgInvScrollDown_Click()
    Call Inventario.ScrollInventory(True)
End Sub

Private Sub imgInvScrollUp_Click()
    Call Inventario.ScrollInventory(False)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
    
    'If Not CursorSelected = 1 Then
        'Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorGeneral, IDC_ARROW)
        'CursorSelected = 1
    'End If
    
    LastButtonPressed.ToggleToNormal

    
   ' If MirandoObjetos Then
     '   FrmObject_Info.Close_Form
   ' End If
    
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub lblHabla_Click(Index As Integer)

    Dim A As Long
    
    If Index = 0 Then
        If picHabla.visible Then
            picHabla.visible = False
        Else
            picHabla.visible = True
        End If
        
        Exit Sub

    End If
    
    For A = 1 To 3
        lblHabla(A).ForeColor = vbWhite
    Next A
    
    Select Case Index

        Case 0

        Case 1
            HablaTemp = vbNullString
            lblHabla(1).ForeColor = vbYellow

        Case 2 ' Grito
            HablaTemp = "-"
            lblHabla(2).ForeColor = vbYellow

        Case 3 ' Susurro
            HablaTemp = "\" & InputBox("Elige el nombre de la persona a la que deseas susurrar") & " "
            lblHabla(3).ForeColor = vbYellow
            
        Case 4 ' Emojis visibles
            Call MsgBox("Algunas computadoras presentan problemas con el componente que utilizamos para los emojis. Cuando se encuentre reparado en su totalidad, volveremos a activarlo")
            'Call ShellExecute(hWnd, "open", "https://es.piliapp.com/facebook-symbols/", vbNullString, vbNullString, 1)
    End Select
    
    If Index > 0 Then picHabla.visible = False
End Sub


Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub Capturar_Guardar()
      
    Dim filePath As String
    filePath = App.path & "\SCREENSHOTS\Prueba.jpg"
    'FilePath = "c:\pantalla.bmp"
    
    If FileExist(filePath, vbArchive) Then
        Kill filePath
    End If
    
    Clipboard.Clear
      
    ' Manda la pulsación de teclas para capturar la imagen de la pantalla
    Call keybd_event(44, 2, 0, 0)
      
    DoEvents
    
    ' Si el formato del clipboard es un bitmap
    If Clipboard.GetFormat(vbCFBitmap) Then
        'Guardamos la imagen en disco
        SavePicture Clipboard.GetData(vbCFBitmap), filePath
    End If
    
   
End Sub

Private Sub MainViewPic_Click()
     
    Dim Pt As Point

    GetCursorPos Pt
    
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        #If Testeo = 1 Then
    
            If FrmBody.visible Then

                Dim Body As Integer
                
                If MapData(tX, tY).CharIndex > 0 Then
                    TempCharIndex = MapData(tX, tY).CharIndex

                End If
                
                If TempCharIndex > 0 Then
                    Body = CharList(TempCharIndex).iBody
                    
                    FrmBody.txt.Text = Body
                    FrmBody.x1.Text = CharList(TempCharIndex).Body.BodyOffSet(1).X
                    FrmBody.y1.Text = CharList(TempCharIndex).Body.BodyOffSet(1).Y
                    
                    FrmBody.x2.Text = CharList(TempCharIndex).Body.BodyOffSet(2).X
                    FrmBody.y2.Text = CharList(TempCharIndex).Body.BodyOffSet(2).Y
                    
                    FrmBody.x3.Text = CharList(TempCharIndex).Body.BodyOffSet(3).X
                    FrmBody.y3.Text = CharList(TempCharIndex).Body.BodyOffSet(3).Y
                    
                    FrmBody.x4.Text = CharList(TempCharIndex).Body.BodyOffSet(4).X
                    FrmBody.y4.Text = CharList(TempCharIndex).Body.BodyOffSet(4).Y
                    
                End If

            End If

        #End If
    

        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then

                '[/ybarra]
                If UsingSkill = 0 Then
                  '  Call CountPacketIterations(packetControl(ClientPacketID.LeftClick), 150)
                    Call WriteLeftClick(tX, tY)
                Else
                      
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        FrmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        Call RestoreLastCursor(IDC_CROSS)
                        Exit Sub

                    End If
                          
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                    
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                                If Not MainTimer.Check(TimersIndex.Arrows) Then
                                    FrmMain.MousePointer = vbDefault
                                    UsingSkill = 0
                                    Call RestoreLastCursor(IDC_CROSS)
                                    Call MainTimer.Restart(TimersIndex.Attack) ' Prevengo flecha-golpe
                                    Call MainTimer.Restart(TimersIndex.CastSpell) ' flecha-hechizo
                                    Exit Sub

                                End If
                        
                            End If

                        End If

                    End If
                          
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then  'Check if attack interval has finished.
                            If MainTimer.Check(TimersIndex.CastSpell) Then  'Corto intervalo de Golpe-Magia
                                'frmMain.MousePointer = vbDefault
                                ' UsingSkill = 0
                                ' Call RestoreLastCursor(IDC_CROSS)
                                
                                Call MainTimer.Restart(TimersIndex.CastAttack)
                                Call MainTimer.Restart(TimersIndex.CastSpell)
                                ' Exit Sub

                            End If

                        End If

                    End If


                    FrmMain.MousePointer = vbDefault
                    Call RestoreLastCursor(IDC_CROSS)
                    Call WriteWorkLeftClick(tX, tY, UsingSkill, Pt.X, Pt.Y)
                    UsingSkill = 0

                End If

            Else

                ' Descastea
                If UsingSkill = Magia Or UsingSkill = Proyectiles Then
                    FrmMain.MousePointer = vbDefault
                    Call RestoreLastCursor(IDC_CROSS)
                    UsingSkill = 0
                Else

                    If Not Comerciando Then
                        If NpcIndex_MouseHover > 0 Then
                            Call Setting_MenuInfo(NpcIndex_MouseHover, False)
                        Else
                            Call WriteRightClick(tX, tY, Pt.X, Pt.Y)

                        End If

                    End If

                End If

            End If
                  
        ElseIf (MouseShift And 1) = 1 Then

            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)

                End If

            End If

        End If

    End If

End Sub

Private Sub MainViewPic_DblClick()
    Form_DblClick

    If SendTxt.visible Then
        SendTxt.SetFocus
    End If

End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseBoton = Button
    MouseShift = Shift

    Call ConvertCPtoTP(X, Y, tX, tY)
          
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
          
    MouseX = X
    MouseY = Y
          
    Call ConvertCPtoTP(X, Y, tX, tY)
          
    If Inventario.sMoveItem And Not vbKeyShift Then
        General_Drop_X_Y tX, tY
        Inventario.uMoveItem = False
    Else

        If Inventario.sMoveItem And vbKeyShift Then
            
                
            FrmCantidad.Show , FrmMain
            Call FrmCantidad.SetDropDragged(X, Y)
        End If
    End If

    If SendTxt.visible Then
        SendTxt.SetFocus
    End If

End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
          
    clicX = X
    clicY = Y
           
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem(0)
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(FrmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub


Private Sub coord_click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If CoordBloqued Then
        CoordBloqued = False
    Else
        CoordBloqued = True
    End If
    
End Sub


Private Sub picHabla_Click()
    
    picHabla.visible = False
End Sub



Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
          
    
    
    With GlobalCounters
        If .StrenghtAndDextery > 0 Then
        
            .StrenghtAndDextery = .StrenghtAndDextery - 1
        
        
        End If
    End With
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With

    Else

        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then

            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
                
                Inventario.uMoveItem = False
                Inventario.sMoveItem = False
            Else

                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    
                    
                    If Not Comerciando Then FrmCantidad.Show , FrmMain
                    Call FrmCantidad.SetDropGround
                End If
            End If
        End If
    End If

End Sub

Private Sub AgarrarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With

    Else
        Call WritePickUp
    End If

End Sub

Private Sub UsarItem(ByVal SecondaryClick As Byte)
    
    
    'If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub
    'If Not CheckInterval(SecondaryClick) Then Exit Sub
    'If (FrameTime - Intervalos(SecondaryClick).ModifyTime) <= 200 Then Exit Sub
          
    'ShowConsoleMsg
    
    Dim Value As Long

    If Comerciando Then Exit Sub
          
    Dim strTemp As String, A As Long
    
    If SecondaryClick Then
        If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            CheckingDouble = CheckingDouble + 1
            CheckingDoubleValue(CheckingDouble) = Value
            
            If CheckingDouble >= 10 Then

                For A = 1 To 10
                    strTemp = strTemp & CheckingDoubleValue(A) & ", "
                Next A
                
                Call WriteDenounce("[SEGURIDAD]: Posible uso de Mouse-Gamer. Velocidades: " & strTemp)
                CheckingDouble = 0
            End If
        Else
            CheckingDouble = 0

        End If
    
    Else
         If Not MainTimer.Check(TimersIndex.UseItemWithU) Then
            CheckingDouble_U = CheckingDouble_U + 1
            CheckingDoubleValue_U(CheckingDouble_U) = Value
            
            If CheckingDouble_U >= 10 Then

                For A = 1 To 10
                    strTemp = strTemp & CheckingDoubleValue_U(A) & ", "
                Next A
                
                Call WriteDenounce("[SEGURIDAD]: Posible uso de Mouse-Gamer. Velocidades: " & strTemp)
                CheckingDouble_U = 0
            Else
                CheckingDouble_U = 0
            End If

        End If
        
       
    End If

    
    Dim ItemIndex As Integer
              
    ItemIndex = Inventario.SelectedItem
          
    If (ItemIndex > 0) And (ItemIndex < MAX_INVENTORY_SLOTS + 1) Then
        If Inventario.ObjType(ItemIndex) <> eOBJType.otBarcos And Inventario.ObjType(ItemIndex) <> eOBJType.otTeleportInvoker Then
            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If

        End If
        
        If Not IsActionParaCliente(Inventario.ObjIndex(ItemIndex)) Then
            Call WriteUseItem(ItemIndex, SecondaryClick, Value)
        End If
        
        Call AssignedInterval(SecondaryClick)
        
    End If

    Call Inventario.DrawInventory
End Sub

Private Sub EquiparItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With

    Else

        If Comerciando Then Exit Sub
   
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteEquipItem(Inventario.SelectedItem)
    End If

End Sub




Private Sub tAnuncios_Timer()
     Anuncio_Update_Next_Text
End Sub


Private Sub Timer1_Timer()
    LoopInterval
    
   
End Sub

Private Sub TimerPing_Timer()

    Static I As Integer

    '//
    I = I + 1


    If I >= 3 Then
        I = 0

    End If

End Sub



Private Sub tMapName_Timer()
    Map_TimeRender = 255
    tMapName.Enabled = False
End Sub

Private Sub tMessage_Timer()
    Static Minutes As Integer
    
    Minutes = Minutes + 1
    
    If Minutes >= 20 Then
       ' Call SelectedSpamMessage
        Minutes = 0
        Call SearchDesterium
    End If
    
End Sub



Private Sub tmrBlink_Timer()

    Dim A As Long
    
    For A = Me.lblFuerza.LBound To Me.lblFuerza.UBound
        If bLastBrightBlink Then
            FrmMain.lblFuerza(A).ForeColor = getStrenghtColor(UserFuerza)
            FrmMain.lblAgilidad(A).ForeColor = getDexterityColor(UserAgilidad)
        Else
            FrmMain.lblFuerza(A).ForeColor = vbWhite
            FrmMain.lblAgilidad(A).ForeColor = vbWhite
        End If
    
    Next A
    bLastBrightBlink = Not bLastBrightBlink
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()

    If Not hlst.visible Then
        DesactivarMacroHechizos

        Exit Sub

    End If
          
    'Macros are disabled if focus is not on Argentum!
    'If Not Application.IsAppActive() Then
    'DesactivarMacroHechizos

    'Exit Sub

    'End If
          
    If Comerciando Then Exit Sub

    If hlst.List(hlst.ListIndex) <> "(Vacio)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
          
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
      
        Dim Pt As Point
    GetCursorPos Pt
    Call WriteWorkLeftClick(tX, tY, UsingSkill, Pt.X, Pt.Y)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()

    If hlst.List(hlst.ListIndex) <> "(Vacio)" Then

        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With

        Else
            If ClientSetup.bConfig(eSetupMods.SETUP_BOTONLANZAR) = 1 Then
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then Exit Sub
            End If
            
            Dim SpellIndex As Integer
            SpellIndex = UserHechizos(hlst.ListIndex + 1)
            
            Call WriteCastSpell(hlst.ListIndex + 1)
            
            If SpellIndex > 0 Then
                If Hechizos(SpellIndex).AutoLanzar = 0 Then
                    Call WriteWork(eSkill.Magia)
                End If
            End If
        End If
    End If
    
    
    
    'Call TestingSound
End Sub

Private Sub TestingSound()
    
    'Static Effect As Byte
    
    'Effect = Effect + 1
    
    'If Effect >= 40 Then Effect = 1
    
    'Call Audio.SetReverb(Effect)
       ' Call Audio.SetReverb(REVERB_Off)  ' EFECTO DESACTIVADO
        
    'End If

End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    SetHand
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub PicInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    MouseX = X
    MouseY = Y
    
    If Not Inventario.uMoveItem Then
        PicInv.MousePointer = vbDefault
    End If
    
    If Not CursorSelected = 2 Then
        Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorInv, IDC_ARROW)
        CursorSelected = 2
    End If
End Sub

Private Sub Form_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then
        
        If NpcIndex_MouseHover > 0 Then
            If NpcList(NpcIndex_MouseHover).NpcType = eNPCType.Banquero Then
                Call Setting_MenuInfo(NpcIndex_MouseHover, False)
            Else
                Call WriteDoubleClick(tX, tY, 0)
            End If
            
        Else
            Call WriteDoubleClick(tX, tY, 0)
        End If
        
            
    End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            
            If FrmMenu.visible Then
                FrmMain.SetFocus
                Unload FrmMenu
            Else
                If Not MirandoComerciar And Not MirandoBanco Then
                    FrmMenu.Show , FrmMain
                End If
                
            End If
             
            UnloadAllForms_ButPrincipal
              
        End If
    
End Sub

'Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
'  KeyCode = 0
'End Sub

'Private Sub hlst_KeyPress(KeyAscii As Integer)
'  KeyAscii = 0
'End Sub

'Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
'   KeyCode = 0
'End Sub

Public Sub Label4_Click()
    
    Dim Pt As Point

    Dim A  As Long

    Call Audio.PlayInterface(SND_CLICK)
    'If Not MainTimer.Check(TimersIndex.Packet250) Then Exit Sub
    
    InvEqu.Picture = Nothing

    Panel = eVentanas.vInventario

    GetCursorPos Pt
          
    'If Panel <> LastPanel Then
    
    If ModoTab Then
        Pt.X = 0
        Pt.Y = 0
        ModoTab = False
    End If
    
    Call WriteSetPanelClient(Panel, 255, Pt.X, Pt.Y)
    LastPanel = Panel
    'End If

    ' Activo controles de inventario
    PicInv.visible = True
    
    #If ModoBig > 0 Then
        FrmMain.GldLbl(0).visible = False
        FrmMain.GldLbl(1).visible = True
    #End If
    ' Desactivo controles de hechizo
    
    #If ModoBig = 0 Then
    picHechiz.visible = False
    #Else
    picHechiz(1).visible = False
    
    #End If
    
    CmdLanzar.visible = False
    ImgInfo.visible = False
    
    imgButton(6).visible = True
    imgButton(7).visible = True
    imgButton(8).visible = True
    Inventario.DrawInventory

End Sub

Private Sub label4_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
    
    SetHand
End Sub

Public Sub Label7_Click()
    
    Call Audio.PlayInterface(SND_CLICK)
        
    'If Not MainTimer.Check(TimersIndex.Packet250) Then Exit Sub

    #If ModoBig = 0 Then
        If ClientSetup.bConfig(eSetupMods.SETUP_INTERFAZTDS) Then
            InvEqu.Picture = LoadPicture(DirInterface & "main\SpellClassic.jpg")
        Else
            InvEqu.Picture = LoadPicture(DirInterface & "main\SpellClassic2.jpg")

        End If
        
    #Else
        InvEqu.Picture = LoadPicture(DirInterface & "main\spellclasicx2.jpg")

    #End If

    Panel = eVentanas.vHechizos

    'If Panel <> LastPanel Then

    Dim TempInv As Byte

    Dim Pt      As Point

    Dim A       As Long
    
    GetCursorPos Pt
          
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        TempInv = CByte(Inventario.SelectedItem)
    Else
        TempInv = 255 ' @@ Pasamos y tenemos ningun slot seleccionado entonces 255 ...

    End If
              
        
    If ModoTab Then
        Pt.X = 0
        Pt.Y = 0
        ModoTab = False
    End If
    
    Call WriteSetPanelClient(Panel, TempInv, Pt.X, Pt.Y)
    LastPanel = Panel

    'End If
          
    ' Activo controles de hechizos
    #If ModoBig = 0 Then
    picHechiz.visible = True
    #Else
        picHechiz(1).visible = True
    #End If
    
    CmdLanzar.visible = True
    ImgInfo.visible = True
    
    #If ModoBig > 0 Then

        For A = FrmMain.GldLbl.LBound To FrmMain.GldLbl.UBound
            FrmMain.GldLbl(A).visible = False
        Next A

    #End If
    
    ' Desactivo controles de inventario
    PicInv.visible = False
    
    imgButton(6).visible = False
    imgButton(7).visible = False
    imgButton(8).visible = False
    'imgInvScrollUp.Visible = False
    'imgInvScrollDown.Visible = False
    
    
    
   ' If MirandoObjetos Then
    '    FrmObject_Info.Close_Form
   ' End If

End Sub

Private Sub Label7_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
    
    SetHand
End Sub

Private Sub PicInv_DblClick()
    
    If Inventario.SelectedItem = 0 Then Exit Sub
    Inventario.DrawInventory

    If (mouse_Down <> False) And (mouse_UP = True) Then Exit Sub

    Dim Value As Long, strTemp As String, A As Long

    mouse_UP = False
    ' x button
    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
     
    Inventario.uMoveItem = False
              
    If MouseInvBoton = vbRightButton Then Exit Sub

    Dim ObjIndex As Integer
    
    ObjIndex = Inventario.ObjIndex(Inventario.SelectedItem)
    
    If ObjIndex > 0 Then
        If (ObjData(ObjIndex).ObjType = otarmadura Or _
            ObjData(ObjIndex).ObjType = otWeapon Or _
            ObjData(ObjIndex).ObjType = otcasco Or _
            ObjData(ObjIndex).ObjType = otescudo Or _
            ObjData(ObjIndex).ObjType = otAnillo Or _
            ObjData(ObjIndex).ObjType = otMagic Or _
            ObjData(ObjIndex).ObjType = otFlechas Or _
            ObjData(ObjIndex).ObjType = otPendienteParty) Then
            
            If Not Inventario.Equipped(Inventario.SelectedItem) Then
                Call EquiparItem
            Else
                Call UsarItem(1)
            End If
        Else
            Call UsarItem(1)
        End If
    
    End If
    
    Inventario.DrawInventory
End Sub

Private Sub PicInv_Click()
    Call Audio.PlayInterface(SND_CLICK)
End Sub

Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
                                 
    '    / x button
    If (mouse_Down = False) Then Exit Sub
    mouse_Down = False
    mouse_UP = True
    '    / x button
           
    Inventario.uMoveItem = False
    MouseInvBoton = Button
    
    If dobleclick.Interval = 0 Then dobleclick.Interval = 1000
    
    If Button = 1 Then
        dobleclick.Interval = 1000
        totalclicks = totalclicks + 1

    End If

End Sub

Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
          
    If SendTxt.visible Then
        SendTxt.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoForo) And (Not MirandoEstadisticas) And (Not MirandoCantidad) And _
           (Not MirandoRank) And (Not MirandoGuildPanel) And (Not MirandoTravel) And _
        (Not MirandoComerciarUsu) And (Not MirandoBanco) And (Not MirandoComerciar) And (Not MirandoConcentracion) And (Not MirandoCuenta) Then
               
       ' If picInv.Visible Then
           ' picInv.SetFocus
       ' ElseIf hlst.Visible Then
        '    hlst.SetFocus
      '  End If
    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

  '  If picInv.Visible Then
       ' picInv.SetFocus
   ' Else
    '    hlst.SetFocus
   ' End If

End Sub

Private Function InGameArea() As Boolean

    '***************************************************
    'Author: NicoNZ
    'Last Modification: 04/07/08
    'Checks if last click was performed within or outside the game area.
    '***************************************************
    If clicX < MainViewPic.Left Or clicX > MainViewPic.Left + MainViewPic.Width Then Exit Function

    If clicY < MainViewPic.Top Or clicY > MainViewPic.Top + MainViewPic.Height Then Exit Function
          
    InGameArea = True
End Function



Private Sub tUpdateInactive_Timer()
    Call WriteUpdateInactive
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call AddtoRichTextBox(FrmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
End Sub

Private Sub PicInv_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Dim Position  As Integer
    Dim I         As Long
    Dim file_path As String
    Dim data()    As Byte
    Dim bmpInfo   As BITMAPINFO
    Dim handle    As Integer
    Dim bmpData   As StdPicture
    
    '    / x button
    mouse_Down = True
    mouse_UP = False
    '    / x button

    If Inventario.SelectedItem = 0 Then Exit Sub
    
    If (Button = vbRightButton) And (Not Comerciando) Then
        If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then

            Last_I = Inventario.SelectedItem

            If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                          
                Position = Search_GhID(3057)
                  
                If Position = 0 Then
                    I = 3057
                    Call Get_Image(DirGraficos & GRH_RESOURCE_FILE_DEFAULT, CStr(3057), data, False)
                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1)
                    FrmMain.ImageList1.ListImages.Add , "g3057", Picture:=bmpData
                    Position = FrmMain.ImageList1.ListImages.Count
                    Set bmpData = Nothing
                    
                End If
                  
                Inventario.uMoveItem = True
                  
                Set PicInv.MouseIcon = FrmMain.ImageList1.ListImages(Position).ExtractIcon
                FrmMain.PicInv.MousePointer = vbCustom

                Exit Sub

            End If
        End If
    End If

End Sub

Private Function Search_GhID(ByVal gh As Integer) As Integer

    Dim I As Long

    For I = 1 To FrmMain.ImageList1.ListImages.Count

        If FrmMain.ImageList1.ListImages(I).Key = "g" & CStr(gh) Then
            Search_GhID = I
            Exit For
        End If

    Next I

End Function

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then

        'Say
        If LenB(stxtbuffercmsg) <> 0 Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = vbNullString
        KeyCode = 0
              
       'If picInv.Visible Then
       '     picInv.SetFocus
      '  Else
      '      hlst.SetFocus
      '  End If
    End If

End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub UnicodeRtfTextBox1_Click()

End Sub
Private Sub SendTxt_Change()
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
    '**************************************************************

    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = vbNullString
    Else

        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I         As Long

        Dim TempStr   As String

        Dim CharAscii As Integer
              
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, I, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                TempStr = TempStr & Chr$(CharAscii)
            End If

        Next I
              
        If TempStr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = TempStr
        End If
        
        stxtbuffer = SendTxt.Text
        
        Commands_Search stxtbuffer
        SendTxt.Text = stxtbuffer
        FrmMain.SendTxt.SetFocus
        
        
    End If

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = CustomKeys.BindedKey(eKeyType.mKeyTalk) Then 'Si se apretó enter entonces:
        
        ' No enviamos mensajes vacios a los clanes pero borramos cartel.
        If stxtbuffer = " " Or stxtbuffer = "  " Then
            Call ParseUserCommand(stxtbuffer)
        Else
            If LenB(stxtbuffer) <> 0 Then
                Call ParseUserCommand(HablaTemp & stxtbuffer)
            End If
        End If
        
        stxtbuffer = vbNullString ' // Mejor vbnullstring que vbnullstring
        SendTxt.Text = vbNullString ' // Mejor vbnullstring que vbnullstring
        KeyCode = 0
        SendTxt.visible = False
             
      '  If picInv.Visible Then
        '    picInv.SetFocus
       ' Else
     '       hlst.SetFocus
     '   End If
    End If
       
End Sub





Private Sub UpdateMapa_Timer()
    If RenderizandoIndex = UserMap Then
        Call wGL_Graphic.Capture(FrmMain.MiniMapa.hWnd, MiniMap_FilePath & RenderizandoIndex & ".png")
    End If
    
    UpdateMapa.Enabled = False
    RenderizandoMap = False
End Sub


Private Sub macrotrabajo_Timer()

    If Inventario.SelectedItem = 0 Or UserMinSTA <= 5 Then
        DesactivarMacroTrabajo

        Exit Sub

    End If
              Dim Pt As Point
    GetCursorPos Pt
    
    If (UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill, Pt.X, Pt.Y)
        UsingSkill = 0
    End If
          
    Call UsarItem(0)
    Inventario.DrawInventory
End Sub


Public Sub ActivarMacroTrabajo()
    MacroTrabajo.Interval = IntervaloUserPuedeTrabajar
    MacroTrabajo.Enabled = True
    Call AddtoRichTextBox(FrmMain.RecTxt, "Empiezas a trabajar", 0, 200, 200, False, False, True)

End Sub

Public Sub DesactivarMacroTrabajo()

    MacroTrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(FrmMain.RecTxt, "Dejas de trabajar", 0, 200, 200, False, False, True)
       
End Sub


' Lista Gráfica de Hechizos
Private Sub picHechiz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then Y = Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1

If X < picHechiz.ScaleWidth - 10 Then
    hlst.ListIndex = Int(Y / hlst.Pixel_Alto) + hlst.Scroll
    hlst.DownBarrita = 0

Else
    hlst.DownBarrita = Y - hlst.Scroll * (picHechiz.ScaleHeight - hlst.BarraHeight) / (hlst.ListCount - hlst.VisibleCount)
End If
End Sub

Private Sub picHechiz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseShift = Shift

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then yy = Int(picHechiz.ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
    If hlst.DownBarrita > 0 Then
        hlst.Scroll = (Y - hlst.DownBarrita) * (hlst.ListCount - hlst.VisibleCount) / (picHechiz.ScaleHeight - hlst.BarraHeight)
    Else
        hlst.ListIndex = Int(yy / hlst.Pixel_Alto) + hlst.Scroll

       ' If ScrollArrastrar = 0 Then
           ' If (Y < yy) Then hlst.Scroll = hlst.Scroll - 1
           ' If (Y > yy) Then hlst.Scroll = hlst.Scroll + 1
        'End If
    End If
ElseIf Button = 0 Then
    hlst.ShowBarrita = X > picHechiz.ScaleWidth - hlst.BarraWidth * 2
End If
End Sub

Private Sub picHechiz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
hlst.DownBarrita = 0
End Sub

Private Sub MouseData_MouseWheel(ByVal wKeys As Long, _
                                 ByVal zDelta As Long, _
                                 ByVal xPos As Long, _
                                 ByVal yPos As Long)

    ' Cuando se mueve la rueda del ratón
    'Call MsgBox("Se ha movido la rueda del ratón pos arriba +" & wKeys & ", " & zDelta)

            If zDelta <= 240 Then
                 hlst.Scroll = hlst.Scroll - 1
            Else

                 hlst.Scroll = hlst.Scroll + 1

            End If

        hlst.ShowBarrita = True
   
End Sub

Private Sub MouseData_Activate()
    ' Cuando se activa
    'Label2(2).Caption = "Activate"
End Sub

Private Sub MouseData_Deactivate()
    ' Cuando se desactiva
    '(2).Caption = "Deactivate"
End Sub

Private Sub MouseData_DisplayChange(ByVal BitsPerPixel As Long, ByVal cxScreen As Long, ByVal cyScreen As Long)
    ' DisplayChange
End Sub

Private Sub MouseData_FontChange()
    ' FontChange
End Sub

Private Sub MouseData_LowMemory()
    ' LowMemory
End Sub

Private Sub MouseData_MenuSelected(ByVal mnuItem As Long, mnuFlags As eWSCMF, ByVal hMenu As Long)
    ' MenuSelected
    'List1.AddItem "Has seleccionado el menú: " & mnuItem
End Sub

Private Sub MouseData_MouseEnter()
    ' Cuando entra el ratón en el formulario
   ' Label1 = "MouseEnter"
End Sub

Private Sub MouseData_MouseEnterOn(unControl As Object)
    ' Cuando entra el ratón en uno de los controles
   ' Label1.Caption = "MouseEnterOn: " & unControl.Name
End Sub

Private Sub MouseData_MouseLeave()
    ' Cuando sale el ratón
   ' Label1 = "MouseLeave"
End Sub

Private Sub MouseData_MouseLeaveOn(unControl As Object)
    ' Cuando sale el ratón de un control
   ' Label1.Caption = "MouseLeaveOn: " & unControl.Name
End Sub

Private Sub MouseData_Move(ByVal wLeft As Long, ByVal wTop As Long)
    ' Cuando se mueve el formulario
    'Label2(0).Caption = "Posición del formulario: Left= " & wLeft & ", Top= " & wTop
End Sub

Private Sub MouseData_SetCursor(unControl As Object, ByVal HitTest As eWSCHitTest, ByVal MouseMsg As Long)
    'Label2(2).Caption = "SetCursor: " & unControl.Name & ", HitTest:" & HitTest & ", MouseMsg: " & MouseMsg
End Sub

Private Sub MouseData_WindowPosChanged(ByVal wLeft As Long, ByVal wTop As Long, ByVal wWidth As Long, ByVal wHeight As Long)
    ' Cuando cambia la posición de la ventana
   ' Label2(1).Caption = "WndPosChanged: L:" & wLeft & ", T:" & wTop & ", W:" & wWidth & ", H:" & wHeight
End Sub

