VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   16785
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1119
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5505
      Index           =   1
      Left            =   3240
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2A3B6
      ScaleHeight     =   367
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   338
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   5070
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
      Left            =   6480
      Picture         =   "frmMain.frx":34F96
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   182
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   2730
      Begin VB.Image imgSocial 
         Height          =   225
         Index           =   3
         Left            =   1680
         Top             =   2580
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
         Index           =   1
         Left            =   1125
         Top             =   2625
         Width           =   225
      End
      Begin VB.Image imgSocial 
         Height          =   225
         Index           =   0
         Left            =   840
         Top             =   2625
         Width           =   225
      End
      Begin VB.Image imgParty 
         Height          =   375
         Left            =   1380
         Top             =   915
         Width           =   1155
      End
      Begin VB.Image imgFight 
         Height          =   375
         Left            =   1380
         Top             =   540
         Width           =   1155
      End
      Begin VB.Image imgResu 
         Height          =   300
         Left            =   1050
         Top             =   2265
         Width           =   315
      End
      Begin VB.Image imgDrag 
         Height          =   300
         Left            =   2265
         Top             =   2280
         Width           =   315
      End
      Begin VB.Image imgSeg 
         Height          =   300
         Left            =   2295
         Top             =   1890
         Width           =   315
      End
      Begin VB.Image imgGoStats 
         Height          =   330
         Left            =   105
         Top             =   60
         Width           =   1275
      End
      Begin VB.Image imgOpciones 
         Height          =   375
         Left            =   1410
         Top             =   1290
         Width           =   1110
      End
      Begin VB.Image imgClanes 
         Height          =   375
         Left            =   210
         Top             =   915
         Width           =   1110
      End
      Begin VB.Image imgEvents 
         Height          =   375
         Left            =   210
         Top             =   1290
         Width           =   1110
      End
      Begin VB.Image imgObjetive 
         Height          =   375
         Left            =   210
         Top             =   540
         Width           =   1110
      End
      Begin VB.Image CMSG 
         Height          =   300
         Left            =   675
         Top             =   1890
         Width           =   315
      End
      Begin VB.Image imgPMSG 
         Height          =   300
         Left            =   1005
         Top             =   1890
         Width           =   315
      End
   End
   Begin VB.PictureBox picHabla 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   7875
      ScaleHeight     =   1125
      ScaleWidth      =   1335
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   315
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
      Left            =   7725
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   150
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
      Left            =   345
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1560
      Visible         =   0   'False
      Width           =   7170
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      Height          =   195
      Left            =   8640
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   480
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      Index           =   0
      Left            =   5880
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   4
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6840
      Top             =   4320
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Top             =   4320
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1380
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2434
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":3FA72
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
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H8000000D&
      Height          =   7200
      Left            =   75
      MousePointer    =   99  'Custom
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   608
      TabIndex        =   10
      Top             =   1755
      Width           =   9120
      Begin VB.Timer tUpdateMS 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   2160
         Top             =   2400
      End
      Begin VB.Timer MacroTrabajo 
         Enabled         =   0   'False
         Left            =   8160
         Top             =   2640
      End
      Begin VB.Timer UpdateMapa 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   7680
         Top             =   2640
      End
      Begin VB.Timer tMapData 
         Enabled         =   0   'False
         Left            =   5040
         Top             =   2520
      End
      Begin VB.Timer tMapName 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5520
         Top             =   2520
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
         Left            =   7200
         Top             =   2640
      End
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
      Height          =   2880
      Left            =   7560
      MousePointer    =   4  'Icon
      Picture         =   "frmMain.frx":3FAEF
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   23
      Top             =   2640
      Width           =   2400
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      Index           =   0
      Left            =   9315
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":43558
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblarmor 
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
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   79
      Top             =   930
      Width           =   855
   End
   Begin VB.Label lblhelm 
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
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3030
      TabIndex        =   78
      ToolTipText     =   " "
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblShielder 
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
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1530
      TabIndex        =   77
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblWeapon 
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
      Height          =   255
      Left            =   30
      TabIndex        =   76
      Top             =   0
      Width           =   855
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
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   315
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
      Left            =   855
      TabIndex        =   73
      Top             =   0
      Width           =   315
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
      Left            =   14760
      TabIndex        =   72
      Top             =   1440
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
      Left            =   14280
      TabIndex        =   71
      Top             =   4110
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
      Left            =   16560
      TabIndex        =   70
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad de Ullathorpe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Index           =   0
      Left            =   11880
      TabIndex        =   69
      Top             =   7080
      Width           =   5715
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
      Left            =   13920
      TabIndex        =   68
      Top             =   840
      Width           =   3450
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
      Left            =   15855
      TabIndex        =   67
      Top             =   6720
      Width           =   840
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
      Left            =   15735
      TabIndex        =   66
      Top             =   6240
      Width           =   840
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
      Left            =   15630
      TabIndex        =   65
      Top             =   5640
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
      Left            =   15480
      TabIndex        =   64
      Top             =   5160
      Width           =   1350
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
      Left            =   15585
      TabIndex        =   63
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "47%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   705
      Index           =   2
      Left            =   12720
      TabIndex        =   62
      Top             =   3840
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
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
      Left            =   13560
      TabIndex        =   61
      Top             =   3000
      Visible         =   0   'False
      Width           =   2280
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
      Left            =   15600
      TabIndex        =   60
      Top             =   2280
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
      Left            =   15600
      TabIndex        =   59
      Top             =   2520
      Visible         =   0   'False
      Width           =   435
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
      Left            =   15600
      TabIndex        =   58
      Top             =   2040
      Visible         =   0   'False
      Width           =   435
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
      Left            =   12720
      TabIndex        =   57
      Top             =   7200
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
      Left            =   13800
      TabIndex        =   56
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgInvisible 
      Height          =   330
      Left            =   13440
      Picture         =   "frmMain.frx":48743
      Top             =   7155
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgParalisis 
      Height          =   330
      Left            =   12960
      Picture         =   "frmMain.frx":49926
      Top             =   7140
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   0
      Left            =   9870
      TabIndex        =   16
      Top             =   7515
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFC0&
      Height          =   165
      Index           =   0
      Left            =   9870
      TabIndex        =   15
      Top             =   7245
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00C0C0FF&
      Height          =   165
      Index           =   0
      Left            =   9870
      TabIndex        =   14
      Top             =   6975
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Index           =   0
      Left            =   9870
      TabIndex        =   12
      Top             =   6435
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFC0&
      Height          =   165
      Index           =   0
      Left            =   9870
      TabIndex        =   13
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   4
      Left            =   12480
      TabIndex        =   55
      Top             =   840
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   3
      Left            =   12240
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   2
      Left            =   12120
      TabIndex        =   53
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblsed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Index           =   1
      Left            =   12120
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFC0&
      Height          =   165
      Index           =   4
      Left            =   12000
      TabIndex        =   51
      Top             =   1200
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFC0&
      Height          =   165
      Index           =   3
      Left            =   12000
      TabIndex        =   50
      Top             =   1200
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFC0&
      Height          =   165
      Index           =   2
      Left            =   12120
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Lblham 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      ForeColor       =   &H00C0FFC0&
      Height          =   165
      Index           =   1
      Left            =   12120
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00C0C0FF&
      Height          =   165
      Index           =   4
      Left            =   12000
      TabIndex        =   47
      Top             =   1680
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00C0C0FF&
      Height          =   165
      Index           =   3
      Left            =   12360
      TabIndex        =   46
      Top             =   1680
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00C0C0FF&
      Height          =   165
      Index           =   2
      Left            =   12480
      TabIndex        =   45
      Top             =   1800
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H00C0C0FF&
      Height          =   165
      Index           =   1
      Left            =   11880
      TabIndex        =   44
      Top             =   1800
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFC0&
      Height          =   165
      Index           =   4
      Left            =   12240
      TabIndex        =   43
      Top             =   2520
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFC0&
      Height          =   165
      Index           =   3
      Left            =   12240
      TabIndex        =   42
      Top             =   2520
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFC0&
      Height          =   165
      Index           =   2
      Left            =   12120
      TabIndex        =   41
      Top             =   2520
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1490/1490"
      ForeColor       =   &H00FFFFC0&
      Height          =   165
      Index           =   1
      Left            =   12360
      TabIndex        =   40
      Top             =   2520
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Index           =   4
      Left            =   12120
      TabIndex        =   39
      Top             =   6000
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Index           =   3
      Left            =   12120
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Index           =   2
      Left            =   12120
      TabIndex        =   37
      Top             =   6000
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "475/475"
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Index           =   1
      Left            =   12120
      TabIndex        =   36
      Top             =   6120
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   8
      Left            =   12120
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   7
      Left            =   12720
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   6
      Left            =   12360
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   5
      Left            =   12600
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   4
      Left            =   12600
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   3
      Left            =   12600
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   2
      Left            =   12120
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   1
      Left            =   12120
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   0
      Left            =   12120
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgInfo 
      Height          =   615
      Left            =   11400
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgMoveSpell 
      Height          =   255
      Index           =   1
      Left            =   11760
      Top             =   5760
      Width           =   255
   End
   Begin VB.Image imgMoveSpell 
      Height          =   255
      Index           =   0
      Left            =   11760
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad de Ullathorpe"
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Index           =   1
      Left            =   9450
      TabIndex        =   34
      Top             =   1635
      Width           =   2235
   End
   Begin VB.Image CmdLanzar 
      Height          =   525
      Left            =   9345
      MouseIcon       =   "frmMain.frx":4ABF8
      MousePointer    =   99  'Custom
      Top             =   5400
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Left            =   11355
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgCerrar 
      Height          =   315
      Left            =   11670
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgGoMenu 
      Height          =   330
      Left            =   10605
      Top             =   5985
      Width           =   1275
   End
   Begin VB.Image imgStats 
      Height          =   345
      Left            =   9375
      Top             =   180
      Width           =   375
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
      Left            =   13920
      TabIndex        =   29
      Top             =   2160
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
      TabIndex        =   24
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
      Left            =   5040
      TabIndex        =   21
      Top             =   105
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   9015
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   90
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
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   0
      Left            =   9420
      TabIndex        =   19
      Top             =   570
      Width           =   2415
   End
   Begin VB.Label lblporclvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   9645
      TabIndex        =   18
      Top             =   1275
      Width           =   2100
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000.000"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   10110
      TabIndex        =   17
      Top             =   5580
      Width           =   1185
   End
   Begin VB.Image STAShp 
      Height          =   180
      Left            =   9420
      Picture         =   "frmMain.frx":4AD4A
      Top             =   6450
      Width           =   2310
   End
   Begin VB.Image MANShp 
      Height          =   180
      Left            =   9420
      Picture         =   "frmMain.frx":4BF9A
      Top             =   6720
      Width           =   2310
   End
   Begin VB.Image Hpshp 
      Height          =   180
      Left            =   9420
      Picture         =   "frmMain.frx":4CFE4
      Top             =   6990
      Width           =   2310
   End
   Begin VB.Image COMIDAsp 
      Height          =   180
      Left            =   9420
      Picture         =   "frmMain.frx":4E024
      Top             =   7260
      Width           =   2310
   End
   Begin VB.Image AGUAsp 
      Height          =   180
      Left            =   9420
      Picture         =   "frmMain.frx":4F062
      Top             =   7530
      Width           =   2310
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13200
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   180
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13470
      MousePointer    =   99  'Custom
      TabIndex        =   8
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
      Height          =   405
      Left            =   10605
      MousePointer    =   4  'Icon
      TabIndex        =   3
      Top             =   1890
      Width           =   1305
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
      Height          =   405
      Left            =   9240
      MousePointer    =   4  'Icon
      TabIndex        =   2
      Top             =   1890
      Width           =   1365
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   10185
      TabIndex        =   1
      Top             =   945
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   9765
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   7170
      Left            =   60
      Top             =   1740
      Visible         =   0   'False
      Width           =   9120
   End
   Begin VB.Image imgExp 
      Height          =   330
      Left            =   9600
      Picture         =   "frmMain.frx":4FFEE
      Top             =   1200
      Width           =   2070
   End
   Begin VB.Image InvEqu 
      Height          =   4110
      Left            =   9210
      Top             =   1860
      Width           =   2790
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ZoomIn As Boolean

Private SpellSelected As Byte
Public CoordBloqued As Boolean
Public PorcBloqued As Boolean
Public CursorSelected As Byte

' Detectar posicion del cursor.
Private Declare Function GetCursorPos Lib "user32.dll" (Pt As Point) As Long


Private Type Point

    X As Long
    Y As Long

End Type

'End Security
        
' x Auto Pots
Public Enum eVentanas

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
    
   ' hSwapCursor = SetClassLong(frmMain.hWnd, GLC_HCURSOR, hSwapCursor)
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
            Call frmOpciones.Show(vbModeless, frmMain)

        Case 1
            Call WriteGuilds_Required(0)

        Case 2
           CMSG_Click
        Case 3
            imgPMSG_Click
        Case 4 ' Retos
            Call ShowConsoleMsg("Ayuda» Los comandos /RETOSON y /RETOSOFF activan un Panel que te ayudará a ver la invitación en una nueva Ventana.", 150, 200, 148, True)
            Call ParseUserCommand("/RETOS")

        Case 5 ' Party
            Call ShowConsoleMsg("Ayuda» Es hora de enviar solicitudes para que usuarios formen un grupo contigo.. Haz clic sobre aquel que desees invitar y luego teclea F3.", 150, 200, 148, True)
            Call WritePartyClient(1)
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

Private Sub imgInfo_Click()

    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)

    End If

End Sub

Private Sub imgMinimize_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Me.WindowState = 1
End Sub

Private Sub imgMoveSpell_Click(Index As Integer)
        
        Call Audio.PlayInterface(SND_CLICK)
        
        If hlst.ListIndex = -1 Then Exit Sub
        
        Dim sTemp As String
    
        Select Case Index
            Case 0 'subir
                If hlst.ListIndex = 0 Then Exit Sub
            Case 1 'bajar
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        
        
        Select Case Index
            Case 0 'subir
                Call WriteMoveSpell(hlst.ListIndex, hlst.ListIndex + 1)
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
                
            Case 1 'bajar
                Call WriteMoveSpell(hlst.ListIndex + 1, hlst.ListIndex + 2)
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
                 
        End Select
End Sub

Private Sub imgMoveSpell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgObjetive_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call WriteQuestRequired
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
    
    Dim url As String
    
    Select Case Index
    
        Case 0 ' Instagram
            url = "https://www.instagram.com/ArgentumGame"
        Case 1 ' Youtube
             url = "https://www.instagram.com/ArgentumGame"
        Case 2 ' Facebook
             url = "https://www.facebook.com/ArgentumGame"
        Case 3 ' Discord
             url = "https://www.discord.argentumgame.com/"
    End Select
    
    Call ShellExecute(hWnd, "open", url, vbNullString, vbNullString, 1)
End Sub

Private Sub imgStats_Click()

  If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)
    
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegaronStats = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
End Sub



Private Sub SetHand()
    If Not CursorSelected = 3 Then
    Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorHand, IDC_ARROW)
    CursorSelected = 3
    End If
End Sub

Private Sub Label1_Click()
    frmOpciones.Show , frmMain

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

    
   ' #If Classic = 1 Then
        PorcBloqued = Not PorcBloqued
    
    '#End If
    
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
    
    
    #If ModoBig = 1 Then
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
        frmScreenShot.Show , frmMain
    #Else
        Call Audio.PlayInterface(SND_CLICK)
        
        #If Classic = 0 Then
            Call frmMapa.Show(vbModeless, frmMain)
        
        #Else
            Call frmMapaClassic.Show(vbModeless, frmMain)
        #End If
    #End If
    
End Sub


Private Sub picHechiz_DblClick(Index As Integer)
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
    
    #If ModoBig = 1 Then
        Me.picHechiz(1).Width = frmMain_Scalled.picHechiz(1).Width
        Me.picHechiz(1).Height = frmMain_Scalled.picHechiz(1).Height
    #End If
    
    ' Lista Gráfica
    Set hlst = New clsGraphicalList
    #If ModoBig = 0 Then
    Call hlst.Initialize(Me.picHechiz(0), RGB(200, 190, 190))
    #Else
    Call hlst.Initialize(Me.picHechiz(1), RGB(200, 190, 190))
    #End If
    
    'Drag And Drop
    Set dragInventory = Inventario

    ' Handles Form movement (drag and dr|op).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, 120
          
    Call LoadButtons
    
    #If ModoBig = 1 Then
        imgExp.visible = True
        imgExp.Picture = LoadPicture(DirInterface & "main\exp.jpg")
        picHechiz(1).Picture = LoadPicture(DirInterface & "main\spellfound_new.jpg")
        
    #Else
        imgExp.Picture = LoadPicture(DirInterface & "main\ExpBar.jpg")
        picHechiz(0).Picture = Nothing
    #End If

    
    
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

    #If Classic = 0 Then
   ' Call cBotonOpciones.Initialize(imgOpciones, vbNullString, GrhPath & "main\OpcionesActivo.jpg", vbNullString, Me)
  '  Call cBotonParty.Initialize(imgParty, vbNullString, GrhPath & "main\PartyActivo.jpg", vbNullString, Me)
    'Call cBotonEventos.Initialize(imgEvents, vbNullString, GrhPath & "main\EventosActivo.jpg", vbNullString, Me)
   ' Call cBotonClanes.Initialize(imgClanes, vbNullString, GrhPath & "main\ClanesActivo.jpg", vbNullString, Me)
   ' Call cBotonCerrar.Initialize(imgCerrar, vbNullString, GrhPath & "generic\BotonCerrarActivo.jpg", vbNullString, Me)
    'Call cBotonMinimizar.Initialize(imgMinimize, vbNullString, GrhPath & "generic\BotonMinimizarActivo.jpg", vbNullString, Me)
    'Call cBotonLanzar.Initialize(CmdLanzar, vbNullString, GrhPath & "main\LanzarActivo.jpg", vbNullString, Me)
    'Call cBotonObjetive.Initialize(imgObjetive, vbNullString, GrhPath & "main\MisionesActivo.jpg", vbNullString, Me)
   ' Call cBotonSkills.Initialize(imgStats, vbNullString, GrhPath & "main\SkillActivo.jpg", vbNullString, Me)
   ' Call cBotonRetos.Initialize(imgFight, vbNullString, GrhPath & "main\RetosActivo.jpg", vbNullString, Me)
    
    #End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '***************************************************
  
    If (Not SendTxt.visible) Then
    
        If esGM(UserCharIndex) Then
            If KeyCode = vbKeyI Then
                Call ParseUserCommand("/INVISIBLE")

            End If

        End If
        
        If KeyCode = vbKeyF5 Then
            Call ShowConsoleMsg("Ayuda» Los comandos /RETOSON y /RETOSOFF activan un Panel que te ayudará a ver la invitación en una nueva Ventana.", 150, 200, 148, True)
            Call ParseUserCommand("/RETOS")
                
            Exit Sub
        ElseIf KeyCode = vbKeyF7 Then
            Call ShowConsoleMsg("Ayuda» Es hora de enviar solicitudes para que usuarios formen un grupo contigo.. Haz clic sobre aquel que desees invitar y luego teclea F3.", 150, 200, 148, True)
            Call WritePartyClient(1)
            Exit Sub

        End If
                    
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            
            Select Case KeyCode
            
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
                        Call AddtoRichTextBox(frmMain.RecTxt, "Haz alcanzado el máximo de envio de 1 FotoDenuncia por minuto. Esperá unos instantes y volve a intentar.", 0, 200, 200, False, False, True)
        
                        Exit Sub
        
                    End If

                    'Aca guardamos el string que nos devuelve FotoD_Capturar.
                    Dim nString As String
        
                    FotoD_Capturar nString
        
                    'Si el string da nullo, es por que nadie esta insultando.
                    If nString = vbNullString Then
                        Call AddtoRichTextBox(frmMain.RecTxt, "Nadie te esta insultando. Las FotoDenuncias solo sirven para denunciar agravios.", 0, 200, 200, False, False, True)
                    Else 'Si no, enviamos.
                        Call AddtoRichTextBox(frmMain.RecTxt, "La FotoDenuncia fue sacada correctamente.", 0, 200, 200, False, False, True)
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

            If frmMain.MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
           
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

            If frmMain.MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
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
    
End Sub

Private Sub CMSG_Click()
    Call Audio.PlayInterface(SND_CLICK)

    If Not CharTieneClan And Not bCMSG Then
        Call AddtoRichTextBox(frmMain.RecTxt, "¡No perteneces a ningún clan!", 0, 200, 200, False, False, True)

    Else
        If PMSGimg Then Call imgPMSG_Click
        
        bCMSG = Not bCMSG

        If bCMSG Then
            
            '#If Classic = 0 Then
          '  CMSG.Picture = LoadPicture(App.path & "\resource\interface\main\ChatClanActivo.jpg")
           ' #Else
            imgButton(2).Picture = LoadPicture(App.path & "\resource\interface\main\CMSG.jpg")
          '  #End If
            Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu clan.", 0, 200, 200, False, False)
            HablaTemp = "/CMSG "
            
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de ser escuchado por tu clan. ", 0, 200, 200, False, False)
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
        Call AddtoRichTextBox(frmMain.RecTxt, "Todo lo que digas sera escuchado por tu party. ", 255, 200, 200, False, False)
        HablaTemp = "/PMSG "
    Else 'si ya estaba apretado lo desactivamos
        PMSGimg = False 'desactivamos el boton
        imgPMSG.Picture = Nothing
        imgButton(3).Picture = Nothing
        Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de ser escuchado por tu party. ", 255, 200, 200, False, False)
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
    frmPanelGm.Show , frmMain
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

Private Sub imgOpciones_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call frmOpciones.Show(vbModeless, frmMain)
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
      
    Dim FilePath As String
    FilePath = App.path & "\SCREENSHOTS\Prueba.jpg"
    'FilePath = "c:\pantalla.bmp"
    
    If FileExist(FilePath, vbArchive) Then
        Kill FilePath
    End If
    
    Clipboard.Clear
      
    ' Manda la pulsación de teclas para capturar la imagen de la pantalla
    Call keybd_event(44, 2, 0, 0)
      
    DoEvents
    
    ' Si el formato del clipboard es un bitmap
    If Clipboard.GetFormat(vbCFBitmap) Then
        'Guardamos la imagen en disco
        SavePicture Clipboard.GetData(vbCFBitmap), FilePath
    End If
    
   
End Sub

Private Sub MainViewPic_Click()
    
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
                    Call CountPacketIterations(packetControl(ClientPacketID.LeftClick), 150)
                    Call WriteLeftClick(tX, tY)
                Else
                      
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        Call RestoreLastCursor(IDC_CROSS)
                        Exit Sub

                    End If
                          
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                    
                        If MainTimer.Check(TimersIndex.AttackSpell, False) Then
                            If MainTimer.Check(TimersIndex.CastAttack, False) Then
                                If Not MainTimer.Check(TimersIndex.Arrows) Then
                                    frmMain.MousePointer = vbDefault
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

                    'Splitted because VB isn't lazy!
                    If UsingSkill = Robar Or UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Then

                        If MainTimer.Check(TimersIndex.CastSpell) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Call RestoreLastCursor(IDC_CROSS)
                            Exit Sub

                        End If
                    End If

                    frmMain.MousePointer = vbDefault
                    Call RestoreLastCursor(IDC_CROSS)
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If

            Else

                ' Descastea
                If UsingSkill = Magia Or UsingSkill = Proyectiles Then
                    frmMain.MousePointer = vbDefault
                    Call RestoreLastCursor(IDC_CROSS)
                    UsingSkill = 0
                Else

                    If Not Comerciando Then
                        Call WriteRightClick(tX, tY)
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
            
                
            FrmCantidad.Show , frmMain
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
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub


Private Sub coord_click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If CoordBloqued Then
        CoordBloqued = False
    Else
        CoordBloqued = True
    End If
    
End Sub

Private Sub coord_dblclick()
    Call Audio.PlayInterface(SND_CLICK)
    Call frmMapa.Show(vbModeless, frmMain)

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
                    
                    
                    If Not Comerciando Then FrmCantidad.Show , frmMain
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
    'If (GetSystemTime() - Intervalos(SecondaryClick).ModifyTime) <= 200 Then Exit Sub
          
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
        If Inventario.ObjType(ItemIndex) <> eOBJType.otBarcos Then
            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If

        End If

        Call WriteUseItem(ItemIndex, SecondaryClick, Value)
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




Private Sub Timer1_Timer()
    LoopInterval
    '  If frmMain.Visible Then RandomMove
    
    Static A As Long

   ' A = A + 1
    
   ' If A = 40 Then
       ' A = 0
    'End If
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
        Call SelectedSpamMessage
        Minutes = 0
        
    End If
    
End Sub



Private Sub tmrBlink_Timer()

    Dim A As Long
    
    For A = Me.lblFuerza.LBound To Me.lblFuerza.UBound
        If bLastBrightBlink Then
            frmMain.lblFuerza(A).ForeColor = getStrenghtColor(UserFuerza)
            frmMain.lblAgilidad(A).ForeColor = getDexterityColor(UserAgilidad)
        Else
            frmMain.lblFuerza(A).ForeColor = vbWhite
            frmMain.lblAgilidad(A).ForeColor = vbWhite
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

    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
          
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
      
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()

    If hlst.List(hlst.ListIndex) <> "(None)" Then

        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With

        Else
            If ClientSetup.bConfig(eSetupMods.SETUP_BOTONLANZAR) = 1 Then
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then Exit Sub
            End If
            
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
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

    If Not Inventario.uMoveItem Then
        picInv.MousePointer = vbDefault
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
        Call WriteDoubleClick(tX, tY)
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
    Call WriteSetPanelClient(Panel, 255, Pt.X, Pt.Y)
    LastPanel = Panel
    'End If

    ' Activo controles de inventario
    picInv.visible = True
    
    #If ModoBig = 1 Then
        frmMain.GldLbl(0).visible = False
        frmMain.GldLbl(1).visible = True
    #End If
    ' Desactivo controles de hechizo
    
    #If ModoBig = 0 Then
    picHechiz(0).visible = False
    #Else
    picHechiz(1).visible = False
    
    #End If
    
    CmdLanzar.visible = False
    imgInfo.visible = False
    
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

Private Sub Label7_Click()
    
    Call Audio.PlayInterface(SND_CLICK)
    
    'If Not MainTimer.Check(TimersIndex.Packet250) Then Exit Sub
    
    ' #If Classic = 0 Then
    '  InvEqu.Picture = LoadPicture(DirInterface & "main\Spell.jpg")
    ' GldLbl.visible = False
    '#Else
    
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
              
    Call WriteSetPanelClient(Panel, TempInv, Pt.X, Pt.Y)
    LastPanel = Panel

    'End If
          
    ' Activo controles de hechizos
    #If ModoBig = 0 Then
    picHechiz(0).visible = True
    #Else
        picHechiz(1).visible = True
    #End If
    
    CmdLanzar.visible = True
    imgInfo.visible = True
    
    #If ModoBig = 1 Then

        For A = frmMain.GldLbl.LBound To frmMain.GldLbl.UBound
            frmMain.GldLbl(A).visible = False
        Next A

    #End If
    
    ' Desactivo controles de inventario
    picInv.visible = False
    
    imgButton(6).visible = False
    imgButton(7).visible = False
    imgButton(8).visible = False
    'imgInvScrollUp.Visible = False
    'imgInvScrollDown.Visible = False

End Sub

Private Sub Label7_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Inventario.uMoveItem = False
    Inventario.sMoveItem = False
    
    SetHand
End Sub

Private Sub picInv_DblClick()
    
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
            ObjData(ObjIndex).ObjType = otFlechas) Then
            
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

Private Sub picInv_Click()
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
   ' Inventario.DrawInventory
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
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
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
                    Call Get_Image(DirGraficos, CStr(3057), data, False)
                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1)
                    frmMain.ImageList1.ListImages.Add , "g3057", Picture:=bmpData
                    Position = frmMain.ImageList1.ListImages.count
                    Set bmpData = Nothing
                    
                End If
                  
                Inventario.uMoveItem = True
                  
                Set picInv.MouseIcon = frmMain.ImageList1.ListImages(Position).ExtractIcon
                frmMain.picInv.MousePointer = vbCustom

                Exit Sub

            End If
        End If
    End If

End Sub

Private Function Search_GhID(ByVal gh As Integer) As Integer

    Dim I As Long

    For I = 1 To frmMain.ImageList1.ListImages.count

        If frmMain.ImageList1.ListImages(I).Key = "g" & CStr(gh) Then
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
        frmMain.SendTxt.SetFocus
        
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





Private Sub tUpdateMS_Timer()
    Call ParseUserCommand("/PING")
End Sub

Private Sub UpdateMapa_Timer()
    Call wGL_Graphic.Capture(frmMain.MiniMapa.hWnd, MiniMap_FilePath & UserMap & ".png")
    UpdateMapa.Enabled = False
End Sub


Private Sub macrotrabajo_Timer()

    If Inventario.SelectedItem = 0 Then
        DesactivarMacroTrabajo

        Exit Sub

    End If
          
    If (UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
          
    Call UsarItem(0)
    Inventario.DrawInventory
End Sub


Public Sub ActivarMacroTrabajo()
    MacroTrabajo.Interval = IntervaloUserPuedeTrabajar
    MacroTrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Empiezas a trabajar", 0, 200, 200, False, False, True)

End Sub

Public Sub DesactivarMacroTrabajo()

    MacroTrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de trabajar", 0, 200, 200, False, False, True)
       
End Sub


' Lista Gráfica de Hechizos
Private Sub picHechiz_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picHechiz(Index).ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then Y = Int(picHechiz(Index).ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
If X < picHechiz(Index).ScaleWidth - 10 Then
    hlst.ListIndex = Int(Y / hlst.Pixel_Alto) + hlst.Scroll
    hlst.DownBarrita = 0

Else
    hlst.DownBarrita = Y - hlst.Scroll * (picHechiz(Index).ScaleHeight - hlst.BarraHeight) / (hlst.ListCount - hlst.VisibleCount)
End If
End Sub

Private Sub picHechiz_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseShift = Shift
If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picHechiz(Index).ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1 Then yy = Int(picHechiz(Index).ScaleHeight / hlst.Pixel_Alto) * hlst.Pixel_Alto - 1
    If hlst.DownBarrita > 0 Then
        hlst.Scroll = (Y - hlst.DownBarrita) * (hlst.ListCount - hlst.VisibleCount) / (picHechiz(Index).ScaleHeight - hlst.BarraHeight)
    Else
        hlst.ListIndex = Int(yy / hlst.Pixel_Alto) + hlst.Scroll

        If ScrollArrastrar = 0 Then
            If (Y < yy) Then hlst.Scroll = hlst.Scroll - 1
            If (Y > yy) Then hlst.Scroll = hlst.Scroll + 1
        End If
    End If
ElseIf Button = 0 Then
    hlst.ShowBarrita = X > picHechiz(Index).ScaleWidth - hlst.BarraWidth * 2
End If
End Sub

Private Sub picHechiz_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
hlst.DownBarrita = 0
End Sub
