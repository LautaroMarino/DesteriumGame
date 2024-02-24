VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmShop 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Shop Argentum"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmShop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicTier 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5880
      Index           =   0
      Left            =   300
      Picture         =   "FrmShop.frx":000C
      ScaleHeight     =   392
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3675
      Begin RichTextLib.RichTextBox RecTxt 
         Height          =   4425
         Index           =   0
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Seleccionar Tier"
         Top             =   720
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   7805
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FrmShop.frx":67E1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgTierDsp 
         Height          =   480
         Index           =   0
         Left            =   960
         Picture         =   "FrmShop.frx":685E
         Top             =   5220
         Width           =   480
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TIER 1"
         BeginProperty Font 
            Name            =   "Booter - Five Zero"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   480
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   3645
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         DataField       =   "+"
         BeginProperty Font 
            Name            =   "Booter - Five Zero"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Top             =   5280
         Width           =   540
      End
   End
   Begin VB.PictureBox PicTier 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5880
      Index           =   2
      Left            =   7920
      Picture         =   "FrmShop.frx":9DE2
      ScaleHeight     =   392
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3675
      Begin RichTextLib.RichTextBox RecTxt 
         Height          =   4425
         Index           =   2
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Seleccionar Tier"
         Top             =   720
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   7805
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FrmShop.frx":105B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         DataField       =   "+"
         BeginProperty Font 
            Name            =   "Booter - Five Zero"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Index           =   2
         Left            =   1440
         TabIndex        =   20
         Top             =   5280
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIER 1"
         BeginProperty Font 
            Name            =   "Booter - Five Zero"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   480
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   60
         Width           =   3645
      End
      Begin VB.Image imgTierDsp 
         Height          =   480
         Index           =   2
         Left            =   960
         Picture         =   "FrmShop.frx":10634
         Top             =   5220
         Width           =   480
      End
   End
   Begin VB.TextBox txtBank 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   2160
      TabIndex        =   9
      Top             =   5040
      Width           =   7680
   End
   Begin VB.TextBox txtEmail 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   2160
      TabIndex        =   8
      Top             =   4110
      Width           =   7680
   End
   Begin VB.Timer tUpdate 
      Interval        =   150
      Left            =   11040
      Top             =   360
   End
   Begin VB.ComboBox cmbPromotion 
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5880
      Width           =   7695
   End
   Begin VB.PictureBox PicTier 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5880
      Index           =   1
      Left            =   4110
      Picture         =   "FrmShop.frx":13BB8
      ScaleHeight     =   392
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3675
      Begin RichTextLib.RichTextBox RecTxt 
         Height          =   4425
         Index           =   1
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Seleccionar Tier"
         Top             =   720
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   7805
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FrmShop.frx":1A38D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         DataField       =   "+"
         BeginProperty Font 
            Name            =   "Booter - Five Zero"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   420
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   5280
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIER 1"
         BeginProperty Font 
            Name            =   "Booter - Five Zero"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   480
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   60
         Width           =   3645
      End
      Begin VB.Image imgTierDsp 
         Height          =   480
         Index           =   1
         Left            =   960
         Picture         =   "FrmShop.frx":1A40A
         Top             =   5220
         Width           =   480
      End
   End
   Begin VB.PictureBox PicDraw 
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
      Height          =   5910
      Left            =   3150
      MousePointer    =   99  'Custom
      ScaleHeight     =   394
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2145
      Width           =   8415
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   7
         Left            =   6240
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   6
         Left            =   6240
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   5
         Left            =   4200
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   4
         Left            =   4200
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   3
         Left            =   2160
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   2
         Left            =   2160
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   1
         Left            =   120
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgPagination 
         Height          =   495
         Index           =   1
         Left            =   4320
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image imgPagination 
         Height          =   495
         Index           =   0
         Left            =   3720
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image imgItem 
         Height          =   2295
         Index           =   0
         Left            =   210
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   4
      Left            =   675
      Picture         =   "FrmShop.frx":1D98E
      Top             =   7780
      Width           =   2010
   End
   Begin VB.Label lblPriceDSP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   6720
      TabIndex        =   23
      Top             =   8565
      Width           =   120
   End
   Begin VB.Label lblPriceGLD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   6720
      TabIndex        =   22
      Top             =   8160
      Width           =   120
   End
   Begin VB.Image imgValueDSP 
      Height          =   360
      Left            =   6360
      Picture         =   "FrmShop.frx":22509
      Top             =   8490
      Width           =   2010
   End
   Begin VB.Image imgValueGLD 
      Height          =   360
      Left            =   6360
      Picture         =   "FrmShop.frx":24F13
      Top             =   8100
      Width           =   2010
   End
   Begin VB.Label lblAlias 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lemondsp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   2880
      TabIndex        =   21
      Top             =   2580
      Width           =   2820
   End
   Begin VB.Image imgNoUsage 
      Height          =   465
      Left            =   3480
      Picture         =   "FrmShop.frx":278AE
      Top             =   8160
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image imgUsage 
      Height          =   465
      Left            =   3480
      Picture         =   "FrmShop.frx":2B984
      Top             =   8160
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   6
      Left            =   675
      Picture         =   "FrmShop.frx":3114B
      Top             =   6855
      Width           =   2010
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   5
      Left            =   675
      Picture         =   "FrmShop.frx":35908
      Top             =   5940
      Width           =   2010
   End
   Begin VB.Image imgPaypal 
      Height          =   735
      Left            =   6480
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblCVU 
      BackStyle       =   0  'Transparent
      Caption         =   "0000007900204083203732"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   2880
      TabIndex        =   12
      Top             =   2880
      Width           =   2820
   End
   Begin VB.Image imgCopy 
      Height          =   255
      Left            =   5640
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblCantDSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   435
      Left            =   2175
      TabIndex        =   11
      Top             =   6840
      Width           =   7680
   End
   Begin VB.Image imgMoney1 
      Height          =   255
      Left            =   720
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image imgMoney2 
      Height          =   255
      Left            =   720
      Top             =   1350
      Width           =   2175
   End
   Begin VB.Image imgPoints 
      Height          =   255
      Left            =   720
      Top             =   1620
      Width           =   2055
   End
   Begin VB.Image imgAdd 
      Height          =   465
      Left            =   9600
      Picture         =   "FrmShop.frx":3A877
      Top             =   8160
      Width           =   1890
   End
   Begin VB.Label lblPoints 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   1080
      TabIndex        =   3
      Top             =   1635
      Width           =   1065
   End
   Begin VB.Image imgDsp 
      Height          =   825
      Left            =   7440
      Picture         =   "FrmShop.frx":3EBDB
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   11640
      Picture         =   "FrmShop.frx":43CAD
      Top             =   15
      Width           =   330
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   3
      Left            =   675
      Picture         =   "FrmShop.frx":44D5F
      Top             =   5025
      Width           =   2010
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   2
      Left            =   675
      Picture         =   "FrmShop.frx":4B3B4
      Top             =   4110
      Width           =   2010
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   1
      Left            =   675
      Picture         =   "FrmShop.frx":514F4
      Top             =   3195
      Width           =   2010
   End
   Begin VB.Image imgMenu 
      Height          =   825
      Index           =   0
      Left            =   675
      Picture         =   "FrmShop.frx":55C10
      Top             =   2280
      Width           =   2010
   End
   Begin VB.Image imgTier 
      Height          =   825
      Left            =   5400
      Picture         =   "FrmShop.frx":5C053
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Image imgGeneral 
      Height          =   825
      Left            =   3360
      Picture         =   "FrmShop.frx":6073B
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Label lblDSP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1335
      Width           =   1605
   End
   Begin VB.Label lblGld 
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   1065
      Width           =   1605
   End
End
Attribute VB_Name = "FrmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ePanel

    eTienda = 1
    eTier = 2
    eCargar = 3
    eChars = 4
    eAvatars = 5
End Enum


Private Enum ePRICE
    eGLD = 0
    eDSP = 1
End Enum

Private SelectedPrice As ePRICE




Public Pagination As Integer


Private Panel As ePanel
Public TierSelected As Byte

' Consola Transparente
Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = -20

Private Const WS_EX_LAYERED = &H80000


Public Lastcopy As Integer
Public MeditationSelected As Integer

Private Const WS_EX_TRANSPARENT As Long = &H20&

Private Sub cmbPromotion_Click()

    Select Case cmbPromotion.ListIndex
    
        Case 0
            lblCantDSP.Caption = "250 DSP"
        Case 1
            lblCantDSP.Caption = "500 DSP"
        Case 2
            lblCantDSP.Caption = "1.000 DSP"
        Case 3
            lblCantDSP.Caption = "2.000 DSP"
        Case 4
            lblCantDSP.Caption = "4.000 DSP"
        Case 5
            lblCantDSP.Caption = "8.000 DSP"
    End Select
End Sub

Private Sub Form_Load()
        
    g_Captions(eCaption.e_Shop) = wGL_Graphic.Create_Device_From_Display(PicDraw.hWnd, PicDraw.ScaleWidth, PicDraw.ScaleHeight)
    lblGld.Caption = PonerPuntos(Account.Gld)
    lblDSP.Caption = PonerPuntos(Account.Eldhir)
    lblPoints.Caption = PonerPuntos(UserPoints)
    
    ReDim ShopCopy(1 To ShopLast) As tShop
    
    MirandoShop = True
    Dim A As Long
    
    For A = 0 To 2
        Call SetWindowLong(RecTxt(A).hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
        SetTier A
    Next A

    SelectedPanel ePanel.eTienda
    
    imgPoints.ToolTipText = "Puntos de Torneo de tu Personaje"
    imgMoney1.ToolTipText = "Monedas de Oro de tu Cuenta"
    imgMoney2.ToolTipText = "Desterium Points de tu Cuenta"
    
    ' Cargamos las promociones de DSP
    cmbPromotion.AddItem "AR$1.000-10OFF DEPOSITÁ AR$900!"
    cmbPromotion.AddItem "AR$1.750-20OFF DEPOSITÁ AR$1.400!"
    cmbPromotion.AddItem "AR$2.500-25OFF DEPOSITÁ AR$1.875!"
    cmbPromotion.AddItem "AR$4.500-30OFF DEPOSITÁ AR$3.150!"
    cmbPromotion.AddItem "AR$7.500-45OFF DEPOSITÁ AR$4.125!"
    cmbPromotion.AddItem "AR$15.000-50OFF DEPOSITÁ AR$7.500!"
    cmbPromotion.ListIndex = 1


    txtEmail.Text = Account.Email
    
    
    lblCVU.Caption = CVU
    lblAlias.Caption = Alias
    
    AjustarClic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    MirandoShop = False
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.e_Shop))

End Sub
Private Sub imgAdd_Click()
    
    If Panel = ePanel.eTienda Then
        If SelectedTienda > 0 Then
            If ShopCopy(SelectedTienda).ID > 0 Then
                
                If MsgBox("¿Estás seguro que deseas comprar: " & Shop(ShopCopy(SelectedTienda).ID).Name & "?", vbYesNo) = vbYes Then
                    ComprarItem

                End If

            End If

        End If
    
    ElseIf Panel = ePanel.eTier Then

        If TierSelected > 0 Then
            If Account.Eldhir < Val(ReadField(1, lblPrice(TierSelected - 1), Asc(" "))) Then
                Call MsgBox("No tienes suficientes DSP.")
                Exit Sub

            End If
            
            If MsgBox("¿Estás seguro que deseas comprar: Tier n°" & TierSelected & "? Deberás dar " & lblPrice(TierSelected - 1), vbYesNo) = vbYes Then
                WriteConfirmTier TierSelected

            End If

        End If

    ElseIf Panel = ePanel.eChars Then

        If SelectedTienda > 0 And SelectedTienda <= UBound(ShopChars) Then
            If ShopChars(SelectedTienda).Dsp > 0 Then
                If MsgBox("¿Estás seguro que deseas comprar el personaje: " & ShopChars(SelectedTienda).Name & "?", vbYesNo) = vbYes Then
                    ComprarChar

                End If

            End If

        End If
    
    ElseIf Panel = ePanel.eCargar Then

        If MsgBox("¿Has abonado la cantidad de: " & cmbPromotion.List(cmbPromotion.ListIndex) & "? En ese caso confirma y recibirás los DSP correspondientes.", vbYesNo) = vbYes Then
            Call CargarDSP

        End If

    End If

End Sub

Private Sub ComprarChar()

    If ShopChars(SelectedTienda).Dsp > Account.Eldhir Then
        Call MsgBox("¡Parece que no tienes los DSP en tu cuenta para comprar el personaje!")
        Exit Sub

    End If

    Call WriteConfirmChar(SelectedTienda)
End Sub
Public Function ApplyDiscount(ByVal Tier As Integer, ByVal Price As Long)
        '<EhHeader>
        On Error GoTo ApplyDiscount_Err
        '</EhHeader>
100     Select Case Tier
    
            Case 0
102             ApplyDiscount = Price
104         Case 1
106             ApplyDiscount = Price - (Price * 0.05)
108         Case 2
110             ApplyDiscount = Price - (Price * 0.07)
112         Case 3
114             ApplyDiscount = Price - (Price * 0.1)
        End Select
        '<EhFooter>
        Exit Function

ApplyDiscount_Err:
        LogError err.Description & vbCrLf & _
               "in ServidorArgentum.mShop.ApplyDiscount " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Private Sub ComprarItem()
    
    With Shop(ShopCopy(SelectedTienda).ID)
        If ApplyDiscount(Account.Premium, .Gld) > Account.Gld And SelectedPrice = ePRICE.eGLD Then
            Call MsgBox("¡Parece que no tienes las Monedas de Oro en tu cuenta!")
            Exit Sub
        End If
    
        If .ObjIndex <> 880 Then
            If ApplyDiscount(Account.Premium, .Dsp) > Account.Eldhir And SelectedPrice = ePRICE.eDSP Then
                Call MsgBox("¡Parece que no tienes los DSP en tu cuenta!")
                Exit Sub
            End If
        End If
        
        If .Points > UserPoints Then
            Call MsgBox("¡Tus puntos de Torneo no son suficientes para realizar el Canje!")
            Exit Sub
        End If
        
    End With
    
    Call WriteConfirmItem(ShopCopy(SelectedTienda).ID, SelectedPrice)
End Sub
Private Sub CargarDSP()

    If txtEmail.Text = vbNullString Then
        Call MsgBox("Ingresa un Email el cual recibirá los DSP. ¡Puede ser el de un amigo!")
        Exit Sub

    End If
    
    If Not CheckMailString(txtEmail.Text) Then
        Call MsgBox("Corrobora que el Email ingresado sea válido.")
        Exit Sub

    End If

    If txtBank.Text = vbNullString Then
         Call MsgBox("Debes proporcionar un nombre bancario, una cuenta de Mercado Pago o bien algo que haga referencia que el dinero que nos llegará, proviene de tu cuenta.")
        Exit Sub
    End If
    
    Call WriteConfirmTransaccion(txtEmail.Text, cmbPromotion.ListIndex, txtBank.Text)
End Sub

Private Sub imgCopy_Click()
    Clipboard.Clear
    
    'Copiar todo el contenido de la caja de texto
    Clipboard.SetText lblCVU.Caption
    
    Call MsgBox("CVU copiado exitosamente")
End Sub

Private Sub imgDsp_Click()
    Call Audio.PlayInterface(SND_CLICK)
      
    SelectedPanel ePanel.eCargar

End Sub

Private Sub imgGeneral_Click()
    Call Audio.PlayInterface(SND_CLICK)
      
    
    SelectedPanel ePanel.eTienda
      
End Sub

Private Sub imgPagination_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0 ' Página anterior
            If (Pagination - 8) < 0 Then Exit Sub
            Pagination = Pagination - 8
        
        Case 1 ' Página siguiente
            If Panel = ePanel.eTienda Then
                If (Pagination + 8) >= Lastcopy Then Exit Sub
            Else
                 If (Pagination + 8) >= ShopCharLast Then Exit Sub
            End If
            Pagination = Pagination + 8
    End Select
    
    SelectedTienda = 0
End Sub

Private Sub imgPaypal_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call MsgBox("Para realizar pagos vía Paypal deberás contactarte de forma privada con nuestros administradores")
End Sub

Private Sub imgTier_Click()
    Call Audio.PlayInterface(SND_CLICK)
      
    SelectedPanel ePanel.eTier

End Sub

' Copias de Clics
Private Sub lblCargar_Click()
    imgDsp_Click

End Sub

Private Sub lblGeneral_Click()
    imgGeneral_Click

End Sub

Private Sub lblTiers_Click()
    imgTier_Click

End Sub

' Fin Copias de clics
Private Sub SelectedPanel(ByVal cPanel As ePanel)

    Dim A As Long
    
    Panel = cPanel
    
    ' Disable All buttons
    imgGeneral.Picture = LoadPicture(DirInterface & "\shop\general.jpg")
    imgDsp.Picture = LoadPicture(DirInterface & "\shop\cargar.jpg")
    imgTier.Picture = LoadPicture(DirInterface & "\shop\tiers.jpg")
    
    lblAlias.visible = False
    PicDraw.visible = False
    
    For A = imgMenu.LBound To imgMenu.UBound
        imgMenu(A).visible = False
    Next A
    
    For A = PicTier.LBound To PicTier.UBound
        PicTier(A).visible = False
    Next A
    
    lblGld.visible = False
    lblDSP.visible = False
    lblPoints.visible = False
    imgGeneral.visible = False
    imgTier.visible = False
    imgDsp.visible = False
    
    txtEmail.visible = False
    txtBank.visible = False
    cmbPromotion.visible = False
    lblCantDSP.visible = False
    lblCVU.visible = False
    imgCopy.visible = False
    
    lblPriceDSP.visible = False
    lblPriceGLD.visible = False
    imgValueGLD.visible = False
    imgValueDSP.visible = False
    
    Select Case cPanel
    
        Case ePanel.eTienda
            Me.Picture = LoadPicture(DirInterface & "\shop\shop.jpg")
            imgGeneral.Picture = LoadPicture(DirInterface & "\shop\general_hover.jpg")
            
            For A = imgMenu.LBound To imgMenu.UBound
                imgMenu(A).visible = True
            Next A
            
            PicDraw.visible = True
            lblGld.visible = True
            lblDSP.visible = True
            lblPoints.visible = True
            imgGeneral.visible = True
            imgTier.visible = True
            imgDsp.visible = True
            lblPriceDSP.visible = True
            lblPriceGLD.visible = True
            imgValueGLD.visible = True
            imgValueDSP.visible = True
    
        Case ePanel.eTier
            Me.Picture = LoadPicture(DirInterface & "\shop\paneltiers.jpg")
            imgTier.Picture = LoadPicture(DirInterface & "\shop\tiers_hover.jpg")
            
            For A = PicTier.LBound To PicTier.UBound
                PicTier(A).visible = True
            Next A
            
            lblGld.visible = True
            lblDSP.visible = True
            lblPoints.visible = True
            imgGeneral.visible = True
            imgTier.visible = True
            imgDsp.visible = True
            
        Case ePanel.eCargar
            Me.Picture = LoadPicture(DirInterface & "\shop\panelconfirm.jpg")
            Me.imgDsp.Picture = LoadPicture(DirInterface & "\shop\cargar_hover.jpg")
            
            txtEmail.visible = True
            txtBank.visible = True
            cmbPromotion.visible = True
            lblCantDSP.visible = True
            imgCopy.visible = True
            lblCVU.visible = True
            lblAlias.visible = True
    End Select

End Sub

' ############ SOLAPA DE ITEMS SHOP ##########################
Private Sub imgItem_Click(Index As Integer)
   Call Audio.PlayInterface(SND_CLICK)
    SelectedTienda = (Index + 1) + Pagination
    
    Call UpdateMeditationLearn
    Call DetectPrize(SelectedTienda)
End Sub

Private Sub imgMenu_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    

    
    imgNoUsage.visible = False
    imgUsage.visible = False
    imgAdd.visible = True
     
    imgMenu(0).Picture = LoadPicture(DirInterface & "shop\avatars.jpg")
    imgMenu(1).Picture = LoadPicture(DirInterface & "shop\skins.jpg")
    imgMenu(2).Picture = LoadPicture(DirInterface & "shop\scrolls.jpg")
    imgMenu(3).Picture = LoadPicture(DirInterface & "shop\meditars.jpg")
    imgMenu(4).Picture = LoadPicture(DirInterface & "shop\canje.jpg")
    imgMenu(5).Picture = LoadPicture(DirInterface & "shop\chars.jpg")
    imgMenu(6).Picture = LoadPicture(DirInterface & "shop\other.jpg")
    
    ReDim ShopCopy(1 To ShopLast) As tShop
    
    Panel = ePanel.eTienda
    
    Select Case Index
    
        Case 0 ' AVATARES
            Call MsgBox("Próximamente.")
            Exit Sub
            imgMenu(0).Picture = LoadPicture(DirInterface & "shop\avatars_hover.jpg")
            Panel = ePanel.eAvatars
            Shop_OrderByAvatars
            
        Case 1
            imgMenu(1).Picture = LoadPicture(DirInterface & "shop\skins_hover.jpg")
            Call ParseUserCommand("/SKINS")
            Unload Me
            
        Case 2 ' SCROLLS
            imgMenu(2).Picture = LoadPicture(DirInterface & "shop\scrolls_hover.jpg")
            Shop_OrdenBy eOBJType.oteffect

        Case 3 ' MEDITARES
            Shop_OrderByMeditares
            Call InitMeditares
            
            imgMenu(3).Picture = LoadPicture(DirInterface & "shop\meditars_hover.jpg")
    
        Case 4 ' CANJEAR
            'Call MsgBox("Próximamente.")
            'Exit Sub
            imgMenu(4).Picture = LoadPicture(DirInterface & "shop\canje_hover.jpg")
            Shop_OrdenBy 0, 1
            
        Case 5 ' Personajes
            Panel = ePanel.eChars
            Call WriteRequiredShopChars
            imgMenu(5).Picture = LoadPicture(DirInterface & "shop\chars_hover.jpg")
            
        Case 6 ' Otros
            imgMenu(6).Picture = LoadPicture(DirInterface & "shop\other_hover.jpg")
            Shop_OrderBy_Other
    End Select
    
    Pagination = 0
    SelectedTienda = 0
    AjustarClic
End Sub
Private Sub InitMeditares()

    Dim A As Long
    
    For A = LBound(ShopCopy) To UBound(ShopCopy)
        With ShopCopy(A)
            If .ObjIndex = 9998 Then
             '   InitGrh .FX, FxData(.ObjAmount).Animacion, , True
            End If
            
        End With
    Next A

End Sub
Private Sub lblMenu_Click(Index As Integer)
    imgMenu_Click Index

End Sub

' Solapa de OTROS
Private Function Shop_Is_ObjIndex_Other(ByVal ObjIndex As Integer) As Boolean
    
    Shop_Is_ObjIndex_Other = False
    
    If ObjIndex = ACTA_NACIMIENTO Then Shop_Is_ObjIndex_Other = True
    If ObjIndex = ESCRITURAS_CLAN Then Shop_Is_ObjIndex_Other = True
    If ObjIndex = PERLA_FORTUNA_1 Then Shop_Is_ObjIndex_Other = True
    If ObjIndex = PERLA_FORTUNA_2 Then Shop_Is_ObjIndex_Other = True
    If ObjIndex = FRAGMENTO_HIELO Then Shop_Is_ObjIndex_Other = True
    If ObjIndex = MANUSCRITO_1 Then Shop_Is_ObjIndex_Other = True
    If ObjIndex = MANUSCRITO_2 Then Shop_Is_ObjIndex_Other = True
End Function
Private Sub Shop_OrderBy_Other()

    Dim A    As Long

    Dim Last As Byte
    
    For A = 1 To ShopLast

        With Shop(A)

            If Shop_Is_ObjIndex_Other(.ObjIndex) Then
                Last = Last + 1
                    
                ShopCopy(Last) = Shop(A)
                ShopCopy(Last).ID = A

            End If
            
        End With
    
    Next A

    Lastcopy = Last
End Sub

Private Sub Shop_OrderByMeditares()

    Dim A    As Long

    Dim Last As Byte
    
    For A = 1 To ShopLast

        With Shop(A)

            If .ObjIndex = 9998 Then
            
                Last = Last + 1
                    
                ShopCopy(Last) = Shop(A)
                ShopCopy(Last).ID = A
                
                InitGrh ShopCopy(Last).fX, FxData(ShopCopy(Last).ObjAmount).Animacion, , True
            End If
            
        End With
    
    Next A

    Lastcopy = Last
End Sub

Private Sub Shop_OrderByAvatars()

    Dim A    As Long

    Dim Last As Byte
    
    For A = 1 To ShopLast

        With Shop(A)

            If .ObjIndex = 9997 Then
            
                Last = Last + 1
                    
                ShopCopy(Last) = Shop(A)
                ShopCopy(Last).ID = A
            End If
            
        End With
    
    Next A

    Lastcopy = Last
End Sub
 ' Ordena por ObjType
Private Sub Shop_OrdenBy(ByVal ObjType As Byte, Optional ByVal Points As Integer = 0)
    
    Dim A    As Long

    Dim Last As Byte
    
    For A = 1 To ShopLast

        With Shop(A)
            
            If Points > 0 Then
                If .Points > 0 Then
                    Last = Last + 1
                    
                    ShopCopy(Last) = Shop(A)
                    ShopCopy(Last).ID = A

                End If
            
            Else
                
                If .ObjIndex <> 9999 And .ObjIndex <> 9998 And .ObjIndex <> 9997 Then
                If ObjData(.ObjIndex).ObjType = ObjType Then
                    Last = Last + 1
                    
                    ShopCopy(Last) = Shop(A)
                    ShopCopy(Last).ID = A
                End If
                End If
            End If
            
        End With
    
    Next A
    
    
        Lastcopy = Last

End Sub

'# FIN SOLAPAS SHOP

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Panel = ePanel.eCargar Then
         SelectedPanel ePanel.eTienda
         Exit Sub
    End If
    
    MirandoShop = False
   If MsgBox("¿Estás seguro que deseas salir?", vbYesNo) = vbYes Then
        Unload Me

    End If
    MirandoShop = True
    
End Sub

' Draw

Private Sub Render_Tienda()

 Dim X          As Long, Y As Long, Map As Integer, C As Long, DescY As Integer

    Dim A          As Long

    Dim OffsetY    As Integer

    Dim x1         As Long

    Dim y1         As Long

    Dim ColourText As Long
    
    Dim InitialX   As Integer

    Dim InitialY   As Integer
    
    
    X = 15
    Y = 30

    For A = 1 To 8
        
        If A <= ShopLast Then

            With ShopCopy(A + Pagination)

                If .ObjIndex > 0 And .ObjIndex <> 9997 Then
                    Call Render_Obj(A + Pagination, X, Y)
                End If
            
            End With
        
        End If
        
        Y = Y + 85
        
        If A = 4 Then
            X = X + 280
            Y = 30
        
        End If
        
    Next A

End Sub
Private Sub Render_Obj(ByVal A As Long, ByVal X As Long, ByVal Y As Long)
    
    Dim C As Long, DescY As Long
    
    With ShopCopy(A)

        ' Efecto Seleccionado
        If SelectedTienda = (A) Then
            Call Draw_Texture_Graphic_Gui(94, X, Y, To_Depth(2), 250, 74, 0, 0, 250, 74, -1, 0, eTechnique.t_Alpha)
            'Call Draw_Texture_Graphic_Gui(61, X, Y, To_Depth(3), 250, 74, 0, 0, 250, 74, ARGB(255, 197, 0, 50), 0, eTechnique.t_Alpha)
        Else
            Call Draw_Texture_Graphic_Gui(81, X, Y, To_Depth(2), 250, 74, 0, 0, 250, 74, -1, 0, eTechnique.t_Alpha)

        End If
                
        ' Texto del Item
        'ObjData(.ObjIndex).
        Draw_Text f_Verdana, 13, X + 5, Y + 5, To_Depth(4), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, .Name, True, True
                
        ' Icono del Item + Cantidad
                    
        Select Case .ObjIndex
                        
            Case 9999 ' Es personaje. Se dibuja la Cabeza del PJ a la venta
                Call Draw_Grh(HeadData(.ObjAmount).Head(E_Heading.SOUTH), X + 8, Y + 23, To_Depth(3, , , 1), 1, 0)

            Case 9998 ' Es una meditacion
                Call Draw_Texture(GRH_MEDITACION, X + 10, Y + 32, To_Depth(4), 32, 32, -1, 0, eTechnique.t_Alpha)
                                
                If SelectedTienda = (A) Then
                    If CharList(UserCharIndex).Body.Walk(E_Heading.SOUTH).GrhIndex > 0 Then
                        Call Draw_Grh(CharList(UserCharIndex).Body.Walk(E_Heading.SOUTH), X + 15 + CharList(UserCharIndex).Body.BodyOffSet(E_Heading.SOUTH).X, Y + 40 + CharList(UserCharIndex).Body.BodyOffSet(E_Heading.SOUTH).Y, To_Depth(5, , , 4), 1, 1, 0, -1, , eTechnique.t_Alpha)
                        Call Draw_Grh(CharList(UserCharIndex).Head.Head(E_Heading.SOUTH), X + 15 + CharList(UserCharIndex).Body.HeadOffset.X, Y + 40 + CharList(UserCharIndex).Body.HeadOffset.Y, To_Depth(5, , , 4), 1, 0, , -1, , eTechnique.t_Alpha)

                    End If

                    Call Draw_Grh(.fX, X + 15 + FxData(.ObjAmount).OffsetX, Y + 40 + (FxData(.ObjAmount).OffsetY), To_Depth(5, , , 5), 1, 1, 1, ARGB(255, 255, 255, ClientSetup.bAlpha), , eTechnique.t_Alpha)

                End If
            
            
            Case Else
                Call Draw_Texture(ObjData(.ObjIndex).GrhIndex, X + 10, Y + 32, To_Depth(3), 32, 32, -1, 0, eTechnique.t_Alpha)
                Draw_Text f_Tahoma, 12, X + 25, Y + 55, To_Depth(4), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "(x" & .ObjAmount & ")", True, True

        End Select
                    
        ' PRECIO
        If .Gld > 0 Then
            Call Draw_Texture_Graphic_Gui(83, X + 230, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
            Draw_Text f_Verdana, 13, X + 230, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(.Gld), True, True
        ElseIf .Dsp > 0 Then
            Call Draw_Texture_Graphic_Gui(84, X + 230, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
            Draw_Text f_Verdana, 13, X + 230, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(.Dsp), True, True
        ElseIf .Points > 0 Then
            Call Draw_Texture_Graphic_Gui(91, X + 230, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
            Draw_Text f_Verdana, 13, X + 230, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(CLng(.Points)), True, True

        End If
                
        ' Desc
        For C = LBound(.Desc) To UBound(.Desc)
            Draw_Text f_Verdana, 13, X + 145, Y + 35 + DescY, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Desc(C), True, True
                 
            DescY = DescY + 15
        Next C
                
        DescY = 0
    
    End With

End Sub

Private Sub Render_Chars()

        Dim X          As Long, Y As Long, Map As Integer, C As Long, DescY As Integer

        Dim A          As Long

        Dim OffsetY    As Integer

        Dim x1         As Long

        Dim y1         As Long

        Dim ColourText As Long
    
        Dim InitialX   As Integer

        Dim InitialY   As Integer
                
108     X = 15
110     Y = 30

112     For A = 1 To 8
        
114         If A <= ShopCharLast Then

116             With ShopChars(A + Pagination)

                    ' Efecto Seleccionado
120                 If SelectedTienda = (A + Pagination) Then
122                     Call Draw_Texture_Graphic_Gui(94, X, Y, To_Depth(2), 250, 74, 0, 0, 250, 74, -1, 0, eTechnique.t_Alpha)
124                     'Call Draw_Texture_Graphic_Gui(61, X, Y, To_Depth(3), 250, 74, 0, 0, 250, 74, ARGB(255, 197, 0, 50), 0, eTechnique.t_Alpha)
                    Else
126                     Call Draw_Texture_Graphic_Gui(81, X, Y, To_Depth(2), 250, 74, 0, 0, 250, 74, -1, 0, eTechnique.t_Alpha)

                    End If
                        
118                 If .Hp > 0 Then
                
                        ' Texto del Item
128                     Draw_Text f_Verdana, 13, X + 5, Y + 5, To_Depth(4), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, UCase$(ListaClases(.Class)) & " " & ListaRazasShort(.Raze) & " LVL" & .Elv & " " & IIf(.Porc <> 0, "(" & .Porc & "%)", vbNullString), True, True
                
                        ' Icono del Item + Cantidad
                        Call Draw_Grh(HeadData(.Head).Head(E_Heading.SOUTH), X + 8, Y + 23, To_Depth(3, , , 1), 1, 0)
                       
146                     Call Draw_Texture_Graphic_Gui(84, X + 230, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
148                     Draw_Text f_Verdana, 13, X + 230, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(CLng(.Dsp)), True, True
150
            
                        Draw_Text f_Verdana, 13, X + 145, Y + 35, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "HP: " & .Hp & IIf(.Ups = 0, vbNullString, IIf(.Ups > 0, " +" & .Ups, .Ups)) & " MAN: " & PonerPuntos(CLng(.Man)), True, True
                        Draw_Text f_Verdana, 13, X + 145, Y + 50, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Sin Items. ¡¡Incluye ORO!!", True, True
                    
                    Else
                        ' Texto del Item
                        Call Draw_Texture_Graphic_Gui(84, X + 230, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
                        Draw_Text f_Verdana, 13, X + 230, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, "0", True, True
                        Draw_Text f_Verdana, 13, X + 5, Y + 5, To_Depth(4), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "<VACIO>", True, True
                        Draw_Text f_Verdana, 13, X + 145, Y + 35, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "-", True, True
                        Draw_Text f_Verdana, 13, X + 145, Y + 50, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "-", True, True

                    End If
            
                End With
        
            End If

166         Y = Y + 85
        
168         If A = 4 Then
170             X = X + 280
172             Y = 30
        
            End If

174     Next A

End Sub

Private Sub Render()
        '<EhHeader>
        On Error GoTo Render_Err
        '</EhHeader>
        
100     Call wGL_Graphic.Use_Device(g_Captions(eCaption.e_Shop))
102     Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmShop.PicDraw.ScaleWidth, FrmShop.PicDraw.ScaleHeight)
104     Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
        Dim X          As Long, Y As Long, Map As Integer, C As Long, DescY As Integer

        Dim A          As Long

        Dim OffsetY    As Integer

        Dim x1         As Long

        Dim y1         As Long

        Dim ColourText As Long
    
        Dim InitialX   As Integer

        Dim InitialY   As Integer

106     Call Draw_Texture_Graphic_Gui(82, 0, 0, To_Depth(1), 561, 394, 0, 0, 561, 394, -1, 0, eTechnique.t_Alpha)


         ' Flechas para Mover de Solapa a la siguiente
         Call Draw_Texture_Graphic_Gui(92, 250, 367, To_Depth(5), 20, 17, 0, 0, 20, 17, -1, 0, eTechnique.t_Alpha)
         Call Draw_Texture_Graphic_Gui(93, 290, 367, To_Depth(5), 20, 17, 0, 0, 20, 17, -1, 0, eTechnique.t_Alpha)
          Draw_Text f_Verdana, 15, 279, 367, To_Depth(5), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Int(Pagination / 8) + 1, True, True
         
        
        Select Case Panel
            
            Case ePanel.eAvatars
                Call Render_Avatars
                
            Case ePanel.eTienda
                Call Render_Tienda

            Case ePanel.eChars
                Call Render_Chars
                
        End Select
         
    
176     Call wGL_Graphic_Renderer.Flush
  
        '<EhFooter>
        Exit Sub

Render_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.FrmShop.Render " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub Render_Avatars()

     Dim X          As Long, Y As Long, Map As Integer, C As Long, DescY As Integer

        Dim A          As Long

        Dim OffsetY    As Integer

        Dim x1         As Long

        Dim y1         As Long

        Dim ColourText As Long
    
        Dim InitialX   As Integer

        Dim InitialY   As Integer
        
108     X = 12
110     Y = 30

112     For A = 1 To 8
        
114         If A <= ShopLast Then

116             With ShopCopy(A + Pagination)

118                 If .ObjIndex > 0 Then
            
                        ' Fondo
        
                        ' Efecto Seleccionado
120                     If SelectedTienda = (A + Pagination) Then
122                         Call Draw_Texture_Graphic_Gui(97, X, Y, To_Depth(2), 134, 155, 0, 0, 134, 155, -1, 0, eTechnique.t_Alpha)
                        Else
126                         Call Draw_Texture_Graphic_Gui(96, X, Y, To_Depth(2), 134, 155, 0, 0, 134, 155, -1, 0, eTechnique.t_Alpha)

                        End If
                        
                         
                          ' Recuadro + Avatar
                          Call Draw_Texture_Graphic_Gui(95, X + 8, Y + 30, To_Depth(3), 118, 118, 0, 0, 118, 118, -1, 0, eTechnique.t_Alpha)
                          Call Draw_Avatar(.ObjAmount, X + 16, Y + 38, To_Depth(4), 102, 102, 0, 0, 102, 102, -1, 0, eTechnique.t_Alpha)
                          
                        ' PRECIO
138                     If .Gld > 0 Then
140                         Call Draw_Texture_Graphic_Gui(83, X + 110, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
142                         Draw_Text f_Verdana, 13, X + 110, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(.Gld), True, True
144                     ElseIf .Dsp > 0 Then
146                         Call Draw_Texture_Graphic_Gui(84, X + 110, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
148                         Draw_Text f_Verdana, 13, X + 110, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(.Dsp), True, True
150                     ElseIf .Points > 0 Then
152                         Call Draw_Texture_Graphic_Gui(91, X + 110, Y + 3, To_Depth(3), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
154                         Draw_Text f_Verdana, 13, X + 110, Y + 5, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(CLng(.Points)), True, True

                        End If

                    End If
            
                End With
        
            End If
        
166         Y = Y + 165
        
168         If A Mod 2 = 0 Then
170             X = X + 135
172             Y = 30
        
            End If
        
174     Next A

End Sub
Private Sub imgUsage_Click()
    If SelectedTienda > 0 Then
        MeditationSelected = Val(ReadField(2, Shop(ShopCopy(SelectedTienda).ID).Name, Asc(" ")))
        
        Call WriteLearnMeditation(1, MeditationSelected)
        
    End If
End Sub
Private Sub imgNoUsage_Click()
    Call WriteLearnMeditation(1, 0)
End Sub

Private Sub DetectPrize(ByVal SelectedTienda As Integer)

    If SelectedTienda = 0 Then Exit Sub
    
    If Panel = ePanel.eChars Then

        If SelectedTienda > 0 And SelectedTienda <= UBound(ShopChars) Then
            
            ' Venta por DSP
            If ShopChars(SelectedTienda).Dsp > 0 Then
                ValuePrize_Selected eDSP
                Exit Sub

            End If
    
            ' Venta por ORO
            If ShopChars(SelectedTienda).Gld > 0 Then
                ValuePrize_Selected eGLD
                Exit Sub

            End If

        End If
            
    Else
        If ShopCopy(SelectedTienda).ID = 0 Then Exit Sub
        
        ' Prioriza la venta por DSP
        If Shop(ShopCopy(SelectedTienda).ID).Dsp > 0 Then
            ValuePrize_Selected eDSP
            Exit Sub

        End If
    
        ' Venta por ORO
        If Shop(ShopCopy(SelectedTienda).ID).Gld > 0 Then
            ValuePrize_Selected eGLD
            Exit Sub

        End If
    
    End If

End Sub

Private Sub ValuePrize_Selected(ByRef Tipo As ePRICE)

    If SelectedTienda = 0 Then Exit Sub
     
    Select Case Tipo
    
        Case ePRICE.eGLD
            
            If Panel = eTienda Then
                If Shop(ShopCopy(SelectedTienda).ID).Gld = 0 Then
                    Call MsgBox("¡El objeto no puede ser comprado por Monedas de Oro!")
                    Exit Sub

                End If
                
                lblPriceGLD.Caption = Shop(ShopCopy(SelectedTienda).ID).Gld
                
            ElseIf Panel = eChars Then

                If ShopChars(SelectedTienda).Gld = 0 Then
                    Call MsgBox("¡El Personaje no puede ser comprado por Monedas de Oro!")
                    Exit Sub

                End If
                
                 lblPriceGLD.Caption = ShopChars(SelectedTienda).Gld
                 
            End If
            
            imgValueGLD.Picture = LoadPicture(DirInterface & "\shop\gld_hover.jpg")
            imgValueDSP.Picture = LoadPicture(DirInterface & "\shop\dspvalue.jpg")
        
        Case ePRICE.eDSP

            If Panel = eTienda Then
                If Shop(ShopCopy(SelectedTienda).ID).Dsp = 0 Then
                    Call MsgBox("¡El objeto no puede ser comprado por Monedas DSP!")
                    Exit Sub

                End If
                
                lblPriceDSP.Caption = Shop(ShopCopy(SelectedTienda).ID).Dsp
                
            ElseIf Panel = eChars Then

                If ShopChars(SelectedTienda).Dsp = 0 Then
                    Call MsgBox("¡El Personaje no puede ser comprado por DSP!")
                    Exit Sub

                End If
                
                lblPriceDSP.Caption = ShopChars(SelectedTienda).Dsp
                
            End If
    
            imgValueDSP.Picture = LoadPicture(DirInterface & "\shop\dspvalue_hover.jpg")
            imgValueGLD.Picture = LoadPicture(DirInterface & "\shop\gld.jpg")

    End Select
    

    SelectedPrice = Tipo

End Sub

Private Sub imgValueGLD_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call ValuePrize_Selected(eGLD)
End Sub

Private Sub imgValueDSP_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call ValuePrize_Selected(eDSP)
End Sub


Private Sub lblPriceDSP_Click()
    imgValueDSP_Click
End Sub

Private Sub lblPriceGLD_Click()
    imgValueGLD_Click
End Sub

Private Sub tUpdate_Timer()
    If Not MirandoShop Then Exit Sub
    
    Render

End Sub

' SOLAPA TIERS

Private Sub SetTier(ByVal Tier As Byte)

    Dim Desc   As String

    Dim Tittle As String

    Dim Price  As String
    
    RecTxt(Tier).Text = vbNullString
    RecTxt(Tier).SelText = vbNullString
    TierSelected = Tier + 1
    Select Case (Tier + 1)
    
        Case 1
            Tittle = "Tier 1"
            
            Call AddtoRichTextBox(RecTxt(Tier), "ELECCION DE CARA:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Al momento de crear un nuevo personaje elige que diseño de cabeza tendrá." & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "DESCUENTO:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "¡5%OFF en TIENDAS y SHOP!" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /RESET:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Si eres inferior a nivel 29 podrás reiniciar tu personaje a los atributos iniciales y comenzar de nuevo sin borrar el personaje." & vbCrLf, 200, 200, 200, 0, 0)
        
        Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /HOGAR: ", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "60 Segundos. Regreso 50%OFF" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "* Todos los personajes de tu cuenta reciben los beneficios del Tier escogido", 2, 51, 223, 0, 1)
            Call AddtoRichTextBox(RecTxt(Tier), "* Suscripción MENSUAL", 200, 200, 200, 0, 1)
            
            Price = "100DSP"
        
        Case 2
            Tittle = "Tier 2"

            Call AddtoRichTextBox(RecTxt(Tier), "INCLUYE TIER 1", 250, 240, 0, 1, 1, False)
              
            Call AddtoRichTextBox(RecTxt(Tier), "DESCUENTO:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "¡7%OFF en TIENDAS y SHOP!" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /RESET:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Si eres inferior a nivel 34 podrás reiniciar tu personaje a los atributos iniciales y comenzar de nuevo sin borrar el personaje." & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "BOVEDA EXCLUSIVA:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Comparte una boveda con todos los personajes de tu cuenta para poder disponer de objetos en todos ellos." & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /HOGAR: ", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "30 Segundos. Regreso 50%OFF" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /DESC:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), " Elige una descripción para tu personaje que será visible cuando hagan clic sobre él." & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "* Si tu cuenta vence los objetos que conservas en tu boveda quedarán bloqueados hasta retomar la membresía PREMIUM de la misma.", 2, 51, 223, 0, 1)
            Call AddtoRichTextBox(RecTxt(Tier), "* Todos los personajes de tu cuenta reciben los beneficios del Tier escogido", 2, 51, 223, 0, 1)
            Call AddtoRichTextBox(RecTxt(Tier), "* Suscripción MENSUAL", 200, 200, 200, 0, 1)
            Price = "250DSP"

        Case 3
            Tittle = "Tier 3"

            Call AddtoRichTextBox(RecTxt(Tier), "INCLUYE TIER 2", 250, 240, 0, 1, 1, False)
              
            Call AddtoRichTextBox(RecTxt(Tier), "DESCUENTO:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "¡10%OFF en TIENDAS y SHOP!" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /RESET:", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Si eres inferior a nivel 39 podrás reiniciar tu personaje a los atributos iniciales y comenzar de nuevo sin borrar el personaje." & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "SUBASTAS: ", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Podrás subastar los objetos que hayas obtenidos y así tener una ganancia en Monedas de Oro que podrás disfrutar en tus personajes." & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "MERCADO: ", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "Permite publicar y ofrecer más de 1 personaje, vender una cuenta completa sin perder nuestro email. + SIN VALOR DE COSTO" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "COMANDO /HOGAR: ", 250, 240, 0, 1, 1, False)
            Call AddtoRichTextBox(RecTxt(Tier), "5 Segundos.  Regreso GRATIS" & vbCrLf, 200, 200, 200, 0, 0)
            
            Call AddtoRichTextBox(RecTxt(Tier), "* Todos los personajes de tu cuenta reciben los beneficios del Tier escogido", 2, 51, 223, 0, 1)
            Call AddtoRichTextBox(RecTxt(Tier), "* Suscripción MENSUAL", 200, 200, 200, 0, 1)
            
            Price = "450DSP"

    End Select
    
    Me.lblPrice(Tier).Caption = Price
     
    Me.lblTitle(Tier).Caption = Tittle
    
    RecTxt(Tier).SelStart = 0

End Sub

Private Sub lblInfoTier_Click(Index As Integer)
    PicTier_Click Index

End Sub

Private Sub lblPrice_Click(Index As Integer)
    PicTier_Click Index

End Sub

Private Sub lblTitle_Click(Index As Integer)
    PicTier_Click Index

End Sub

Private Sub PicTier_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)

    Dim A As Long
    
    For A = 0 To 2
        PicTier(A).Picture = LoadPicture(DirInterface & "\shop\tier.jpg")
        SetTier A
    Next A
    
    PicTier(Index).Picture = LoadPicture(DirInterface & "\shop\tierselected.jpg")
    
    SetTier Index

End Sub

Private Sub RecTxt_Click(Index As Integer)
    PicTier_Click Index

End Sub

Public Sub UpdateMeditationLearn()

    Dim MeditationID As Integer
    
    If SelectedTienda > 0 Then
        If ShopCopy(SelectedTienda).ID > 0 And ShopCopy(SelectedTienda).ObjIndex = 9998 Then
            If ShopCopy(SelectedTienda).ObjAmount = ClientMeditation Then
                imgNoUsage.visible = True
                imgUsage.visible = False
                imgAdd.visible = False
            Else
                MeditationID = Val(ReadField(2, Shop(ShopCopy(SelectedTienda).ID).Name, Asc(" ")))
            
                If MeditationUser(MeditationID) = 1 Then
                    imgNoUsage.visible = False
                    imgUsage.visible = True
                    imgAdd.visible = False
                Else
                    imgNoUsage.visible = False
                    imgUsage.visible = False
                    imgAdd.visible = True
                End If

            End If

        End If

    End If

End Sub

Private Sub AjustarClic()
    
    Dim A As Long
    Dim T As Long, L As Long, W As Long, H As Long
    Dim Last As Long
    Dim X As Long, Y As Long
    Dim AddX As Long, AddY As Long
    
    Select Case Panel
    
        Case ePanel.eAvatars
            T = 32
            L = 14
            H = 153
            W = 129
            Last = 2
            AddX = 129
            AddY = 153
            
        Case Else
            T = 40
            L = 24
            H = 65
            W = 233
            Last = 4
            AddX = 300
            AddY = 40
    End Select
    
    For A = imgItem.LBound To imgItem.UBound
        imgItem(A).Width = W
        imgItem(A).Height = H
        
        imgItem(A).Top = T + Y
        imgItem(A).Left = L + X
            
        If (A + 1) Mod Last = 0 Then
            X = X + AddX
            Y = 0
        Else
            Y = Y + (imgItem(0).Top + AddY)
        End If
      
    Next A
End Sub
