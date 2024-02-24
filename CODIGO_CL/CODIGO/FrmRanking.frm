VERSION 5.00
Begin VB.Form FrmRanking 
   BorderStyle     =   0  'None
   Caption         =   "Ranking de Usuarios"
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
   Icon            =   "FrmRanking.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tUpdate 
      Interval        =   250
      Left            =   4200
      Top             =   480
   End
   Begin VB.PictureBox picRank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   360
      MousePointer    =   99  'Custom
      Picture         =   "FrmRanking.frx":000C
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4560
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmRanking.frx":ABEC
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblSubTittle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIVEL"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   420
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label lblSubTittle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   420
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EVENTOS"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Index           =   3
      Left            =   1875
      TabIndex        =   2
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETOS"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   315
      Index           =   2
      Left            =   585
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NIVEL"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   420
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Width           =   1185
   End
End
Attribute VB_Name = "FrmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario          As clsFormMovementManager


Public Sub DefaultTittle()
    Dim A As Long
    
    For A = lblTitle.LBound To lblTitle.UBound
        lblTitle(A).ForeColor = RGB(240, 220, 175)
    Next A
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    g_Captions(eCaption.RankTop) = wGL_Graphic.Create_Device_From_Display(picRank.hWnd, picRank.ScaleWidth, picRank.ScaleHeight)
    
    Me.Picture = LoadPicture(DirInterface & "menucompacto\rank.jpg")
    lblTitle(1).ForeColor = RGB(238, 190, 0)
    
        #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoRank = False
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub lblTitle_Click(Index As Integer)
        
    Call DefaultTittle
    lblTitle(Index).ForeColor = RGB(238, 190, 0)
    
    Select Case Index
    
        Case 1 ' Nivel
            lblSubTittle(0).Caption = "NIVEL"
            
        Case 2 ' Retos
            lblSubTittle(0).Caption = "BALANCE"
        Case 3 ' Torneo
            lblSubTittle(0).Caption = "PUNTOS"
    End Select
    
    Call WriteRequestRank(Index)
    
End Sub

Private Sub picRank_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

' Lista Gráfica de Hechizos
Private Sub picRank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picRank.ScaleHeight / RankList.Pixel_Alto) * RankList.Pixel_Alto - 1 Then Y = Int(picRank.ScaleHeight / RankList.Pixel_Alto) * RankList.Pixel_Alto - 1
If X < picRank.ScaleWidth - 10 Then
    RankList.ListIndex = Int(Y / RankList.Pixel_Alto) + RankList.Scroll
    RankList.DownBarrita = 0

Else
    RankList.DownBarrita = Y - RankList.Scroll * (picRank.ScaleHeight - RankList.BarraHeight) / (RankList.ListCount - RankList.VisibleCount)
End If
End Sub

Private Sub picRank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picRank.ScaleHeight / RankList.Pixel_Alto) * RankList.Pixel_Alto - 1 Then yy = Int(picRank.ScaleHeight / RankList.Pixel_Alto) * RankList.Pixel_Alto - 1
    If RankList.DownBarrita > 0 Then
        RankList.Scroll = (Y - RankList.DownBarrita) * (RankList.ListCount - RankList.VisibleCount) / (picRank.ScaleHeight - RankList.BarraHeight)
    Else
        RankList.ListIndex = Int(yy / RankList.Pixel_Alto) + RankList.Scroll

        'If ScrollArrastrar = 0 Then
          '  If (Y < yy) Then RankList.Scroll = RankList.Scroll - 1
         '   If (Y > yy) Then RankList.Scroll = RankList.Scroll + 1
       ' End If
    End If
ElseIf Button = 0 Then
    RankList.ShowBarrita = X > picRank.ScaleWidth - RankList.BarraWidth * 2
End If
End Sub

Private Sub picRank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RankList.DownBarrita = 0
End Sub
