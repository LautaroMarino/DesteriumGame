VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmStatsUser 
   BorderStyle     =   0  'None
   Caption         =   "Panel del Personaje"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "FrmStatsUser.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
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
      Height          =   1395
      Left            =   3315
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1410
      Width           =   1395
   End
   Begin VB.PictureBox PicInv 
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
      Left            =   720
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4095
      Visible         =   0   'False
      Width           =   3840
   End
   Begin RichTextLib.RichTextBox ListView 
      Height          =   2820
      Left            =   615
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   4095
      Visible         =   0   'False
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   4974
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmStatsUser.frx":14C15
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
   Begin VB.Image Button 
      Height          =   375
      Index           =   7
      Left            =   3765
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   6
      Left            =   3300
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   5
      Left            =   2835
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   4
      Left            =   2385
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   3
      Left            =   1935
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   2
      Left            =   1485
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   1
      Left            =   1035
      Top             =   3645
      Width           =   390
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   0
      Left            =   585
      Top             =   3645
      Width           =   390
   End
   Begin VB.Label lblMap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1-50-50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   3480
      TabIndex        =   13
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Label lblFrags 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2070
      TabIndex        =   12
      Top             =   3285
      Width           =   915
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999.999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   720
      TabIndex        =   11
      Top             =   3285
      Width           =   915
   End
   Begin VB.Label lblDsp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999.999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   2070
      TabIndex        =   10
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label lblGld 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99.999.999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label lblHasta 
      BackStyle       =   0  'Transparent
      Caption         =   "Durante 3hs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   1800
      TabIndex        =   8
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label lblBlocked 
      BackStyle       =   0  'Transparent
      Caption         =   "SI."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1420
      TabIndex        =   7
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label lblHp 
      BackStyle       =   0  'Transparent
      Caption         =   "480 [+16]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   900
      TabIndex        =   6
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label lblElv 
      BackStyle       =   0  'Transparent
      Caption         =   "43 (39%)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label lblRaze 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo Oscuro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   4
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      Caption         =   "Clerigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblGenero 
      BackStyle       =   0  'Transparent
      Caption         =   "Mujer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   1620
      Width           =   1395
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lion Ragnarok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1110
      Width           =   3135
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmStatsUser.frx":14C93
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmStatsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Inventory As clsGrapchicalInventory
Private clsFormulario As clsFormMovementManager

Private Enum eButtons
    eInventory = 0
    eSpells = 1
    eBank = 2
    eAbilities = 3
    eBonus = 4
    ePenas = 5
    eSkins = 6
    eLogros = 7
End Enum

Private Const MAX_INVENTORY As Byte = 48
Private Const MAX_PICS As Byte = 7

Private picButtons(MAX_PICS) As Picture
Private ButtonActive(MAX_PICS) As Byte

' Define RGB colors for different types of text
Dim TitleColor() As Variant
Dim DescColor() As Variant
Dim ValueColor() As Variant
    
 
Private Sub Button_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    Call Button_Selected(Index)
End Sub

Private Sub Button_Reset_All()
    
    Dim A As Long
    
    For A = 0 To MAX_PICS
        ButtonActive(A) = 0
        Set Button(A).Picture = Nothing
    Next A
    
End Sub
Private Sub Button_Selected(Index As Integer)
    
    Call Button_Reset_All
    
    If ButtonActive(Index) = 0 Then
        Set Button(Index).Picture = picButtons(Index)
        Call Button_Action(Index)
    End If
    
End Sub

Private Sub Button_Action(Index As Integer)

    PicInv.visible = False
    ListView.visible = False
    
    Select Case Index
        Case eButtons.eInventory
            PicInv.visible = True
            
        Case eButtons.eSpells
            ListView.visible = True
            
        Case eButtons.eBank
            PicInv.visible = True
            
        Case eButtons.eAbilities
            ListView.visible = True
            
        Case eButtons.eBonus
            ListView.visible = True
            
        Case eButtons.ePenas
            ListView.visible = True
            
        Case eButtons.eSkins
            PicInv.visible = True
            
        Case eButtons.eLogros
            ListView.visible = True
    End Select
    
    Call WriteRequiredStatsUser(Index, lblName.Caption)
    
End Sub
Private Sub Form_Load()
    TitleColor = Array(50, 205, 50)    ' Green
    DescColor = Array(70, 130, 180) ' Steel Blue
    ValueColor = Array(255, 165, 0) ' Orange
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadInterface
    Call Update_Info
    
    Set Inventory = New clsGrapchicalInventory
    Call Inventory.Initialize(PicInv, MAX_INVENTORY, MAX_INVENTORY, eCaption.eMercader_Inv, 32, 32)
    
    MirandoStatsUser = True
    
        
    g_Captions(eCaption.eMercader_Inv) = wGL_Graphic.Create_Device_From_Display(PicInv.hWnd, PicInv.ScaleWidth, PicInv.ScaleHeight)
End Sub

Public Sub Update_Info()

    ' # Reseteamos los botones
    Call Button_Reset_All
    
    Call Load_Stats
    
    
End Sub

' # Seteamos las estadísticas básicas en los labels
Private Sub Load_Stats()
    
    Dim ExpConverted As String
    Dim TempUP As Single
    Dim TempUPs As String
    
    
    If InfoUser.Elu > 0 Then
        ExpConverted = Round(CDbl(InfoUser.Exp) * CDbl(100) / CDbl(InfoUser.Elu), 2) & "%"
    End If
    
    TempUP = UserCheckPromedy(InfoUser.Elv, InfoUser.Hp, InfoUser.Clase, ModRaza(InfoUser.Raza).Constitucion)

    If TempUP >= 0 Then
        TempUPs = " +" & TempUP
    Else
        TempUPs = TempUP
    End If

    lblName.Caption = UCase$(InfoUser.UserName)
    lblClass.Caption = ListaClases(InfoUser.Clase)
    lblRaze.Caption = ListaRazas(InfoUser.Raza)
    lblGenero.Caption = IIf((InfoUser.Genero = 1), "Hombre", "Mujer")
        
    lblElv.Caption = InfoUser.Elv
        
    If InfoUser.Elv <> STAT_MAXELV Then
        lblElv.Caption = lblElv.Caption & " " & ExpConverted
    End If
        
    lblBlocked.Caption = IIf((InfoUser.Blocked = 1), "SI", "NO")
        
    lblBlocked.ForeColor = IIf((InfoUser.Blocked = 1), vbGreen, vbRed)
    
    If InfoUser.Blocked = 1 Then
        lblHasta.Caption = "Durante " & SecondsToHMS(InfoUser.BlockedHasta) ' Transformar a HH:MM
    Else
        lblHasta.visible = False
    End If
        
    lblGld.Caption = PonerPuntos(InfoUser.Gld)
    lblDsp.Caption = PonerPuntos(InfoUser.Dsp)
    lblPoints.Caption = PonerPuntos(InfoUser.Points)
        
    lblFrags.Caption = PonerPuntos(CLng(InfoUser.Frags))
        
    lblHp.Caption = InfoUser.Hp & TempUPs
        
    If InfoUser.Map = 0 Then
        lblMap = "Bloqueado"
    Else
        lblMap = InfoUser.Map & "-" & InfoUser.X & "-" & InfoUser.Y
    End If
End Sub


' # Seteamos el inventario
Public Sub Load_Inventory(ByVal IsBank As Boolean)


    Dim A As Long
    Dim ObjIndex As Integer
    
    
    If Not IsBank Then
        For A = 1 To Inventory.MaxObjs
            
            
            If A <= InfoUser.Inventory.NroItems Then
                ObjIndex = InfoUser.Inventory.Object(A).ObjIndex
                If ObjIndex > 0 Then
                    Call Inventory.SetItem(A, ObjIndex, InfoUser.Inventory.Object(A).Amount, InfoUser.Inventory.Object(A).Equipped, ObjData(ObjIndex).GrhIndex, 0, 0, 0, 0, 0, 0, ObjData(ObjIndex).Name, 0, True, 0, 0, 0, 0)
                Else
                    Call Inventory.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
                End If
            Else
                Call Inventory.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
            End If
            
        Next A
    Else
        
        For A = 1 To Inventory.MaxObjs
            
            
            If A <= InfoUser.Bank.NroItems Then
                ObjIndex = InfoUser.Bank.Object(A).ObjIndex
                If ObjIndex > 0 Then
                    Call Inventory.SetItem(A, ObjIndex, InfoUser.Bank.Object(A).Amount, InfoUser.Bank.Object(A).Equipped, ObjData(ObjIndex).GrhIndex, 0, 0, 0, 0, 0, 0, ObjData(ObjIndex).Name, 0, True, 0, 0, 0, 0)
                Else
                    Call Inventory.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
                End If
            Else
                Call Inventory.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
            End If
            
        Next A
    End If
    
    
    Inventory.DrawInventory

End Sub

' # Seteamos los hechizos
Public Sub Load_Spells()
    Dim A As Long
    
    ListView.Text = vbNullString
    ListView.SelStart = 0
    
    AddtoRichTextBox ListView, "'Hechizos que posee el personaje'", TitleColor(0), TitleColor(1), TitleColor(2), False, False, True
    
    For A = 1 To MAXHECHI
        If InfoUser.Spells(A) > 0 Then
            AddtoRichTextBox ListView, "Hechizo: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            AddtoRichTextBox ListView, Hechizos(InfoUser.Spells(A)).Nombre, ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
            
        End If
    Next A
    
End Sub

' # Seteamos los Skils
Public Sub Load_Skills()
    Dim A As Long
    
    ListView.Text = vbNullString
    ListView.SelStart = 0
    
    AddtoRichTextBox ListView, "'Habilidades que posee el personaje'", TitleColor(0), TitleColor(1), TitleColor(2), False, False, True
    
    For A = 1 To NUMSKILLS
        AddtoRichTextBox ListView, "Habilidad: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
        AddtoRichTextBox ListView, "'" & SkillsNames(A) & "'", ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
        AddtoRichTextBox ListView, " |" & InfoUser.Skills(A), ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
    Next A
    
End Sub
' # Seteamos las Penas
Public Sub Load_Penas()
    Dim A As Long
    
    ListView.Text = vbNullString
    ListView.SelStart = 0
    
    AddtoRichTextBox ListView, "'Penas que tiene el personaje'", TitleColor(0), TitleColor(1), TitleColor(2), False, False, True
    
    For A = 1 To InfoUser.PenasLast
        AddtoRichTextBox ListView, A & ") ", DescColor(2), False, False, True
        AddtoRichTextBox ListView, A, InfoUser.Penas(A), ValueColor(1), ValueColor(2), False, False, False
    Next A
    
End Sub

' # Seteamos los skins
Public Sub Load_Skins()
    Dim A As Long
    Dim ObjIndex As Integer
    
    For A = 1 To Inventory.MaxObjs
        If A <= InfoUser.Skins.Last Then
            ObjIndex = InfoUser.Skins.ObjIndex(A)
            
            Call Inventory.SetItem(A, ObjIndex, 0, 0, ObjData(ObjIndex).GrhIndex, 0, 0, 0, 0, 0, ObjData(ObjIndex).ValueGLD, ObjData(ObjIndex).Name, ObjData(ObjIndex).ValueDSP, True, 0, 0, 0, 0)
        Else
            Call Inventory.SetItem(A, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, True, 0, 0, 0, 0)
        End If
    Next A
        
        
    Inventory.DrawInventory
End Sub

' # Seteamos los Logros()
Public Sub Load_Logros()
    Dim A As Long
    
    ListView.Text = vbNullString
    ListView.SelStart = 0
    
    AddtoRichTextBox ListView, "'Logros que consiguió el personaje'", TitleColor(0), TitleColor(1), TitleColor(2), False, False, True
    
    'For A = 1 To MAXHECHI
        'If InfoUser.Spells(A) > 0 Then
            'AddtoRichTextBox ListView, "Hechizo: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
            'AddtoRichTextBox ListView, Hechizos(InfoUser.Spells(A)).Nombre, ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
            
        'End If
    'Next A
    
End Sub

' # Seteamos las bonificaciones
Public Sub Load_Bonus()
    
     
    Dim strTemp  As String
    Dim A As Long

    ListView.Text = vbNullString
    ListView.SelStart = 0
    
    For A = 1 To InfoUser.BonusLast
        With InfoUser.BonusUser(A)
            Select Case .Tipo
                Case eBonusType.eObj
                    AddtoRichTextBox ListView, "'" & ObjData(.Value).Name & "'", TitleColor(0), TitleColor(1), TitleColor(2), False, False, True
                    
                    AddtoRichTextBox ListView, "Duración: ", DescColor(0), DescColor(1), DescColor(2), False, False, True
                    
                    If .DurationSeconds > 0 Then
                        AddtoRichTextBox ListView, SecondsToHMS(.DurationSeconds), ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
                        
                    ElseIf .DurationDate <> vbNullString Then
                        AddtoRichTextBox ListView, .DurationDate, ValueColor(0), ValueColor(1), ValueColor(2), False, False, False
                    End If
                Case Else
            End Select
        End With
    Next A
    
    
End Sub










Private Sub LoadInterface()

    Dim filePath As String
    Dim A As Long
    
    filePath = DirInterface & "menucompacto\"
    
    For A = 0 To MAX_PICS
        Set picButtons(A) = LoadPicture(filePath & "CharInfo_Pic_" & A + 1 & ".jpg")
    Next A
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoStatsUser = False
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMercader_Inv))
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Unload Me
End Sub


Private Function ConsoleTipe(ByRef Tipo As eBonusType) As String
  
End Function
