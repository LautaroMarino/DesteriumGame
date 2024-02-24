VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMapaNpc 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   Picture         =   "FrmMapaNpc.frx":0000
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tUpdate 
      Interval        =   100
      Left            =   4440
      Top             =   2760
   End
   Begin VB.PictureBox PicNpc 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   1200
      ScaleHeight     =   202
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   1
      Top             =   1620
      Width           =   2955
   End
   Begin RichTextLib.RichTextBox Console 
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes de eventos"
      Top             =   5520
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmMapaNpc.frx":11034
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
   Begin VB.Image Button 
      Height          =   375
      Index           =   2
      Left            =   3480
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   1
      Left            =   1920
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Image Button 
      Height          =   375
      Index           =   0
      Left            =   360
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Criatura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   1170
      TabIndex        =   0
      Top             =   1110
      Width           =   3015
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMapaNpc.frx":110B2
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMapaNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private clsFormulario As clsFormMovementManager
Private picCheckBox          As Picture
Private picCheckBoxNulo      As Picture

Public NpcIndex As Integer
Public Mirando As Byte

Private Sub Button_Click(Index As Integer)

    
    Console.visible = False
    
    
    Mirando = Index
    
    Select Case Index
        Case 0 ' Estadisticas
            Console.visible = True
            
            Call LoadStats
        Case 1 ' Drops
        
        Case 2 ' Info extra, descripciones
    End Select
End Sub

Private Sub LoadStats()


    Console.Text = vbNullString
    Console.SelStart = 0
    
    With NpcList(NpcIndex)
        Call AddtoRichTextBox(Console, "Estadísticas de la criatura: ", 255, 255, 255, True, False)
        
        ' # Vida
        If .MaxHp > 0 Then
            Call AddtoRichTextBox(Console, "Vida: ", 255, 255, 255, True, False)
            Call AddtoRichTextBox(Console, PonerPuntos(.MaxHp), 255, 255, 0, True, False, False)
        End If
        
        ' Hit
        If .MinHit > 0 Then
            Call AddtoRichTextBox(Console, "Hit: ", 255, 255, 255, True, False)
            Call AddtoRichTextBox(Console, PonerPuntos(CLng(.MinHit)) & "/" & PonerPuntos(CLng(.MaxHit)), 255, 255, 0, True, False, False)
        End If
        
         ' Defensa
         If .Def > 0 Then
            Call AddtoRichTextBox(Console, "Defensa: ", 255, 255, 255, True, False)
            Call AddtoRichTextBox(Console, PonerPuntos(CLng(.Def)), 255, 255, 0, True, False, False)
        End If
        
        ' Exp
         If .GiveExp > 0 Then
            Call AddtoRichTextBox(Console, "Experiencia: ", 255, 255, 255, True, False)
            Call AddtoRichTextBox(Console, PonerPuntos(CLng(.GiveExp)), 255, 255, 0, True, False, False)
        End If
    End With
End Sub

Public Sub UpdateNpc(ByVal INpc As Integer)
    NpcIndex = INpc
    lblName.Caption = NpcList(INpc).Name
    
    If Mirando = 0 Then ' Estadissticas
        Call LoadStats
    ElseIf Mirando = 1 Then ' Drops
    
    Else ' Info
        
    End If
    
    
    Render
End Sub
Private Sub Form_Load()

    g_Captions(eCaption.eMapaNpc) = wGL_Graphic.Create_Device_From_Display(PicNpc.hWnd, PicNpc.ScaleWidth, PicNpc.ScaleHeight)
        
    Me.Picture = LoadPicture(App.path & "\resource\interface\menucompacto\VentanaNpc.jpg")
        
        
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    
    Dim GrhPath As String
    
    GrhPath = DirInterface
                        
    Set picCheckBox = LoadPicture(DirInterface & "options\CheckBoxOpciones.jpg")
    Set picCheckBoxNulo = LoadPicture(DirInterface & "options\CheckBoxOpcionesNulo.jpg")
    
    MirandoNpc = True
    
    NpcIndex = FrmMapa.NpcIndexSelected
    Call UpdateNpc(NpcIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eMapaNpc))
    MirandoNpc = False
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub



' # Render

' # Renderizado del mapa
Private Sub Render()
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.eMapaNpc))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, PicNpc.ScaleWidth, PicNpc.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    
    ' Personaje
    
    Dim X As Long
    Dim Y As Long
    Dim GrhIndex As Long
    Dim A As Long
    
        
    X = 80
    Y = 170
    
    
    Dim Width As Long
    Dim Height As Long
    Dim ChangeTamaño As Boolean
    
    If NpcIndex > 0 Then
        With NpcList(NpcIndex)
        
            If .Body > 0 Then
                GrhIndex = BodyData(.Body).Walk(E_Heading.SOUTH).GrhIndex
                  
                Width = GrhData(GrhIndex).pixelWidth
                Height = GrhData(GrhIndex).pixelHeight
    
                If Width > 200 Or Height > 200 Then
                    Width = Width * 0.7
                    Height = 200 * 0.7

                    ChangeTamaño = True
                End If
     
                
              
                Call Draw_Grh(BodyData(.Body).Walk(E_Heading.SOUTH), X + BodyData(.Body).BodyOffSet(E_Heading.SOUTH).X, Y + BodyData(.Body).BodyOffSet(E_Heading.SOUTH).Y, To_Depth(6), 1, 1, 0, , , eTechnique.t_Alpha, _
                           Width, Height)
            End If
                
            If .Head > 0 Then
                GrhIndex = HeadData(.Head).Head(E_Heading.SOUTH).GrhIndex
                 
                Width = IIf(ChangeTamaño, GrhData(GrhIndex).pixelWidth * 0.7, GrhData(GrhIndex).pixelWidth)
                Height = IIf(ChangeTamaño, GrhData(GrhIndex).pixelHeight * 0.7, GrhData(GrhIndex).pixelHeight)

                Call Draw_Grh(HeadData(.Head).Head(E_Heading.SOUTH), X + BodyData(.Body).HeadOffset.X, Y + BodyData(.Body).HeadOffset.Y, To_Depth(6), 1, 1, , , , , _
                Width, Height)
            End If
        End With
   End If
    
    Call wGL_Graphic_Renderer.Flush
End Sub

Private Sub tUpdate_Timer()
    Render
End Sub
