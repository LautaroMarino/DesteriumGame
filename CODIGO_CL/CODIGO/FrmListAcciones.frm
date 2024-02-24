VERSION 5.00
Begin VB.Form FrmListAcciones 
   BorderStyle     =   0  'None
   Caption         =   "Acciones"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmListAcciones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2160
      Top             =   0
   End
   Begin VB.Image imgAction 
      Height          =   255
      Index           =   5
      Left            =   120
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Image imgAction 
      Height          =   255
      Index           =   4
      Left            =   120
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Image imgAction 
      Height          =   255
      Index           =   3
      Left            =   120
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image imgAction 
      Height          =   255
      Index           =   2
      Left            =   120
      Top             =   960
      Width           =   2655
   End
   Begin VB.Image imgAction 
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   720
      Width           =   2655
   End
   Begin VB.Image imgAction 
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "FrmListAcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const PIXEL_HEIGHT As Integer = 220

Public pixelHeight As Long

Private Height_Original As Integer
Private Width_Original As Integer

Public FormMovement As clsFormMovementManager

Private Enum eAction
    ACTION_NONE = 0
    COMMERCE_NPC = 1                ' La criatura comercia
    BANK_INIT = 2                           ' Inicia el banco
    BANK_INIT_ACCOUNT = 3          ' Inicia el banco compartido
    BANK_MERCADER = 4                 ' Inicia el Mercader
    CRAFT_NPC = 5
    RESU_NPC = 6
    
    INFO_NPC = 100 ' Criaturas hostiles del mundo.
End Enum
        
Private Type tLine
        Text As String
        Color As Long
        Font As Integer
        Size As Integer
        Action As eAction
End Type


Public LastLine As Integer

Private Hover_Action() As Boolean
Private lines() As tLine

Public NpcIndex_Selected As Integer

Private Sub Hover_Reset()
    Dim A As Long
    
    For A = LBound(Hover_Action) To UBound(Hover_Action)
        Hover_Action(A) = False
    Next A
    
End Sub
Private Sub Add_Line(ByVal Text As String, ByVal Color As Long, Font As Byte, Size As Byte, Optional ByRef Actione As eAction = ACTION_NONE)
    
    LastLine = LastLine + 1
    ReDim Preserve lines(0 To LastLine) As tLine
    ReDim Preserve Hover_Action(0 To LastLine) As Boolean
    
    With lines(LastLine)
        .Text = Text
        .Color = Color
        .Font = Font
        .Size = Size
        .Action = Actione
    End With
    
    pixelHeight = pixelHeight + PIXEL_HEIGHT
End Sub

Private Sub Prepare_Npcs()
    
    LastLine = 0
    pixelHeight = 0
    ReDim Preserve lines(LastLine) As tLine
    ReDim Preserve Hover_Action(0) As Boolean
     
    Dim Npc As tNpcs
    
    Npc = NpcList(SelectedNpcIndex)

    ' Nombre de la criatura
    Call Add_Line(Npc.Name, ARGB(255, 255, 255, 255), eFonts.f_Tahoma, 14)

    Select Case Npc.NpcType
    
        Case eNPCType.Banquero
            Call Add_Line("Banco del Personaje", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14, BANK_INIT)
            Call Add_Line("Banco de la Cuenta", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14, BANK_INIT_ACCOUNT)
            Call Add_Line("Mercado Global", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14, BANK_MERCADER)
        
        Case eNPCType.Revividor, eNPCType.ResucitadorNewbie
            Call Add_Line("Curar Personaje", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14, RESU_NPC)
        
        
        Case Else

            If Npc.Comercia = 1 Then
                Call Add_Line("Comercio", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14, COMMERCE_NPC)

            End If
            
            
            If Npc.Craft > 0 Then
                Call Add_Line("Fabricación", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14, CRAFT_NPC)

            End If
            
            
            If Npc.MaxHp > 0 Then
                Call Add_Line("Vida: " & PonerPuntos(Npc.MaxHp), ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)
            End If
            
            If Npc.GiveExp > 0 Then
                Call Add_Line("Exp: " & PonerPuntos(Npc.GiveExp), ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)
            End If
            
            If Npc.GiveGld > 0 Then
                Call Add_Line("Oro: " & PonerPuntos(Npc.GiveGld), ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)
            End If
            
            
            If Npc.MinHit > 0 Then
                Call Add_Line("Hit: " & Npc.MinHit & "/" & Npc.MaxHit, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

            End If
        
            If Npc.Def > 0 Then
                Call Add_Line("Def: " & Npc.Def, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

            End If
        
            If Npc.DefM > 0 Then
                Call Add_Line("RM: " & Npc.DefM, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

            End If

    End Select

    Me.Height = Me.Height + pixelHeight
    MirandoOpcionesNpc = True

End Sub

Public Sub Initial_Form()
     
    
    Me.visible = False

    Width_Original = 2900
    Height_Original = 490
    
    Me.Width = Width_Original
    Me.Height = Height_Original

    If MirandoOpcionesNpc Then
        Close_Form
    End If
    
    Call Prepare_Npcs
    g_Captions(eCaption.cCriaturaInfo) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    Render_List
    Me.visible = True
    FrmMain.SetFocus
    
End Sub
Public Sub Close_Form()
    MirandoOpcionesNpc = False
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cCriaturaInfo))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Set FormMovement = New clsFormMovementManager
    Call FormMovement.Initialize(Me, 32)
    
    Initial_Form
End Sub

Public Sub Render_List()
    
    Dim A        As Long

    Dim Y_Avance As Long
    
    Dim Color    As Long

    Dim Tier     As Byte
    
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cCriaturaInfo))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, Me.ScaleWidth, Me.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    ' Borde Superior
    Call Draw_Texture_Graphic_Gui(129, 0, 0, To_Depth(1), 193, 16, 0, 0, 193, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
    Y_Avance = 16
      
    For A = 1 To UBound(lines)

        With lines(A)
            Call Draw_Text(.Font, .Size, 15, Y_Avance + 1, To_Depth(3), 0, IIf(Hover_Action(A) = True, ARGB(255, 166, 0, 255), .Color), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, .Text, False, True)
            Call Draw_Texture_Graphic_Gui(130, 0, Y_Avance, To_Depth(1), 193, 16, 0, 0, 193, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
            Y_Avance = Y_Avance + 16

        End With

    Next A

    Call Draw_Texture_Graphic_Gui(131, 0, Y_Avance - 3, To_Depth(2), 193, 16, 0, 0, 193, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
    
    Call wGL_Graphic_Renderer.Flush

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Hover_Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Close_Form
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
     Unload Me
End Sub

Private Sub Click_Action(ByRef Action As eAction)
    
    Dim X As Long, Y As Long
    
    Select Case Action
    
        Case eAction.ACTION_NONE
        
        Case eAction.BANK_INIT
            Call WriteBankStart(E_BANK.e_User)

        Case eAction.BANK_INIT_ACCOUNT
            Call WriteBankStart(E_BANK.e_Account)

        Case eAction.BANK_MERCADER
            Call WriteMercader_Required(1, 1, 255)
            
        Case eAction.COMMERCE_NPC

            If CharIndex_MouseHover > 0 Then
                X = CharList(CharIndex_MouseHover).Pos.X
                Y = CharList(CharIndex_MouseHover).Pos.Y
                    
                Call WriteDoubleClick(X, Y, 1)

            End If
            
        Case eAction.CRAFT_NPC

            If CharIndex_MouseHover > 0 Then
                X = CharList(CharIndex_MouseHover).Pos.X
                Y = CharList(CharIndex_MouseHover).Pos.Y
                    
                Call WriteDoubleClick(X, Y, 2)

            End If
            
        Case eAction.RESU_NPC
            If CharIndex_MouseHover > 0 Then
                Call ParseUserCommand("/RESUCITAR")
            End If
            
        Case eAction.INFO_NPC
    
    End Select
    
    Unload Me

End Sub

Private Sub imgAction_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    If Index + 2 > UBound(lines) Then Exit Sub
    Call Click_Action(lines(Index + 2).Action)
    
End Sub

Private Sub imgAction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Hover_Reset
    
     If Index + 2 > UBound(lines) Then Exit Sub
    Hover_Action(Index + 2) = True
End Sub

Private Sub Timer1_Timer()
    If MirandoOpcionesNpc Then
        Render_List
    End If
End Sub
