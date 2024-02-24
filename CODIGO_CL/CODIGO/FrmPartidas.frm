VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmPartidas 
   BorderStyle     =   0  'None
   Caption         =   "Partidas"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "FrmPartidas.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbEvents 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   360
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1140
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox Console 
      Height          =   3540
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   3240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6244
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmPartidas.frx":12D71
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
   Begin RichTextLib.RichTextBox ConsoleEvent 
      Height          =   720
      Left            =   540
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   1680
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   1270
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FrmPartidas.frx":12DEE
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
   Begin VB.Image imgInscribirse 
      Height          =   375
      Left            =   2280
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Image imgReturn 
      Height          =   255
      Left            =   1440
      Top             =   6960
      Width           =   2415
   End
End
Attribute VB_Name = "FrmPartidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Consola Transparente
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&

Private Sub cmbEvents_Click()
    
    If cmbEvents.ListIndex < 0 Then Exit Sub
    
    Dim ID As Integer
    ConsoleEvent.Text = vbNullString
    
    ID = SearchID(cmbEvents.List(cmbEvents.ListIndex))
    
    Call Events_GenerateSpam(ID, Me.ConsoleEvent)
End Sub
Private Function SearchID(ByVal Name As String) As Integer
    Dim A As Long
    
    For A = 1 To MAX_EVENT_SIMULTANEO
        With Events(A)
            If StrComp(.Name, Name) = 0 Then
                SearchID = A
                Exit Function
            End If
        End With
    Next A
    
End Function
Private Sub Form_Load()
    
    MirandoPartidas = True
    
    Call SetWindowLong(Console.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Call SetWindowLong(ConsoleEvent.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Call Tournaments_List
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoPartidas = False
End Sub

Public Sub Tournaments_List()

    Dim A As Long, TextSpam As String
    Dim Selected As Boolean
    
    
    ConsoleEvent.Text = vbNullString
    ConsoleEvent.SelStart = 0
    cmbEvents.Clear
    Console.Text = vbNullString
    Console.SelStart = 0
    
    For A = 1 To MAX_EVENT_SIMULTANEO
        With Events(A)
            
            If Not (.AllowedClasses(UserClase)) = 0 Then
                If .Name <> vbNullString Then
                    cmbEvents.AddItem .Name
                    Call Events_GenerateSpam(A, Me.Console)
                    Selected = True
                End If
            End If
        End With
    Next A
 
    If Selected Then
        cmbEvents.ListIndex = 0
    End If
End Sub

Private Sub imgInscribirse_Click()
    Call Audio.PlayInterface(SND_CLICK)
    If cmbEvents.ListIndex < 0 Then Exit Sub
    
    Call WriteParticipeEvent(cmbEvents.List(cmbEvents.ListIndex))
    
    
End Sub

Private Sub imgReturn_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    
    #If ModoBig = 1 Then
        dockForm FrmMenu.hWnd, FrmMain.PicMenu, True
        Unload Me
    #Else
        Unload Me
    #End If
End Sub
