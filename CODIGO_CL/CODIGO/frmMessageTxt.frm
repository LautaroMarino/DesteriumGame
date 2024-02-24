VERSION 5.00
Begin VB.Form frmMessageTxt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mensajes Predefinidos"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMessageTxt.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMessageTxt.frx":000C
   ScaleHeight     =   4695
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   3720
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   3000
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   2630
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   2250
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   1870
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   " "
      Top             =   1490
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1120
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   780
      Width           =   3165
   End
   Begin VB.TextBox messageTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   420
      Width           =   3165
   End
   Begin VB.Image ImgCancelar 
      Height          =   375
      Left            =   120
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Image ImgGuardar 
      Height          =   375
      Left            =   3240
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmMessageTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario  As clsFormMovementManager

Private cBotonGuardar  As clsGraphicalButton

Private cBotonCancelar As clsGraphicalButton

Public LastPressed     As clsGraphicalButton

Private Const MAX_LEN_MESSAGE As Long = 60

Private Sub Form_Load()

    Dim I As Long
          
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
          
    For I = 0 To 9
        messageTxt(I) = CustomMessages.Message(I)
    Next I

    '  Me.Picture = LoadPicture(App.path & "\graficos\VentanaMensajesPersonalizados.jpg")
          
    LoadButtons
          
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String
          
    GrhPath = DirGraficos
          
    Set cBotonGuardar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
          
    Set LastPressed = New clsGraphicalButton

    ' Call cBotonGuardar.Initialize(imgGuardar, GrhPath & "BotonGuardarCustomMsg.jpg", GrhPath & "BotonGuardarRolloverCustomMsg.jpg", GrhPath & "BotonGuardarClickCustomMsg.jpg", Me)
    ' Call cBotonCancelar.Initialize(ImgCancelar, GrhPath & "BotonCancelarCustomMsg.jpg", GrhPath & "BotonCancelarRolloverCustomMsg.jpg", GrhPath & "BotonCancelarClickCustomMsg.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgGuardar_Click()

    On Error GoTo ErrHandler

    Dim I As Long
    
    For I = 0 To 9
        CustomMessages.Message(I) = messageTxt(I)
    Next I
          
    Unload Me

    Exit Sub

ErrHandler:

    'Did detected an invalid message??
    If err.Number = CustomMessages.InvalidMessageErrCode Then
        Call MsgBox("El Mensaje " & CStr(I + 1) & " es inválido. Modifiquelo por favor.")
    End If

End Sub

Private Sub messageTxt_MouseMove(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    'LastPressed.ToggleToNormal
End Sub
