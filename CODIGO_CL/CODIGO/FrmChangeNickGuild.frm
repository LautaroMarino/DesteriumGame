VERSION 5.00
Begin VB.Form FrmChangeNickGuild 
   Caption         =   "Cambio de Líder"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChangeNickGuild.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbConfirm 
      Caption         =   "Confirmar"
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   990
   End
   Begin VB.TextBox txtName 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4425
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje:"
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   2625
   End
End
Attribute VB_Name = "FrmChangeNickGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbConfirm_Click()
    Dim UserName As String
    
    UserName = txtName.Text
    
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
    End If
    
    If Not CheckNickValid(UserName) Then Exit Sub

    If Not TieneObjetos(ESCRITURAS_CLAN) > 0 Then
        Call MsgBox("¡No tienes las escrituras del clan para cambiar el liderazgo!")
        Exit Sub
    End If
    
    Call WriteChangeNick(UserName, True)
    Unload Me
End Sub

Private Sub Form_Load()
    lblDesc.Caption = "Con las escrituras del clan en tus manos podrás pasarle el liderazgo a la persona que elijas."
    
End Sub


Public Function CheckNickValid(ByVal UserName As String) As Boolean

    If Len(UserName) < ACCOUNT_MIN_CHARACTER_CHAR Then
        Call MsgBox("Nick inválido")
        Exit Function
    End If
    
    If Len(UserName) > ACCOUNT_MAX_CHARACTER_CHAR Then
        Call MsgBox("Nick inválido")
        Exit Function
    End If

    CheckNickValid = True
End Function

Private Sub txtName_Change()
    If txtName.Text <> vbNullString Then
        If Not ValidarNombre(txtName.Text) Then
             txtName.Text = Left(txtName.Text, Len(txtName.Text) - 1)
             txtName.SelStart = Len(txtName.Text)
        End If
    End If
End Sub

