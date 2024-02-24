VERSION 5.00
Begin VB.Form FrmChangeNick 
   Caption         =   "Acta de Nacimiento"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChangeNick.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmbConfirm 
      Caption         =   "Confirmar"
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   990
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
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2625
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4425
   End
End
Attribute VB_Name = "FrmChangeNick"
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

    If Not TieneObjetos(ACTA_NACIMIENTO) > 0 Then
        Call MsgBox("¡No tienes el Acta requerida para cambiar tu nombre!")
        Exit Sub
    End If
    
    Call WriteChangeNick(UserName, False)
    Unload Me
End Sub

Private Sub Form_Load()
    lblDesc.Caption = "Con el Acta de Nacimiento en tus manos podrás cambiar el nombre de tu personaje por otro que escojas dentro de las reglas generales de la comunidad."
    
End Sub


Public Function CheckNickValid(ByVal UserName As String) As Boolean

    If Len(UserName) < ACCOUNT_MIN_CHARACTER_CHAR Then
        Call MsgBox("El nombre debe contener más de " & ACCOUNT_MIN_CHARACTER_CHAR & " caracteres.")
        Exit Function
    End If
    
    If Len(UserName) > ACCOUNT_MAX_CHARACTER_CHAR Then
        Call MsgBox("El nombre debe contener menos de " & ACCOUNT_MAX_CHARACTER_CHAR & " caracteres.")
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
