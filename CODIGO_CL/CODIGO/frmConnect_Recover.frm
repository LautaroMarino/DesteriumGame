VERSION 5.00
Begin VB.Form frmConnect_Recover 
   BorderStyle     =   0  'None
   Caption         =   "Recordar Datos"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmConnect_Recover.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLastName 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   4650
      TabIndex        =   4
      Top             =   5460
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtDNI 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   4710
      TabIndex        =   2
      Top             =   7050
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtDateBirth 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   6210
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
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
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   4200
      Width           =   3675
   End
   Begin VB.Image imgSend 
      Height          =   405
      Left            =   4710
      Top             =   7680
      Width           =   2505
   End
   Begin VB.Image imgUnload 
      Height          =   435
      Left            =   10620
      Top             =   720
      Width           =   1005
   End
End
Attribute VB_Name = "frmConnect_Recover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
          
End Sub

Private Sub imgSend_Click()
    Call Audio.PlayInterface(SND_CLICK)
        
    Account.Email = txtEmail.Text
    'Account.key = txtKey.Text
    'Account.FirstName = txtFirstName.Text
    'Account.LastName = txtLastName.Text
    'Account.DateBirth = txtDateBirth.Text
    'Account.DNI = txtDNI.Text
    
    If Not CheckData Then Exit Sub
    Prepare_And_Connect E_MODO.e_LoginAccountRecover, Me
    
    Unload Me
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    frmConnect.Visible = True
    Unload Me
End Sub

Private Function CheckData() As Boolean
    
    If (Len(Account.Email) <= 0) Or Not CheckMailString(Account.Email) Then
        MsgBox "Email inválido."
        
        Exit Function
    End If
    
    'If Len(Account.key) <= 0 Then
         'MsgBox "Clave pin incorrecta. La misma contiene 20 caracteres aleatoreos y se envío a tu email con la creación de tu cuenta."
        
        'Exit Function
    'End If
    
    'If Len(Account.FirstName) <= 0 Then
         'MsgBox "Apellido inválido."
        
        'Exit Function
    'End If
    
    
    'If Len(Account.DateBirth) <= 0 Then
        'MsgBox "Fecha de nacimiento inválida."
        'Exit Function
    'End If
    
    'If Len(Account.DNI) <= 0 Or Not IsNumeric(Account.DNI) Then
        'MsgBox "DNI o número identificador inválido."
        
        'Exit Function
    'End If
    
    CheckData = True
End Function
