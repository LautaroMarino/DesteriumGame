VERSION 5.00
Begin VB.Form FrmConnect_NewAccount 
   BorderStyle     =   0  'None
   Caption         =   "Nueva cuenta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   HasDC           =   0   'False
   LinkTopic       =   "Nueva Cuenta"
   Picture         =   "FrmConnect_NewAccount.frx":0000
   ScaleHeight     =   9490.334
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8100
      TabIndex        =   4
      Top             =   5940
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
      Left            =   8070
      TabIndex        =   3
      Top             =   5100
      Visible         =   0   'False
      Width           =   2715
   End
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
      Left            =   8040
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   8040
      TabIndex        =   1
      Top             =   3480
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
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   2640
      Width           =   2715
   End
   Begin VB.Image imgUnload 
      Height          =   405
      Left            =   10710
      Top             =   750
      Width           =   885
   End
   Begin VB.Image imgAccept 
      Height          =   405
      Left            =   3930
      Top             =   7650
      Width           =   3945
   End
End
Attribute VB_Name = "FrmConnect_NewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Image1_Click()

End Sub

Private Sub imgAccept_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Account.Email = txtEmail.Text
    'Account.DateBirth = txtDateBirth.Text
    'Account.DNI = txtDNI.Text
    'Account.FirstName = txtFirstName.Text
    'Account.LastName = txtLastName.Text
    
    Dim Temp As String
    
    Temp = "Email: " & Account.Email & vbCrLf & vbCrLf & "¿Los datos son correctos? ¡No podrás cambiarlo!"
    'Temp = Temp & "Nombre y Apellido: " & Account.FirstName & " " & Account.LastName & vbCrLf
    'Temp = Temp & "DNI: " & Account.DNI & vbCrLf
    'Temp = Temp & "Fecha de nacimiento: " & Account.DateBirth & vbCrLf
    'Temp = Temp & vbCrLf & vbCrLf & "¿Los datos son correctos? ¡No podrás cambiarlos!"
    
    If MsgBox(Temp, vbYesNo) = vbYes Then
        Prepare_And_Connect E_MODO.e_LoginAccountNew
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        FrmConnect_Account.Visible = True
        FrmConnect_Account.SelectedPanelAccount (ePrincipal)
        Unload Me
    End If

End Sub


Private Function CheckData() As Boolean
    
    If (Len(Account.Email) <= 0) Or Not CheckMailString(UserEmail) Then
        MsgBox "Email inválido."
        
        Exit Function
    End If
    
    'If Len(Account.FirstName) <= 0 Then
         'MsgBox "Nombre inválido."
        
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
        
       ' Exit Function
    'End If
    
    CheckData = True
End Function

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub
