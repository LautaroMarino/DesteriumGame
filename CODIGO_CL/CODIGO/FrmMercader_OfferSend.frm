VERSION 5.00
Begin VB.Form FrmMercader_OfferSend 
   BorderStyle     =   0  'None
   Caption         =   "Nueva Oferta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercader_OfferSend.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tCode 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10800
      Top             =   3840
   End
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   2190
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   3960
      Width           =   2505
   End
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   2190
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Top             =   3960
      Width           =   2505
   End
   Begin VB.TextBox txtGld 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "0"
      Top             =   7380
      Width           =   1755
   End
   Begin VB.TextBox txtEldhir 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "0"
      Top             =   7920
      Width           =   1755
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   0
      Left            =   8160
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Index           =   1
      Left            =   8520
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Image imgWeb 
      Height          =   255
      Left            =   6840
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Image imgValidate 
      Height          =   375
      Left            =   8160
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Image imgConfirm 
      Enabled         =   0   'False
      Height          =   405
      Left            =   8040
      Top             =   6480
      Width           =   2475
   End
   Begin VB.Image imgAdd 
      Height          =   390
      Left            =   1740
      Top             =   6480
      Width           =   840
   End
   Begin VB.Image imgRemove 
      Height          =   390
      Left            =   5160
      Top             =   6480
      Width           =   960
   End
   Begin VB.Image imgUnload 
      Height          =   435
      Left            =   10680
      Top             =   720
      Width           =   915
   End
End
Attribute VB_Name = "FrmMercader_OfferSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim A As Long
    
    lstChars(0).Clear
    
    For A = 1 To ACCOUNT_MAX_CHARS
        If Account.Chars(A).Name <> vbNullString Then
            lstChars(0).AddItem (Account.Chars(A).Name)
        End If
    Next A
End Sub

Private Sub imgAdd_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If lstChars(0).ListIndex = -1 Then
        Call MsgBox("Selecciona el personaje que deseas agregar")
        Exit Sub
    End If
    
    If SearchRepeat(lstChars(0).List(lstChars(0).ListIndex)) Then
        Call MsgBox("El personaje ya se encuentra en la lista.")
        Exit Sub
    End If
    
    If lstChars(1).ListCount = MAX_MERCADER_CHARS Then
        Call MsgBox("¡No puedes ofrecer más de " & MAX_MERCADER_CHARS & " personajes!")
        Exit Sub
    End If
    
    If Account.Premium = 0 Then
        If lstChars(1).ListCount = 1 Then
            Call MsgBox("Tu cuenta debe ser PREMIUM para ofrecer más de un personaje.")
            Exit Sub
        End If
    End If
    
    lstChars(1).AddItem (lstChars(0).List(lstChars(0).ListIndex))
End Sub

Private Function SearchRepeat(ByVal UserName As String) As Boolean
    Dim A As Long
    
    For A = 0 To lstChars(1).ListCount - 1
        If StrComp(UserName, lstChars(1).List(A)) = 0 Then
            SearchRepeat = True
            Exit Function
        End If
    Next A
End Function

Private Sub imgRemove_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If lstChars(1).ListIndex = -1 Then
        Call MsgBox("Selecciona el personaje que deseas quitar de la lista")
        Exit Sub
    End If
    
    lstChars(1).RemoveItem (lstChars(1).ListIndex)
    
End Sub

Private Sub imgConfirm_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    
    If Not CheckData(False) Then Exit Sub
    
    Dim A As Long
    
    
    UserName = vbNullString

    For A = 0 To lstChars(1).ListCount - 1
        UserName = UserName & lstChars(1).List(A) & "-"
    Next A
    
    If Not lstChars(1).ListCount = 0 Then
    UserName = Left$(UserName, Len(UserName) - 1)
    End If
    
    Account.key = txtKey(0).Text
    Account.Gld = Val(txtGld.Text)
    Account.Eldhir = Val(txtEldhir.Text)
    Account.KeyMao = txtKey(1).Text
    
    'Prepare_And_Connect E_MODO.e_MercaderNewOffer
    Call WriteMercader_NewOffer(False)

    Unload Me
End Sub

Private Function CheckData(ByVal ValidateCode As Boolean) As Boolean
    If Val(txtGld.Text) < 0 Then
        Call MsgBox("¡Monedas de Oro inválidas!")
        Exit Function
    End If
    
    If Val(txtEldhir.Text) < 0 Then
        Call MsgBox("¡Monedas de Eldhir inválidas!")
        Exit Function
    End If
    
   If LenB(txtKey(0).Text) < ACCOUNT_MIN_CHARACTER_KEY Then
        Call MsgBox("La clave de seguridad es un texto con caracteres básicos aleatoreos y contiene 20 caracteres.")
        Exit Function
    End If
    
    If Val(txtGld.Text) = 0 And Val(txtEldhir.Text) = 0 And lstChars(1).ListCount = 0 Then
        Call MsgBox("Debes ofrecer un personaje y/o Monedas")
        Exit Function
    End If
    
    If Not ValidateCode Then
        If LenB(txtKey(1).Text) = 0 Then
            Call MsgBox("Debes escribir el código de seguridad que has recibido vía email.")
            Exit Function
        End If
    End If
    
    CheckData = True
    
End Function


Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Unload Me
End Sub

Private Sub imgValidate_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If Not MainTimer.Check(TimersIndex.Packet500) Then Exit Sub
    If Not CheckData(True) Then Exit Sub
    If tCode.Enabled Then Exit Sub
    
    Dim A As Long
    
    UserName = vbNullString

    For A = 0 To lstChars(1).ListCount - 1
        UserName = UserName & lstChars(1).List(A) & "-"
    Next A
    
    If Not lstChars(1).ListCount = 0 Then
    UserName = Left$(UserName, Len(UserName) - 1)
    End If
    
    Account.key = txtKey(0).Text
    Account.Gld = Val(txtGld.Text)
    Account.Eldhir = Val(txtEldhir.Text)
    Account.KeyMao = txtKey(1).Text
    tCode.Enabled = True
    'Prepare_And_Connect E_MODO.e_MercaderNewOffer
    Call WriteMercader_NewOffer(True)

    imgConfirm.Enabled = True
End Sub

Private Sub imgWeb_Click()
    Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/", vbNullString, vbNullString, 1)
End Sub

Private Sub tCode_Timer()
    Static Second As Long
    
    Second = Second + 1
    
    If Second = 60 Then
        Call MsgBox("Ya puedes volver a enviar otro código.")
        tCode.Enabled = False
        Second = 0
    End If
End Sub
