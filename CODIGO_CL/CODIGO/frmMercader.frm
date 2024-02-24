VERSION 5.00
Begin VB.Form frmMercader 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Mercado de Personajes"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmMercader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tSendCode 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9870
      Top             =   1365
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   9450
      TabIndex        =   2
      Top             =   7395
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   9450
      TabIndex        =   1
      Top             =   6795
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox txtGld 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   195
      Left            =   9345
      TabIndex        =   0
      Text            =   "1"
      Top             =   8130
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Timer tUpdate 
      Interval        =   175
      Left            =   1785
      Top             =   1995
   End
   Begin VB.Image imgSendCode 
      Height          =   435
      Left            =   9480
      Top             =   6000
      Width           =   1995
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   10
      Left            =   8190
      Top             =   5250
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   9
      Left            =   6720
      Top             =   5250
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   8
      Left            =   5250
      Top             =   5250
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   7
      Left            =   3885
      Top             =   5355
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   6
      Left            =   2415
      Top             =   5250
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   5
      Left            =   8190
      Top             =   3255
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   4
      Left            =   6720
      Top             =   3150
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   3
      Left            =   5355
      Top             =   3255
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   2
      Left            =   3885
      Top             =   3255
      Width           =   750
   End
   Begin VB.Image imgChar 
      Height          =   1170
      Index           =   1
      Left            =   2415
      Top             =   3255
      Width           =   750
   End
   Begin VB.Image imgBlocked 
      Height          =   1065
      Left            =   5145
      Top             =   7770
      Width           =   1590
   End
   Begin VB.Image imgSelected 
      Height          =   1065
      Index           =   5
      Left            =   11040
      Top             =   4920
      Width           =   750
   End
   Begin VB.Image imgSelected 
      Height          =   1065
      Index           =   4
      Left            =   7980
      Top             =   5460
      Width           =   750
   End
   Begin VB.Image imgSelected 
      Height          =   1305
      Index           =   2
      Left            =   1575
      Top             =   5850
      Width           =   960
   End
   Begin VB.Image imgSelected 
      Height          =   540
      Index           =   0
      Left            =   4515
      Top             =   630
      Width           =   2955
   End
   Begin VB.Image imgUnload 
      Height          =   465
      Left            =   11550
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "frmMercader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MouseX                   As Long
Public MouseY                   As Long


Private Sub Form_Load()
    Dim A As Long
    
    
    'MercaderSelected = ePanelInitial
    
    For A = 1 To ACCOUNT_MAX_CHARS
        MercaderUser.bChars(A) = 0
        'MercaderUser.IDCHARS(A) = 0
    Next A
    
    MercaderGld = 0
    MercaderUser.Gld = 0
    MercaderUser.Char = 0
    MercaderUser.Blocked = 0
    g_Captions(eCaption.cMercader) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    
    EsperandoValidacion = False
    Call CambioSolapa(MercaderSelected)
    Call Render_Mercader
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cMercader))

End Sub

Private Sub imgBlocked_Click()
    If MercaderSelected <> eMercaderSelected.ePanelPublication And MercaderSelected <> eMercaderSelected.ePanelOffer Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    If MercaderUser.Blocked = 1 Then
        MercaderUser.Blocked = 0
    Else
        MercaderUser.Blocked = 1
    End If
    
End Sub

Private Function Searching(ByVal Name As String) As Integer
    Dim A As Long
    
    For A = 1 To ACCOUNT_MAX_CHARS
        If UCase$(Account.Chars(A).Name) = Name Then
            Searching = A
            Exit Function
        End If
    Next A
End Function
Private Sub imgChar_Click(Index As Integer)
    If MercaderSelected <> eMercaderSelected.ePanelPublication And MercaderSelected <> eMercaderSelected.ePanelOffer Then Exit Sub
    If Account.Chars(Index).PosMap = 0 Then Exit Sub
    
    Dim ID As Integer
    ID = Index
    
    Call Audio.PlayInterface(SND_CLICK)
    
    If Account.Chars(Index).Elv < MERCADER_MIN_LVL Then
        Call MsgBox("¡Solo puedes ofrecer personajes superiores a Nivel " & MERCADER_MIN_LVL & "!")
        Exit Sub
    End If
     

    If MercaderUser.bChars(ID) = 0 Then
        MercaderUser.IDCHARS(Account.Chars(Index).ID) = 1
        MercaderUser.bChars(ID) = 1
        MercaderUser.Char = MercaderUser.Char + 1
    Else
        MercaderUser.IDCHARS(Account.Chars(Index).ID) = 0
        MercaderUser.bChars(ID) = 0
        MercaderUser.Char = MercaderUser.Char - 1
        
        
    End If
    
    MercaderGld = CalculatePrice
End Sub

Private Function CalculatePrice() As Long
    Dim A As Long
    Dim Temp As Long
    
    For A = 1 To ACCOUNT_MAX_CHARS
        If MercaderUser.bChars(A) = 1 Then
                Temp = Temp + (MERCADER_GLD_SALE * Account.Chars(A).Elv)
        End If
    Next A
    
    If Account.Premium > 2 Then Temp = 0
    CalculatePrice = Temp
End Function

Private Sub Mercader_Send_Publication(ByVal ValidCode As Boolean, ByVal IsOffer As Integer)
    
    
     
    If Not IsOffer > 0 Then
        If MercaderUser.Char = 0 Then
            Call MsgBox("¡Antes de realizar una publicación selecciona al menos un personaje.")
            Exit Sub
        End If
    End If
    
    If Val(txtGld.Text) < 0 Then txtGld.Text = "0"

    If Not ValidCode Then
        If Len(txtCode.Text) <= 0 Then
            Call MsgBox("Debes ingresar el código que has recibido vía email.")
            Exit Sub
        End If
        
       
    End If

    Account.Key = txtKey.Text
    Account.KeyMao = txtCode.Text
    MercaderUser.Gld = Val(txtGld.Text)
    
    Call WriteMercader_New(IsOffer, MercaderUser)
End Sub
Private Sub imgSelected_Click(Index As Integer)
    If MercaderSelected = eMercaderSelected.ePanelPublication Or MercaderSelected = eMercaderSelected.ePanelOffer Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Call CambioSolapa(Index)
End Sub
Private Sub CambioSolapa(ByVal Index As Integer)
   MercaderSelected = Index
    
    txtGld.visible = False

    Dim A As Long
    
   ' imgSelected(5).visible = True
  '  imgSelected(0).visible = True
  '  imgSelected(2).visible = True
        
 '   For A = imgChar.LBound To imgChar.UBound
   '     imgChar(A).Enabled = False
  '  Next A
    
    txtKey.visible = False
    txtCode.visible = False
    
    For A = 1 To ACCOUNT_MAX_CHARS
        MercaderUser.bChars(A) = 0
    Next A
    
    MercaderGld = 0
    MercaderUser.Gld = 0
    MercaderUser.Char = 0
    
    Select Case MercaderSelected
        
        Case eMercaderSelected.ePanelMercaderList
            
        Case eMercaderSelected.ePanelSearch
            Call WriteMercader_Required(1, 1, 255)
        
        Case eMercaderSelected.ePanelPublication, eMercaderSelected.ePanelOffer

            txtGld.visible = True
                
           ' imgSelected(0).visible = False
         '   imgSelected(2).visible = False
            
          '  For A = imgChar.LBound To imgChar.UBound
              '  imgChar(A).Enabled = True
         '   Next A
 
            txtKey.visible = True
            txtCode.visible = True
        Case Else
            txtGld.visible = False
    End Select
End Sub

Private Sub imgSendCode_Click()

    If MercaderSelected <> eMercaderSelected.ePanelPublication And MercaderSelected <> eMercaderSelected.ePanelOffer Then Exit Sub
    
            If Account.SaleSlot > 0 And Not MercaderSelected = eMercaderSelected.ePanelOffer Then
                Call MsgBox("¡Ya tienes una publicación vigente!")
                CambioSolapa 0
                Exit Sub
            End If
            
    If MercaderUser.Char > 1 Then
        If Account.Premium < 3 Then
            Call MsgBox("Debes ser al menos Tier 3 para poder publicar mas de 1 personaje. ¡Mientras tanto podrás vender u ofrecer de a uno!")
            Exit Sub
    
        End If
    End If
    
    If Len(txtKey.Text) <= 0 Then
        Call MsgBox("Debes ingresar la clave de seguridad que has recibido al momento de registrarte. Luego solicitar un código vía EMAIL, ingresarlo y presiona el botón nuevamente.")
        Exit Sub

    End If
    
    If MercaderGld > Account.Gld Then
            Call MsgBox("Realizar esta operación requiere tener en la cuenta " & Format(MercaderGld, "##,##") & " Monedas de Oro y al parecer tu no dispones de eso. Si crees que no es así, relogea tu personaje para actualizar el oro en tu cuenta...")
            Exit Sub

        End If
        
    If MercaderSelected = eMercaderSelected.ePanelOffer Then

        If MercaderList_Copy(MercaderID).Gld > Account.Gld Then
            Call MsgBox("La publicación requiere un mínimo de " & Format(MercaderList_Copy(MercaderID).Gld, "##,##") & " Monedas de Oro y al parecer tu no dispones de eso. Si crees que no es así, relogea tu personaje para actualizar el oro en tu cuenta...")
            Exit Sub

        End If

    End If
    
    If EsperandoValidacion Then

        ' Confirma la publicación
        If MercaderSelected = eMercaderSelected.ePanelPublication Then

            If MsgBox("¿Estás seguro que deseas realizar la publicación. ¡Fijate muy bien que estás publicando!", vbYesNo) = vbYes Then
                Call Mercader_Send_Publication(False, 0)

            End If
            
        ElseIf MercaderSelected = eMercaderSelected.ePanelOffer Then

            If MsgBox("¿Estás seguro que deseas ofrecer a la publicación seleccionada? ¡Fijate que es lo que ofertas! No nos haremos responsables...", vbYesNo) = vbYes Then
                Call Mercader_Send_Publication(False, MercaderID)

            End If

        End If

    Else

        If tSendCode.Enabled Then
            Call MsgBox("Ya has enviado un código recientemente. Espera algunos segundos y vuelve a intentarlo")

        End If
    
        Call Mercader_Send_Publication(True, MercaderID)
        
    End If

End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim A As Long
    
    If MercaderSelected = eMercaderSelected.ePanelPublication Then
        MercaderSelected = eMercaderSelected.ePanelInitial
        imgSelected_Click (0)
        EsperandoValidacion = False
        Exit Sub

    End If
    
    If FrmConnect_Account.visible Then
            FrmConnect_Account.SetFocus
    Else
        FrmMain.SetFocus

    End If
    
    Unload Me

End Sub

Private Sub tSendCode_Timer()
    tSendCode.Enabled = False
End Sub

Private Sub tUpdate_Timer()
    Call Render_Mercader
End Sub

Private Sub txtGld_Change()
    
    If Not IsNumeric(txtGld.Text) Then
        txtGld.Text = "0"
    End If
    
    If Val(txtGld.Text) < 1 Then
        txtGld.Text = "1"
    End If
    
    If Val(txtGld.Text) > MERCADER_MAX_GLD Then
        txtGld.Text = MERCADER_MAX_GLD
        txtGld.SelStart = Len(txtGld.Text)
    End If
End Sub

