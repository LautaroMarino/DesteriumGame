VERSION 5.00
Begin VB.Form FrmConnect_AccountBig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Perfil de Cuenta"
   ClientHeight    =   16230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   28830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConnect_AccountBig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1082
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1922
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tBlocked 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picUnload 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   28080
      Picture         =   "FrmConnect_AccountBig.frx":000C
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   330
   End
   Begin VB.Timer tUpdate 
      Interval        =   150
      Left            =   180
      Top             =   3045
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   870
      Left            =   9330
      TabIndex        =   0
      Top             =   12150
      Visible         =   0   'False
      Width           =   10245
   End
   Begin VB.Timer tInactive 
      Interval        =   10000
      Left            =   360
      Top             =   3960
   End
   Begin VB.Image imgPlay 
      Height          =   855
      Left            =   24000
      Top             =   14280
      Width           =   3615
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   10
      Left            =   22320
      Top             =   11760
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   9
      Left            =   17280
      Top             =   11760
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   8
      Left            =   11880
      Top             =   11760
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   7
      Left            =   6960
      Top             =   11760
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   6
      Left            =   1920
      Top             =   11760
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   1
      Left            =   1800
      Top             =   7560
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   2
      Left            =   6840
      Top             =   7560
      Width           =   5010
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   3
      Left            =   12120
      Top             =   7560
      Width           =   5010
   End
   Begin VB.Image imgCrearPj 
      Height          =   855
      Left            =   1200
      Top             =   14280
      Width           =   3615
   End
   Begin VB.Image lblPJ 
      Height          =   3900
      Index           =   7
      Left            =   7080
      Top             =   8880
      Width           =   4965
   End
   Begin VB.Image lblPJ 
      Height          =   4020
      Index           =   6
      Left            =   1920
      Top             =   8880
      Width           =   4965
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   5
      Left            =   22320
      Top             =   7560
      Width           =   5010
   End
   Begin VB.Image lblPJ 
      Height          =   3900
      Index           =   5
      Left            =   22440
      Top             =   4680
      Width           =   4965
   End
   Begin VB.Image imgName 
      Height          =   585
      Index           =   4
      Left            =   17160
      Top             =   7560
      Width           =   5010
   End
   Begin VB.Image lblPJ 
      Height          =   3900
      Index           =   4
      Left            =   17160
      Top             =   4680
      Width           =   4965
   End
   Begin VB.Image lblPJ 
      Height          =   3900
      Index           =   3
      Left            =   12120
      Top             =   4680
      Width           =   4965
   End
   Begin VB.Image lblPJ 
      Height          =   3900
      Index           =   2
      Left            =   6960
      Top             =   4680
      Width           =   4965
   End
   Begin VB.Image lblPJ 
      Height          =   4020
      Index           =   1
      Left            =   1920
      Top             =   4680
      Width           =   4965
   End
   Begin VB.Image imgRemove 
      Height          =   705
      Left            =   5160
      Top             =   14400
      Width           =   3555
   End
   Begin VB.Image lblPJ 
      Height          =   3900
      Index           =   8
      Left            =   12120
      Top             =   8880
      Width           =   4965
   End
   Begin VB.Image lblPJ 
      Height          =   4020
      Index           =   9
      Left            =   17280
      Top             =   8880
      Width           =   4965
   End
   Begin VB.Image lblPJ 
      Height          =   4020
      Index           =   10
      Left            =   22440
      Top             =   8880
      Width           =   4965
   End
   Begin VB.Image imgRaza 
      Height          =   675
      Index           =   1
      Left            =   16560
      Top             =   4680
      Width           =   1155
   End
   Begin VB.Image imgGenero 
      Height          =   675
      Index           =   0
      Left            =   4560
      Top             =   4680
      Width           =   1155
   End
   Begin VB.Image imgGenero 
      Height          =   675
      Index           =   1
      Left            =   10080
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Image imgClass 
      Height          =   675
      Index           =   0
      Left            =   9960
      Top             =   5640
      Width           =   1260
   End
   Begin VB.Image imgClass 
      Height          =   675
      Index           =   1
      Left            =   17640
      Top             =   5640
      Width           =   1140
   End
   Begin VB.Image imgRaza 
      Height          =   735
      Index           =   0
      Left            =   11280
      Top             =   4680
      Width           =   915
   End
   Begin VB.Image imgLogin 
      Height          =   1755
      Left            =   1320
      Top             =   14400
      Width           =   1890
   End
   Begin VB.Image imgNewChar 
      Height          =   1005
      Left            =   10440
      Top             =   13800
      Width           =   8175
   End
   Begin VB.Image imgWeb 
      Height          =   2865
      Left            =   10440
      Top             =   600
      Width           =   8775
   End
   Begin VB.Image imgHead 
      Height          =   450
      Index           =   0
      Left            =   19680
      Top             =   4800
      Width           =   705
   End
   Begin VB.Image imgHead 
      Height          =   450
      Index           =   1
      Left            =   21720
      Top             =   4800
      Width           =   705
   End
   Begin VB.Image imgHeading 
      Height          =   345
      Index           =   1
      Left            =   19560
      Top             =   4320
      Width           =   2985
   End
   Begin VB.Image imgMercaderOffer 
      Height          =   975
      Left            =   23880
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Image imgSubClass 
      Height          =   495
      Index           =   1
      Left            =   480
      Top             =   15120
      Width           =   1575
   End
   Begin VB.Image imgSubClass 
      Height          =   495
      Index           =   2
      Left            =   3600
      Top             =   15720
      Width           =   1575
   End
   Begin VB.Image imgSubClass 
      Height          =   495
      Index           =   3
      Left            =   3000
      Top             =   15480
      Width           =   1575
   End
End
Attribute VB_Name = "FrmConnect_AccountBig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TimeButton As Long
Private Sub Form_Load()
    Dim A As Long
    
    g_Captions(eCaption.eCharAccount) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    

    'Call SwitchMap_Copy(1)
    Call Render_CharAccount
    
End Sub

Public Sub SelectedPanelAccount(ByVal Panel As eAccount_PanelSelected)

        '<EhHeader>
        On Error GoTo SelectedPanelAccount_Err

        '</EhHeader>

        Dim A As Long
    
100     Account_PanelSelected = Panel
    
        ' False All
102     For A = lblPJ.LBound To lblPJ.UBound
104         lblPJ(A).visible = False
106     Next A
    
108     For A = imgGenero.LBound To imgGenero.UBound
110         imgGenero(A).visible = False
112     Next A
    
114     For A = imgClass.LBound To imgClass.UBound
116         imgClass(A).visible = False
118     Next A
    
120     For A = imgRaza.LBound To imgRaza.UBound
122         imgRaza(A).visible = False
124     Next A
    
        For A = imgSubClass.LBound To imgSubClass.UBound
            imgSubClass(A).visible = False
        Next A

128    imgRemove.visible = False
        imgCrearPj.visible = False
    
132     For A = imgHead.LBound To imgHead.UBound
134         imgHead(A).visible = False
136     Next A
    
138     For A = imgHeading.LBound To imgHeading.UBound
140         imgHeading(A).visible = False
142     Next A
    
144     txtName.visible = False
146     txtName.Text = vbNullString
148     imgLogin.visible = False
150     imgNewChar.visible = False
152     imgNewChar.visible = False
154     TimeButton = 0
156     imgLogin.Enabled = True
    
        ' End False all
158     Select Case Account_PanelSelected
    
            Case eAccount_PanelSelected.ePanelAccount
160            ' SwitchMap_Copy (85)
            
162             For A = lblPJ.LBound To lblPJ.UBound
164                 lblPJ(A).visible = True
166             Next A

170                 imgRemove.visible = True
imgCrearPj.visible = True

174         Case eAccount_PanelSelected.ePanelAccountCharNew
176             'SwitchMap_Copy (67)
178             UserClase = 1
180             UserRaza = 1
182             UserSexo = 1
184             Account.RenderHeading = E_Heading.SOUTH
186             Call DarCuerpoYCabeza
            
                For A = imgSubClass.LBound To imgSubClass.UBound
                    imgSubClass(A).visible = True
                Next A

188             For A = imgRaza.LBound To imgRaza.UBound
190                 imgRaza(A).visible = True
192             Next A

194             For A = imgClass.LBound To imgClass.UBound
196                 imgClass(A).visible = True
198             Next A

200             For A = imgHead.LBound To imgHead.UBound
202                 imgHead(A).visible = True
204             Next A

206             For A = imgHeading.LBound To imgHeading.UBound
208                 imgHeading(A).visible = True
210             Next A

212             For A = imgGenero.LBound To imgGenero.UBound
214                 imgGenero(A).visible = True
216             Next A
            
218             txtName.visible = True
220             imgNewChar.visible = True

        End Select
    
222     Call Render_CharAccount
        '<EhFooter>
        Exit Sub

SelectedPanelAccount_Err:
        LogError err.Description & vbCrLf & "in ARGENTUM.SelectedPanelAccount " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eCharAccount))
    
    Account.SelectedChar = 0
End Sub

Private Sub imgClass_Click(Index As Integer)

    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0
            If UserClase > 1 Then
                UserClase = UserClase - 1
            Else
                UserClase = NUMCLASES
            End If
    
        Case 1
    
            If UserClase < NUMCLASES Then
                UserClase = UserClase + 1
            Else
                UserClase = 1
            End If
    End Select
    
    SubClass = 0
    Render_CharAccount
End Sub

Private Sub imgCrearPj_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim SlotFree As Integer
        If Account.CharsAmount = ACCOUNT_MAX_CHARS Then
            Call MsgBox("No tienes más espacio para crear un nuevo personaje. Borra alguno o utiliza otra cuenta.")
            Exit Sub
        End If

        SelectedPanelAccount (ePanelAccountCharNew)
End Sub

Private Sub imgGenero_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
    
    If UserSexo = 1 Then
        UserSexo = 2
    Else
        UserSexo = 1
    End If
    
    Call DarCuerpoYCabeza
    Call Render_CharAccount
End Sub

Private Sub imgHead_Click(Index As Integer)
       Call Audio.PlayInterface(SND_CLICK)
    
    If Account.Premium <= 0 Then
        Call MsgBox("¡Debes ser Usuario Tier 1 para tener este beneficio!")
        Exit Sub
    End If
    
    Select Case Index
    
        Case 0
            UserHead = CheckCabeza(UserHead - 1)
        Case 1
            UserHead = CheckCabeza(UserHead + 1)
    End Select
    
End Sub

Private Sub imgHeading_Click(Index As Integer)
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0
            Account.RenderHeading = Account.RenderHeading + 1
        Case 1
            Account.RenderHeading = Account.RenderHeading - 1
    End Select
    
    If Account.RenderHeading < 1 Then Account.RenderHeading = 4
    If Account.RenderHeading > 4 Then Account.RenderHeading = 1
End Sub


Private Sub imgName_Click(Index As Integer)
    If Account.SelectedChar <= 0 Then Exit Sub
    If Account.SelectedChar <> Index Then Exit Sub
    If Account.Chars(Account.SelectedChar).PosMap = 0 Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Temp As String
    Temp = Account.Chars(Index).Name
    
    Temp = InputBox("Escriba su nombre con las minúsculas y/o mayúsculas que desee.", App.Title, Account.Chars(Index).Name)
    
    If UCase$(Account.Chars(Index).Name) = UCase$(Temp) Then
        Account.Chars(Index).Name = Temp
    End If
    
End Sub

Private Sub imgNewChar_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim I         As Integer

    Dim CharAscii As Byte
    
    UserName = txtName.Text
    
    
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
    End If
    
    If Len(UserName) < ACCOUNT_MIN_CHARACTER_CHAR Then
        Call MsgBox("El nombre debe contener más de " & ACCOUNT_MIN_CHARACTER_CHAR & " caracteres.")
        Exit Sub
    End If
    
    If Len(UserName) > ACCOUNT_MAX_CHARACTER_CHAR Then
        Call MsgBox("El nombre debe contener menos de " & ACCOUNT_MAX_CHARACTER_CHAR & " caracteres.")
        Exit Sub
    End If
    

    Prepare_And_Connect E_MODO.e_LoginAccountNewChar, Me
    CreandoPersonaje = True
End Sub



Private Sub imgPlay_Click()
    Call Audio.PlayInterface(SND_CLICK)
    If Account.SelectedChar <= 0 Then Exit Sub
    If Account.Chars(Account.SelectedChar).PosMap = 0 Then Exit Sub
    lblPJ_DblClick (Account.SelectedChar)
End Sub

Private Sub imgRaza_Click(Index As Integer)

    Call Audio.PlayInterface(SND_CLICK)
    
    Select Case Index
    
        Case 0
            If UserRaza > 1 Then
                UserRaza = UserRaza - 1
            Else
                UserRaza = NUMRAZAS
            End If
        Case 1
            If UserRaza < NUMRAZAS Then
                UserRaza = UserRaza + 1
            Else
                UserRaza = 1
            End If
    End Select
    
    Call DarCuerpoYCabeza
    Render_CharAccount
End Sub


Private Sub imgRemove_Click()
    If Account.SelectedChar <= 0 Then Exit Sub
    If Account.Chars(Account.SelectedChar).PosMap = 0 Then Exit Sub
    
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim Elv As Byte
    
    Elv = 29
    
    If Account.Chars(Account.SelectedChar).Elv > Elv Then
        Call MsgBox("No puedes borrar personajes mayores a nivel " & Elv)
        Exit Sub
    End If
    
    If MsgBox("¿Estás seguro que deseas borrar el personaje '" & Account.Chars(Account.SelectedChar).Name & "'?", vbYesNo) = vbYes Then
        
        MirandoCuenta = False
        Dim Text As String
        
        Text = InputBox("Escribe la Clave Pin de tu Cuenta.", "Clave Pin")
          
        If LenB(Text) <= 0 Then
            Call MsgBox("Clave pin inválida.")
            MirandoCuenta = True
            Exit Sub
        End If
    
        Account.Key = Text
        Prepare_And_Connect E_MODO.e_LoginAccountRemove
    
        MirandoCuenta = True
        
    End If
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyEscape Then
        picUnload_Click
        Exit Sub
    End If
    
End Sub


Private Sub imgWeb_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/", vbNullString, vbNullString, 1)
End Sub

Private Sub lblPJ_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)
 
    Account.SelectedChar = Account.Chars(Index).ID
End Sub
Private Sub lblPJ_DblClick(Index As Integer)
    Dim Temp As String
    
    If tBlocked.Enabled Then Exit Sub
    
    Account.SelectedChar = Account.Chars(Index).ID
    Temp = Account.Chars(Index).Name
    
    ' Nuevo Personaje
    If Temp <> vbNullString Then
        UserName = Temp
        
        #If ClienteGM = 1 Then
            Account.Key = InputBox("A continuación escriba la clave pin de la cuenta.")
        #End If
    
        Prepare_And_Connect E_MODO.e_LoginAccountChar
    
    End If
   
   tBlocked.Enabled = True
End Sub


Private Sub picUnload_Click()
   Call Audio.PlayInterface(SND_CLICK)
    
    If Account_PanelSelected = ePrincipal Then
         
        If MsgBox("¿Desea cerrar el juego?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            prgRun = False
        Else

            Exit Sub

        End If
    ElseIf Account_PanelSelected = ePanelAccountRecover Or Account_PanelSelected = ePanelAccountRegister Then
        Call SelectedPanelAccount(ePrincipal)
        Exit Sub
    End If
    
    If IsConnected Then
        If Account_PanelSelected = ePanelAccountCharNew Then
            Call SelectedPanelAccount(ePanelAccount)
            Exit Sub
       
        End If
        
        If MsgBox("¿Estás seguro que deseas cerrar tu cuenta?", vbYesNo) = vbYes Then
            Call Disconnect
            prgRun = False
            'FrmConnect.visible = True
        End If
    End If
End Sub

Private Sub picUnload_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        picUnload_Click
        Exit Sub
    End If
    
End Sub

Private Sub tBlocked_Timer()
    
    tBlocked.Enabled = False
End Sub

Private Sub tInactive_Timer()
    Call WriteUpdateInactive
End Sub

Private Sub tUpdate_Timer()
    If MirandoCuenta Then
        Call Render_CharAccount
    End If
End Sub

Private Sub txtName_Change()
    If txtName.Text <> vbNullString Then
        If Not ValidarNombre(txtName.Text) Then
             txtName.Text = Left(txtName.Text, Len(txtName.Text) - 1)
             txtName.SelStart = Len(txtName.Text)
        End If
    End If
End Sub


'
' Crear Personaje Subs
'
'

Private Function CheckCabeza(ByVal Head As Integer) As Integer

    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    If Head > HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo
                    If Head > ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro
                    If Head > DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano
                    If Head > ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo
                    If Head > GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    UserRaza = 1
                    CheckCabeza = CheckCabeza(Head)
            End Select
        
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    If Head > HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo
                    If Head > ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro
                    If Head > DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano
                    If Head > ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo
                    If Head > GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    UserRaza = 1
                    CheckCabeza = CheckCabeza(Head)
            End Select
        Case Else
            UserSexo = 1
            CheckCabeza = CheckCabeza(Head)
    End Select
End Function

Private Sub DarCuerpoYCabeza()

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer
    
    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select
    
End Sub



