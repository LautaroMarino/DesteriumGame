VERSION 5.00
Begin VB.Form FrmMenuAccount 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar de Cuenta"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMenuAccount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAccount 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5550
      Left            =   330
      MousePointer    =   99  'Custom
      Picture         =   "FrmMenuAccount.frx":000C
      ScaleHeight     =   370
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   990
      Width           =   4575
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMenuAccount.frx":ABEC
      Top             =   0
      Width           =   330
   End
   Begin VB.Image imgConnect 
      Height          =   495
      Left            =   1680
      Picture         =   "FrmMenuAccount.frx":BC9E
      Top             =   6780
      Width           =   1680
   End
End
Attribute VB_Name = "FrmMenuAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Lista As clsGraphicalList
Private clsFormulario          As clsFormMovementManager


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        #If ModoBig = 1 Then
            dockForm FrmMenu.hWnd, FrmMain.PicMenu, True
            Unload Me
        #Else
            FrmMain.SetFocus
            Unload Me
        #End If
    End If
End Sub

Private Sub Form_Load()

    Dim A As Long
    
    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "account.jpg")
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    Set Lista = New clsGraphicalList
    
    #If ModoBig = 0 Then
        Call Lista.Initialize(picAccount, RGB(255, 255, 255), 32, 120, 20)
    #Else
        Call Lista.Initialize(picAccount, RGB(255, 255, 255), 16, 60, 10)
    #End If
    
    Lista.Clear

    For A = 1 To NUMPASSWD

        With ListPasswd(A)
            Lista.AddItem (.Account)
        End With
            
    Next A
    
    

End Sub

Private Sub imgConnect_Click()
        '<EhHeader>
        On Error GoTo imgConnect_Click_Err
        '</EhHeader>
            
          If Lista.ListIndex < 0 Then Exit Sub
          If Lista.List(Lista.ListIndex) = vbNullString Then Exit Sub
100     If MsgBox("¿Estás seguro que deseas cambiar de cuenta?", vbYesNo) = vbYes Then
102         LogearCuenta = True
        
            LastDataAccount = Lista.List(Lista.ListIndex)
            LastDataPasswd = mDataPasswd.SearchPasswd(LastDataAccount)
            'Call MsgBox("Logear cuenta: " & LastDataAccount & " PW: " & LastDataPasswd)
104         'TempAccount.Email = Lista.List(Lista.ListIndex)
106         'TempAccount.Passwd = mDataPasswd.SearchPasswd(TempAccount.Email)
108         Call WriteQuit(True)

        End If
    
        '<EhFooter>
        Exit Sub

imgConnect_Click_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.FrmMenuAccount.imgConnect_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub picAccount_Click()
    imgConnect_Click
End Sub

Private Sub picAccount_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
        #If ModoBig = 1 Then
            dockForm FrmMenu.hWnd, FrmMain.PicMenu, True
            Unload Me
        #Else
            FrmMain.SetFocus
            Unload Me
  
        #End If
    End If
End Sub

' Lista Gráfica de Hechizos
Private Sub picAccount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Then Y = 0
If Y > Int(picAccount.ScaleHeight / Lista.Pixel_Alto) * Lista.Pixel_Alto - 1 Then Y = Int(picAccount.ScaleHeight / Lista.Pixel_Alto) * Lista.Pixel_Alto - 1
If X < picAccount.ScaleWidth - 10 Then
    Lista.ListIndex = Int(Y / Lista.Pixel_Alto) + Lista.Scroll
    Lista.DownBarrita = 0

Else
    Lista.DownBarrita = Y - Lista.Scroll * (picAccount.ScaleHeight - Lista.BarraHeight) / (Lista.ListCount - Lista.VisibleCount)
End If
End Sub

Private Sub picAccount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    If yy < 0 Then yy = 0
    If yy > Int(picAccount.ScaleHeight / Lista.Pixel_Alto) * Lista.Pixel_Alto - 1 Then yy = Int(picAccount.ScaleHeight / Lista.Pixel_Alto) * Lista.Pixel_Alto - 1
    If Lista.DownBarrita > 0 Then
        Lista.Scroll = (Y - Lista.DownBarrita) * (Lista.ListCount - Lista.VisibleCount) / (picAccount.ScaleHeight - Lista.BarraHeight)
    Else
        Lista.ListIndex = Int(yy / Lista.Pixel_Alto) + Lista.Scroll

      '  If ScrollArrastrar = 0 Then
            'If (Y < yy) Then Lista.Scroll = Lista.Scroll - 1
          '  If (Y > yy) Then Lista.Scroll = Lista.Scroll + 1
        'End If
    End If
ElseIf Button = 0 Then
    Lista.ShowBarrita = X > picAccount.ScaleWidth - Lista.BarraWidth * 2
End If
End Sub

Private Sub picAccount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lista.DownBarrita = 0
End Sub
