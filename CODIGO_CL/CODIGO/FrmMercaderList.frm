VERSION 5.00
Begin VB.Form FrmMercaderList 
   BorderStyle     =   0  'None
   Caption         =   "Mercado"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   Picture         =   "FrmMercaderList.frx":0000
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbClass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2835
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1305
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   540
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.PictureBox PicMercader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   540
      MousePointer    =   99  'Custom
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1740
      Width           =   4245
   End
   Begin VB.Image chkDSP 
      Height          =   225
      Left            =   2040
      Picture         =   "FrmMercaderList.frx":178E6
      Top             =   5310
      Width           =   210
   End
   Begin VB.Image chkChange 
      Height          =   225
      Left            =   2040
      Picture         =   "FrmMercaderList.frx":1873F
      Top             =   5040
      Width           =   210
   End
   Begin VB.Image imgCambio 
      Height          =   255
      Left            =   2280
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Image ImgDsp 
      Height          =   255
      Left            =   2280
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Image ImgOro 
      Height          =   255
      Left            =   720
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Image chkGld 
      Height          =   225
      Left            =   510
      Picture         =   "FrmMercaderList.frx":19598
      Top             =   5310
      Width           =   210
   End
   Begin VB.Image imgSecure 
      Height          =   375
      Left            =   3360
      MousePointer    =   14  'Arrow and Question
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   1560
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Image ButtonView 
      Height          =   375
      Left            =   3480
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Image imgLevel 
      Height          =   255
      Left            =   720
      MousePointer    =   14  'Arrow and Question
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Image imgView 
      Height          =   255
      Left            =   720
      MousePointer    =   14  'Arrow and Question
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Image chkLevel 
      Height          =   225
      Left            =   510
      Picture         =   "FrmMercaderList.frx":1A3F1
      Top             =   5055
      Width           =   210
   End
   Begin VB.Image chkView 
      Height          =   225
      Left            =   510
      Picture         =   "FrmMercaderList.frx":1B24A
      Top             =   4800
      Width           =   210
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmMercaderList.frx":1C0A3
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmMercaderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private clsFormulario          As clsFormMovementManager

Private ListMercader As clsGraphicalList
Private picCheckBox          As Picture
Private picCheckBoxNulo      As Picture

Private OrdenLevel As Boolean
Private ViewFast As Boolean





Private Sub ButtonView_Click()
    
    Call Audio.PlayInterface(SND_CLICK)
    
    If ListMercader.ListIndex = -1 Then
        Call MsgBox("Selecciona una publicación para ver su detalle. Recuerda que puedes activar la vista rápida para visualizar todas las publicaciones de una manera más rápida.", vbInformation)
        Exit Sub
    End If

    If Not MirandoMercader Then
        Call FrmMercaderInfo.Show(, FrmMain)
    Else
        Call FrmMercaderInfo.UpdateInfo
    End If
    
    
End Sub

Private Sub chkGld_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    
    
    'chkLevel.Picture = picCheckBox
    'Set chkLevel.Picture = picCheckBoxNulo
    
    Call MsgBox("En desarrollo")
End Sub
Private Sub chkDsp_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call MsgBox("En desarrollo")
End Sub
Private Sub chkChange_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Call MsgBox("En desarrollo")
End Sub

Private Sub chkView_Click()
     Call Audio.PlayInterface(SND_CLICK)
     
    If ViewFast Then
        ViewFast = False
        Set chkView.Picture = picCheckBoxNulo
    Else
        ViewFast = True
        chkView.Picture = picCheckBox
    End If
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String
    
    GrhPath = DirInterface
    Set picCheckBox = LoadPicture(DirInterface & "options\CheckBoxOpciones.jpg")
    Set picCheckBoxNulo = LoadPicture(DirInterface & "options\CheckBoxOpcionesNulo.jpg")
    
End Sub

Private Sub cmbClass_Click()

    ListMercader.Clear
    
    If cmbClass.ListIndex = 0 Then
        Call Mercader_List_Virgen
    Else
        Call Filter_Class(cmbClass.ListIndex)
        Call Mercader_List
    End If
End Sub

Private Sub Form_Load()
    
    Dim filePath As String
    
    filePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(filePath & "Mercader_View.jpg")
    
    Call LoadButtons
    
    #If ModoBig = 0 Then
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    Set ListMercader = New clsGraphicalList
    
    Call ListMercader.Initialize(PicMercader, RGB(200, 190, 190))
    
    Dim A As Long
    
    For A = 0 To MERCADER_MAX_LIST - 1

        If MercaderList(A).Char > 0 Then
            ListMercader.AddItem Mercader_Prepare_List(A)
        End If
        
    Next A

    
    cmbClass.AddItem "(Ninguna)"
    
    For A = 1 To NUMCLASES
        cmbClass.AddItem ListaClases(A)
    Next A
    
    cmbClass.ListIndex = 0
    
    
    imgView.ToolTipText = "Activa la vista rápida para seleccionar la publicación y visualizarla."
    imgLevel.ToolTipText = "Ordena las publicaciones según el nivel del primer personaje."
    ImgSecure.ToolTipText = "Ofrece DSP. Concreta la venta y le haremos llegar de manera segura el dinero real. (solo AR$)"
    
    If ViewFast Then
        chkView.Picture = picCheckBox
    Else
        Set chkView.Picture = picCheckBoxNulo
    End If

    If OrdenLevel Then
        chkLevel.Picture = picCheckBox
    Else
        Set chkLevel.Picture = picCheckBoxNulo
        
    End If
    
    
    MercaderSelected = 0
    
    Dim Temp As String
    
    If ListMercader.ListIndex > 0 Then
        Temp = ListMercader.List(ListMercader.ListIndex)
        
        MercaderSelected = Val(ReadField(1, Temp, Asc("°")))
        
        If ListMercader.ListCount > 0 Then
            MercaderSelected = MercaderSelected
        End If
    End If
End Sub

Private Function Mercader_Prepare_List(ByVal A As Long) As String
    
    Mercader_Prepare_List = IIf(MercaderUserSlot = A, "[MERCADO_USER]", vbNullString) & MercaderList(A).ID & "º " & MercaderList(A).Chars(1).Desc & IIf(MercaderList(A).Char > 1, " +" & MercaderList(A).Char - 1 & " pjs", vbNullString)
    
End Function

Private Sub imgNew_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    FrmMercaderPublication.Show , FrmMain
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Form_KeyDown vbKeyEscape, 0
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub



Private Sub PicMercader_Click()

    If ListMercader.ListIndex = -1 Then Exit Sub
    
    Dim Temp As String
    
    Temp = ListMercader.List(ListMercader.ListIndex)
    
    Temp = Replace(Temp, "[MERCADO_USER]", "")
    MercaderSelected = Val(ReadField(1, Temp, Asc("°")))
    
    If ViewFast Then
        If Not MirandoMercader Then
            Call FrmMercaderInfo.Show(, FrmMain)
        Else
            Call FrmMercaderInfo.UpdateInfo
        End If
        
        If MirandoOffer Then
            
            MercaderOff = 3
            Call WriteMercader_Required(MercaderOff, MercaderSelected, 0)
        End If
        
    Else
        If MirandoMercader Then
            Call FrmMercaderInfo.UpdateInfo
        End If

        If MirandoOffer Then
             MercaderOff = 3
            Call WriteMercader_Required(MercaderOff, MercaderSelected, 0)
        End If
        
        
    End If
    
End Sub

Private Sub txtSearch_Change()
    
    MercaderID_Selected = 1
    ListMercader.Clear
    
    If Len(txtSearch.Text) <= 0 Then
        If cmbClass.ListIndex = 0 Then
            Call Mercader_List_Virgen
        Else
            Call Filter_Class(cmbClass.ListIndex)
            Call Mercader_List
        End If
    Else
        Call Filter_Mercader(txtSearch.Text)
    End If
End Sub


' # Filtra una publicación a través de un "NOMBRE DE PERSONAJE"
Private Sub Filter_Mercader(ByRef sCompare As String)

    Dim lIndex As Long, b As Long, C As Long
    Dim MaoNull As tMercader
    Dim Slot As Long
    Dim A As Long
    
    For A = 1 To MERCADER_MAX_LIST
        MercaderList(A) = MercaderList_Copy(A)
    Next A
        
    If UBound(MercaderList) <> 0 Then
        For lIndex = 1 To UBound(MercaderList)
            For b = 1 To ACCOUNT_MAX_CHARS
                If InStr(1, UCase$(MercaderList_Copy(lIndex).Chars(b).Name), UCase$(sCompare), vbBinaryCompare) Then
                    Slot = Slot + 1
                    MercaderList(Slot) = MercaderList_Copy(lIndex)
                    ListMercader.AddItem Mercader_Prepare_List(lIndex)
                End If
            Next b
        Next lIndex
    End If
    
End Sub


' # Habilita/Deshabilita el orden por NIVEL.
Private Sub chkLevel_Click()
     Call Audio.PlayInterface(SND_CLICK)
     
    ListMercader.Clear
            
    If OrdenLevel Then
        OrdenLevel = False
        Set chkLevel.Picture = picCheckBoxNulo
        
        If cmbClass.ListIndex = 0 Then
            ' # Listamos la lista sin modificaciones
            Call Mercader_List_Virgen
        Else
            
            'Call Mercader_List_Virgen
            
           ' ListMercader.Clear
            ' # Filtramos por la clase
            Call Filter_Class(cmbClass.ListIndex)
        End If
    Else
        OrdenLevel = True
        chkLevel.Picture = picCheckBox
       ' ListMercader.Clear
        
        ' # Filtramos por la clase
        If cmbClass.ListIndex > 0 Then
            Call Filter_Class(cmbClass.ListIndex)
        Else
            ' # Ordenamos SOLO por Nivel
            Call Chars_OrdenateLevel(True)
        End If
        
       
    End If
   
    
    Mercader_List
End Sub

' # Listamos la lista sin modificaciones
Private Sub Mercader_List_Virgen()

    Dim A As Long
    
    For A = 1 To MERCADER_MAX_LIST
        MercaderList(A) = MercaderList_Copy(A)
    Next A
    
    If OrdenLevel Then
        Call Chars_OrdenateLevel(False)
    End If
        
    For A = 1 To MERCADER_MAX_LIST
    
        If MercaderList(A).Char > 0 Then
            ListMercader.AddItem Mercader_Prepare_List(A)
        End If
    Next A
End Sub

' # Listamos las publicaciones con los filtros aplicados.
Private Sub Mercader_List()

    Dim A As Long, C As Long
    
    For A = 1 To MERCADER_MAX_LIST
        If MercaderList(A).Char > 0 Then
            If cmbClass.ListIndex = 0 Then
                ListMercader.AddItem Mercader_Prepare_List(A)
            Else
                For C = 1 To ACCOUNT_MAX_CHARS
                    If MercaderList(A).Chars(C).Class = cmbClass.ListIndex Then
                        ListMercader.AddItem Mercader_Prepare_List(A)
                        Exit For
                    End If
                Next C
            End If
            
        End If
    Next A
End Sub

' # Filtro por Nivel
Public Sub Chars_OrdenateLevel(ByVal Adding As Boolean)

    Dim A    As Long, b As Long, C As Long
    
    Dim Temp As tMercader
    
    For A = 1 To MERCADER_MAX_LIST - 1
        For b = 1 To MERCADER_MAX_LIST - A

            With MercaderList(b)
                If .Chars(1).Elv < MercaderList(b + 1).Chars(1).Elv Then
                    Temp = MercaderList(b)
                    MercaderList(b) = MercaderList(b + 1)
                    MercaderList(b + 1) = Temp
                End If
            End With
        Next b
    Next A
End Sub

' # Buscar por clase
Private Sub Filter_Class(ByVal Clase As Byte)

    Dim lIndex As Long, b As Long, C As Long
    Dim MaoNull As tMercader
    Dim Slot As Long
    Dim A As Long
    
    For A = 1 To MERCADER_MAX_LIST
       MercaderList(A) = MaoNull
    Next A
        
    If UBound(MercaderList) <> 0 Then
        For lIndex = 1 To UBound(MercaderList)
            
            For b = 1 To ACCOUNT_MAX_CHARS
                If MercaderList_Copy(lIndex).Chars(b).Class = Clase Then
                    Slot = Slot + 1
                    MercaderList(Slot) = MercaderList_Copy(lIndex)
                End If
            Next b
        Next lIndex
    End If

    
    If OrdenLevel Then
        Call Chars_OrdenateLevel(False)
    End If
    
    
End Sub


' Lista Gráfica de Hechizos
Private Sub PicMercader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 0 Then Y = 0
    
    If Y > Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1 Then Y = Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1
    
    If X < PicMercader.ScaleWidth - 10 Then
        ListMercader.ListIndex = Int(Y / ListMercader.Pixel_Alto) + ListMercader.Scroll
        ListMercader.DownBarrita = 0
    
    Else
        ListMercader.DownBarrita = Y - ListMercader.Scroll * (PicMercader.ScaleHeight - ListMercader.BarraHeight) / (ListMercader.ListCount - ListMercader.VisibleCount)
    End If
    
End Sub

Private Sub PicMercader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    Dim yy As Integer
    yy = Y
    
    If yy < 0 Then yy = 0
    
    If yy > Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1 Then yy = Int(PicMercader.ScaleHeight / ListMercader.Pixel_Alto) * ListMercader.Pixel_Alto - 1
    
    If ListMercader.DownBarrita > 0 Then
        ListMercader.Scroll = (Y - ListMercader.DownBarrita) * (ListMercader.ListCount - ListMercader.VisibleCount) / (PicMercader.ScaleHeight - ListMercader.BarraHeight)
    Else
        ListMercader.ListIndex = Int(yy / ListMercader.Pixel_Alto) + ListMercader.Scroll
    End If
ElseIf Button = 0 Then
    ListMercader.ShowBarrita = X > PicMercader.ScaleWidth - ListMercader.BarraWidth * 2
End If
End Sub

Private Sub PicMercader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListMercader.DownBarrita = 0
End Sub


