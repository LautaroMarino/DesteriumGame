VERSION 5.00
Begin VB.Form FrmSkin 
   BorderStyle     =   0  'None
   Caption         =   "Skins del Personaje"
   ClientHeight    =   7605
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
   Icon            =   "FrmSkin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tUpdate 
      Interval        =   200
      Left            =   4200
      Top             =   480
   End
   Begin VB.PictureBox PicInv 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   4725
      Left            =   600
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3675
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmSkin.frx":000C
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblBuy 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COMPRAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   3000
      TabIndex        =   1
      Top             =   6840
      Width           =   1800
   End
   Begin VB.Image imgScroll 
      Height          =   375
      Index           =   0
      Left            =   4440
      Top             =   3960
      Width           =   255
   End
   Begin VB.Image imgScroll 
      Height          =   375
      Index           =   1
      Left            =   4440
      Top             =   4440
      Width           =   255
   End
End
Attribute VB_Name = "FrmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkinSelected As Integer


Private Enum eType
        eWeapon = 1
        eArmour = 2
        eShield = 3
        eHelm = 4
End Enum

Private Enum eModo
        eBuy = 1
        eUsage = 2
        eDesequipar = 3
End Enum


Private Modo As eModo
Private SelectedType As eType

Private Skin_ObjIndex As Integer
Private Armadura As Integer
Private Arma As Integer
Private ArmaSecundaria As Integer
Private Escudo As Integer
Private Casco As Integer

Private Const MAX_SKINS_VIEW As Byte = 6

Private TemporalObj As Long

Private Heading As E_Heading

Public MouseX As Long
Public MouseY As Long

Private ObjIndex_Selected As Integer
Private clsFormulario          As clsFormMovementManager


Private ListObj As clsGraphicalList

Private LastList As Integer
Dim CopyList() As tObjData


' # Mueve el Heading del Personaje
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyLeft Then
        Heading = E_Heading.WEST
    ElseIf KeyCode = vbKeyDown Then
        Heading = E_Heading.SOUTH
    ElseIf KeyCode = vbKeyRight Then
        Heading = E_Heading.EAST
    ElseIf KeyCode = vbKeyUp Then
        Heading = E_Heading.NORTH
    End If
    
End Sub
' # Se fija si puede usar el item según su género biologico.
Function SexoPuedeUsarItem(ByVal UserSexo As Byte, _
                           ByVal ObjIndex As Integer, _
                           Optional ByRef sMotivo As String) As Boolean

    On Error GoTo ErrHandler
    
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserSexo <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserSexo <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."

    Exit Function

ErrHandler:
    Call LogError("SexoPuedeUsarItem")
End Function
' # Se fija si la clase tiene que ver el objeto en la lista
Private Function ClasePuedeUsarItem(ByVal ObjIndex As Integer, ByVal Clase As Byte) As Boolean

    Dim A As Long
    
    
    ClasePuedeUsarItem = True

    With ObjData(ObjIndex)
        If .CP_Valid Then
            For A = LBound(.CP) To UBound(.CP)
                 If .CP(A) = Clase Then
                    ClasePuedeUsarItem = False
                    Exit Function
                 End If
                 
            Next A
        End If
    End With
End Function

' # Rellena las skins en el PictureBox
Public Sub Skins_Load()
    
    Dim A As Long
    
    
    If InventorySkins Is Nothing Then
        Set InventorySkins = New clsGrapchicalInventory
    
        Call InventorySkins.Initialize(PicInv, 63, SkinLast, eCaption.eInvSkin1, , , , , , , , , , , , True)
    End If
    
    LastList = 0
    
    ReDim CopyList(0) As tObjData
    
    For A = 1 To NumObjDatas
        With ObjData(A)
            If .Skin > 0 Then
                If ClasePuedeUsarItem(A, UserClase) And SexoPuedeUsarItem(UserSexo, A) Then
                    LastList = LastList + 1
                    
                    ReDim Preserve CopyList(0 To LastList) As tObjData
                    
                    CopyList(LastList) = ObjData(A)
                    'Dim ExistSkin As Integer
                    'ExistSkin = Skin_SearchUser(CopyObjs(A).ID)
                    
                    'If InventorySkins Is Nothing Then
                       ' Call InventorySkins.SetItem(LastList, CopyObjs(A).ID, 1, 0, .GrhIndex, .ObjType, 0, 0, 0, 0, .ValueGLD, "Skin", .ValueDSP, True, 0, 0, 0, 0, , , , , ExistSkin)
                    'Else
                        
                   ' End If
                End If
            
            End If
            
        End With
    Next A
    
    Call Objs_OrdenatePrice(0)
    Call Objs_OrdenatePrice(1)
    
    Dim ExistSkin As Integer
    For A = 1 To LastList
        With CopyList(A)
            ExistSkin = Skin_SearchUser(.ID)

            Call InventorySkins.SetItem(A, .ID, 1, 0, .GrhIndex, .ObjType, 0, 0, 0, 0, .ValueGLD, "Skin", .ValueDSP, True, 0, 0, 0, 0, , , , , ExistSkin)
        End With
    Next A
    
    InventorySkins.DrawInventory
End Sub

' # Ordena los objetos por precio
Public Sub Objs_OrdenatePrice(ByVal Tipo As Byte)

    Dim A    As Long, b As Long
    Dim Temp As tObjData
    
    For A = 1 To LastList - 1
        For b = 1 To LastList - A

            With CopyList(b)
                If Tipo = 0 Then
                    ' # Ordena por Oro
                    If .ValueGLD > CopyList(b + 1).ValueGLD Then
                        Temp = CopyList(b)
                        CopyList(b) = CopyList(b + 1)
                        CopyList(b + 1) = Temp
                    End If
                Else
                    ' # Ordena por DSP
                    If .ValueDSP > CopyList(b + 1).ValueDSP Then
                        Temp = CopyList(b)
                        CopyList(b) = CopyList(b + 1)
                        CopyList(b + 1) = Temp
                    End If
                End If
            End With
        Next b
    Next A
                
End Sub

' # Recorre las diferentes páginas de skins
Private Sub imgScroll_Click(Index As Integer)

    Call Audio.PlayInterface(SND_CLICK)
    
      If (InventorySkins Is Nothing) Then Exit Sub
        
    Select Case Index
    
        Case 0
             InventorySkins.ScrollInventory (False)
        Case 1
            InventorySkins.ScrollInventory (True)
    End Select
End Sub


Private Sub Form_Load()
    
    Dim FilePath As String
    
    FilePath = DirInterface & "menucompacto\"
    Me.Picture = LoadPicture(FilePath & "skins.jpg")
    
    ' Inventario de las skins
    g_Captions(eCaption.eInvSkin1) = wGL_Graphic.Create_Device_From_Display(PicInv.hWnd, PicInv.ScaleWidth, PicInv.ScaleHeight)

    ' Personaje con el SET equipado
    g_Captions(eCaption.eInvSkin2) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    
    Heading = E_Heading.SOUTH
    MirandoSkins = True
    
    #If ModoBig = 0 Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    ReDim CopyList(0) As tObjData
    
    ' # Carga las skins que puede obtener el personaje.
    'Call Skins_Load
    
    ' # Solicita actualizar la lista de skins que tiene el personaje, para ver cual usa y cual no.
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseX = X
    MouseY = Y
    
    
   If MirandoObjetos Then
        FrmObject_Info.Close_Form
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set InvSkin = Nothing
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eInvSkin1))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.eInvSkin2))
    
    MirandoSkins = False
End Sub

Private Sub imgAdd_Click()
    
    If Not MainTimer.Check(TimersIndex.Packet250) Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    If Inventario.SelectedItem = 0 Then Exit Sub
    If InvSkin.SelectedItem = 0 Then Exit Sub
    
   
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Unload Me
End Sub

Private Sub lblBuy_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
    If InventorySkins.SelectedItem = 0 Then Exit Sub
    Dim ObjIndex As Integer
    
    ObjIndex = InventorySkins.ObjIndex(InventorySkins.SelectedItem)
    
    If Modo = eModo.eBuy Then
        If UserGLD < ObjData(ObjIndex).ValueGLD Then
            Call ShowConsoleMsg("¡Oro insuficiente!", 247, 222, 10)
            Exit Sub
        End If
        
        If UserDSP < ObjData(ObjIndex).ValueDSP Then
            Call ShowConsoleMsg("Dsp insuficiente!", 247, 122, 35)
            Exit Sub
        End If
    End If
    
    WriteRequiredSkins ObjIndex, Modo

End Sub

Private Sub PicInv_Click()
    If Not MainTimer.Check(TimersIndex.Packet250) Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    Dim ObjIndex As Integer
    
    If InventorySkins.SelectedItem = 0 Then Exit Sub
    
    ObjIndex = InventorySkins.ObjIndex(InventorySkins.SelectedItem)
    If ObjIndex = 0 Then Exit Sub
    
    If Skin_SearchExist(ObjIndex) Then
        If Skins_CheckingItems(ObjIndex) Then
            Modo = eDesequipar
            lblBuy.Caption = "DESEQUIPAR"
        Else
            Modo = eUsage
            lblBuy.Caption = "USAR"
        End If
    Else
        Modo = eBuy
        lblBuy.Caption = "COMPRAR"
    End If
    
    Skin_DeterminateTipo ObjIndex
    
    ShowConsoleMsg "Objeto: " & ObjIndex & "(" & InventorySkins.ItemName(InventorySkins.SelectedItem) & ")"
    
    
End Sub

Private Function Skin_DeterminateTipo(ByVal ObjIndex As Integer)
    With ObjData(ObjIndex)
        Select Case .ObjType
            Case eOBJType.otarmadura
                Armadura = .Anim
                
            Case eOBJType.otWeapon
                If .Proyectil > 0 Then
                    Arma = .Anim
                Else
                    ArmaSecundaria = .Anim
                End If
                
                
            Case eOBJType.otcasco
                Casco = .Anim
                
            
            Case eOBJType.otescudo
                Escudo = .Anim
        End Select
    End With
End Function
Private Function Skin_SearchExist(ByVal ObjIndex As Integer) As Boolean
    
    Dim A As Long
    
    For A = 1 To ClientInfo.Skin.Last
        If ClientInfo.Skin.ObjIndex(A) = ObjIndex Then
            Skin_SearchExist = True
            Exit Function
        End If
    Next A
End Function

Private Sub PicInv_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub


Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub tUpdate_Timer()
    InventorySkins.DrawInventory
End Sub

