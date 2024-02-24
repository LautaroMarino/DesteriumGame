VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private ListPic As PictureBox
Private Lista() As String
Private p_ListIndex As Integer
Private p_ListCount As Integer
Private p_Visible As Boolean
Private p_ForeColor As Long
Private p_Pixel_Alto As Integer
Private p_Scroll As Integer
Private p_OldY As Integer
Private p_DownBarrita As Integer
Private p_ShowBarrita As Boolean
Public VisibleCount As Integer
Public BarraHeight As Integer
Public BarraWidth As Integer

Private lastRendered As Integer

Private CaptionIndex As Integer

Private Sub Class_Initialize()
p_Visible = True
Clear
End Sub

Public Sub Initialize(ByRef LPic As PictureBox, _
                              ByVal ForeColor As Long, _
                              Optional ByVal PixelAlto As Byte = 16, _
                              Optional ByVal Height As Byte = 60, _
                              Optional ByVal Width As Byte = 8, _
                              Optional ByVal Caption As eCaption = 0, _
                              Optional ByVal DrawAvatars As Boolean = False)
    Set ListPic = LPic

    p_ForeColor = ForeColor
    Pixel_Alto = PixelAlto
    BarraHeight = Height
    BarraWidth = Width

    ListPic.ForeColor = ForeColor
    ListPic.Cls
    
    CaptionIndex = Caption
    
End Sub

Public Property Let SetForeColor(ByVal ForeColor As Long)
    p_ForeColor = ForeColor
    Dibujar
End Property
Public Property Get List(ByVal Index As Integer) As String
    List = Lista(Index)
End Property
Public Property Let List(ByVal Index As Integer, ByVal Valor As String)
    Lista(Index) = Valor
    Dibujar
End Property

Public Property Get visible() As Integer
    visible = p_Visible
End Property
Public Property Let visible(ByVal Valor As Integer)
    p_Visible = Valor
    ListPic.visible = p_Visible
End Property

Public Property Get ListIndex() As Integer
    ListIndex = p_ListIndex
End Property
Public Property Let ListIndex(ByVal Valor As Integer)
    If Valor >= p_ListCount Then Valor = p_ListCount - 1
    If p_ListIndex <> Valor Then
        p_ListIndex = Valor
        'If p_ListIndex + Scroll + 1 > VisibleCount Then Scroll = p_ListIndex + Scroll + 1 - VisibleCount
        Dibujar (True)
        lastRendered = Valor
    End If
End Property

Public Property Get OldY() As Integer
    OldY = p_OldY
End Property
Public Property Let OldY(ByVal Valor As Integer)
    p_OldY = Valor
End Property

Public Property Get DownBarrita() As Integer
    DownBarrita = p_DownBarrita
End Property
Public Property Let DownBarrita(ByVal Valor As Integer)
    If p_DownBarrita <> Valor Then
        p_DownBarrita = Valor
        Dibujar
    End If
End Property

Public Property Get Pixel_Alto() As Integer
    Pixel_Alto = p_Pixel_Alto
End Property
Public Property Let Pixel_Alto(ByVal Valor As Integer)
    p_Pixel_Alto = Valor
    VisibleCount = ListPic.ScaleHeight / Valor
End Property

Public Property Get Scroll() As Integer
    Scroll = p_Scroll
End Property

Public Property Get ShowBarrita() As Boolean
    ShowBarrita = p_ShowBarrita
End Property
Public Property Let ShowBarrita(ByVal Valor As Boolean)
    If p_ShowBarrita <> Valor Then
        p_ShowBarrita = Valor
        Dibujar
    End If
End Property
Public Property Let Scroll(ByVal Valor As Integer)
    If Valor > p_ListCount - ListPic.ScaleHeight / p_Pixel_Alto Then Valor = p_ListCount - VisibleCount
    If Valor < 0 Then Valor = 0
    If p_Scroll <> Valor Then
        p_Scroll = Valor
        Dibujar
    End If
End Property

Public Property Get ListCount() As Integer
    ListCount = p_ListCount
End Property

Public Sub AddItem(Item As String)
ReDim Preserve Lista(p_ListCount)
Lista(p_ListCount) = Item
p_ListCount = p_ListCount + 1
If p_ListCount = 1 Then p_ListIndex = 0
Dibujar
End Sub
Public Sub Clear()
ReDim Lista(0)
p_ListCount = 0
p_ListIndex = -1
p_Scroll = 0
If Not ListPic Is Nothing Then Dibujar
End Sub
Private Sub Dibujar(Optional refreshPic As Boolean = True)
Dim I As Integer
Dim PosY As Integer
If refreshPic Then ListPic.Cls
Dim drawText As Boolean
'If Not refreshPic And p_ListIndex = lastRendered Then Exit Sub
For I = p_Scroll To p_Scroll + VisibleCount
    If I < p_ListCount Then
        PosY = (I - p_Scroll) * p_Pixel_Alto
        
        drawText = refreshPic Or lastRendered = I Or p_ListIndex = I

        If Not refreshPic And lastRendered = I And p_ListIndex <> I Then
            ListPic.Line (1, PosY)-(ListPic.ScaleWidth - 1, PosY + p_Pixel_Alto + 1), 0, BF
        End If
    
        Dim prefix As String
        prefix = "[MERCADO_USER]"
        Dim listItem As String
        listItem = Lista(I)
            
        If InStr(1, listItem, prefix) = 1 Then
            listItem = Replace(listItem, prefix, "")  ' Eliminar el prefijo de la lista
         '   Lista(I) = Replace(Lista(A), prefix, vbNullString)
            ListPic.Line (1, PosY)-(ListPic.ScaleWidth - 1, PosY + p_Pixel_Alto + 1), &H800000, BF ' Fondo bordó oscuro
            ListPic.Line (1, PosY)-(ListPic.ScaleWidth - 1, PosY + p_Pixel_Alto + 1), &H400000, B ' Borde bordó más oscuro
        Else
            If I = p_ListIndex Then
                ListPic.Line (1, PosY)-(ListPic.ScaleWidth - 1, PosY + p_Pixel_Alto + 1), &H202020, BF
                ListPic.Line (1, PosY)-(ListPic.ScaleWidth - 1, PosY + p_Pixel_Alto + 1), &H404040, B
                ListPic.ForeColor = &H828282
                
            Else
                ' Cambiar el color del texto después de cada punto
                ListPic.ForeColor = p_ForeColor
                    
            End If
        End If
        
        ListPic.CurrentX = 4
        ListPic.CurrentY = PosY
        
        If drawText Then
            ListPic.Print listItem 'Lista(I)
        End If
    End If
Next I
ListPic.DrawStyle = 0
PosY = 0
If ShowBarrita Then
    If VisibleCount < ListCount Then PosY = p_Scroll * (ListPic.ScaleHeight - BarraHeight) / (p_ListCount - VisibleCount)
    Dim barritaColor As Long
    If DownBarrita Then
        barritaColor = &H505050
    Else
        barritaColor = &H404040
    End If
    ListPic.Line (ListPic.ScaleWidth - BarraWidth, PosY + 1)-(ListPic.ScaleWidth - 1, PosY + BarraHeight - 2), barritaColor, BF
    ListPic.Line (ListPic.ScaleWidth - BarraWidth, PosY + 1)-(ListPic.ScaleWidth - 1, PosY + BarraHeight - 2), &H505050, B
End If
End Sub
