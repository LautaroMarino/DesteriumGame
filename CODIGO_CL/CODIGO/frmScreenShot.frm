VERSION 5.00
Begin VB.Form frmScreenShot 
   BorderStyle     =   0  'None
   Caption         =   "Screenshot de Pantalla para Fotos"
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
   Icon            =   "frmScreenShot.frx":0000
   LinkTopic       =   "Screenshot de Pantalla para Fotos"
   ScaleHeight     =   1082
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1922
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tUpdate 
      Interval        =   250
      Left            =   120
      Top             =   2880
   End
End
Attribute VB_Name = "frmScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

   ' Render_MapGrandev2
    Call wGL_Graphic.Capture(frmScreenShot.hWnd, App.path & Maps_FilePath & "screenshots\" & UserMap & ".png")
    
    Unload Me
End Sub

Private Sub Form_Load()
    g_Captions(eCaption.cMapGrande) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    
    SwitchMap_Copy (UserMap)
    Call Render_MapGrandev2
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cMapGrande))

End Sub

Private Sub tUpdate_Timer()
    'Call RenderScreen_Graphic
    Render_MapGrande
End Sub

Sub RenderScreen_Graphic()

    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cMapGrande))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, Me.ScaleWidth, Me.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    Dim ScreenMinY As Long  'Start Y pos on current screen
    Dim ScreenMaxY As Long  'End Y pos on current screen
    Dim ScreenMinX As Long  'Start X pos on current screen
    Dim ScreenMaxX As Long  'End X pos on current screen

    Dim MinY       As Long  'Start Y pos on current map
    Dim MaxY       As Long  'End Y pos on current map
    Dim MinX       As Long  'Start X pos on current map
    Dim MaxX       As Long  'End X pos on current map

    Dim X          As Long
    Dim Y          As Long

    Dim Drawable   As Long
    Dim DrawableX  As Long
    Dim DrawableY  As Long

    'Figure out Ends and Starts of screen
    ScreenMinY = 1 '(UserPos.Y - AddtoUserPos.Y) - HalfWindowTileHeight
    ScreenMaxY = 100 '(UserPos.Y - AddtoUserPos.Y) + HalfWindowTileHeight
    ScreenMinX = 1 '(UserPos.X - AddtoUserPos.X) - HalfWindowTileWidth
    ScreenMaxX = 100 '(UserPos.X - AddtoUserPos.X) + HalfWindowTileWidth
    
    'Figure out Ends and Starts of map
    MinY = ScreenMinY
    MaxY = ScreenMaxY
    MinX = ScreenMinX
    MaxX = ScreenMaxX
    
    If OffsetCounterY < 0 Then
        MaxY = MaxY + 1
    ElseIf OffsetCounterY > 0 Then
        MinY = MinY - 1
    End If

    If OffsetCounterX < 0 Then
        MaxX = MaxX + 1
    ElseIf OffsetCounterX > 0 Then
        MinX = MinX - 1
    End If
    
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

    For Y = 1 To 100
        DrawableY = (Y - ScreenMinY) * 4
    
        For X = 1 To 100
            DrawableX = (X - ScreenMinX) * 4
            With MapData(X, Y)
                If (.Graphic(1).GrhIndex <> 0) Then
                    Call Draw_Grh(.Graphic(1), DrawableX, DrawableY, To_Depth(1, X, Y), 0, 1, , , , , 4, 4)
                End If
                    
                 If (.Graphic(2).GrhIndex <> 0) Then
                    Call Draw_Grh(.Graphic(2), DrawableX, DrawableY, To_Depth(2, X, Y), 0, 1, , , , , 4, 4)
                End If
                    
                If (.Graphic(3).GrhIndex <> 0) Then
                 Call Draw_Grh(.Graphic(3), DrawableX, DrawableY, To_Depth(3, X, Y), 0, 1, , , , , 4, 4)
                End If
                    
                               If (.Graphic(4).GrhIndex <> 0) Then
                 Call Draw_Grh(.Graphic(4), DrawableX, DrawableY, To_Depth(4, X, Y), 0, 1, , , , , 4, 4)
                End If
            End With
        Next X
    Next Y

     Call wGL_Graphic_Renderer.Flush
End Sub

