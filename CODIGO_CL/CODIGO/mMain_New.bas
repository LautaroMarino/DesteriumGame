Attribute VB_Name = "mMain_New"
Option Explicit


Public Type tConfigMain
    ViewStats As Boolean
    ViewCounters As Boolean
    ViewInventory As Boolean
    
End Type


Public ConfigMain As tConfigMain



Public Sub RenderScreen_Stats()
    
    Dim X As Long
    Dim Y As Long
    
    Dim Width As Long
    Dim Height As Long
    
    Dim Div As Byte ' División para usar los gráficos grandes
    
    #If ModoBig = 1 Then
        Div = 1
    #Else
        Div = 2
    #End If
    
    ' Graficos
    
    ' x2
    X = 391
    Y = 805
    
    ' x1
    X = 391 / 2
    Y = 489
    Width = 820 / 2
    Height = 226 / 2
    
    Call Draw_Texture_Graphic_Gui(138, X, Y, To_Depth(9), Width, Height, 0, 0, Width * 2, Height * 2, -1, 0, eTechnique.t_Alpha)
    
    
End Sub

