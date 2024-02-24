Attribute VB_Name = "mAnuncios"
Option Explicit

Public Enum eAnuncios
    A_MISION_COMPLETADA = 1
    
End Enum

Public Type tAnuncios

    Alpha As Byte
    Active As Boolean
    Tittle As String ' Título de los Anuncios
    
    Last As Integer ' Ultimo Texto seleccionado para mostrar
    Text() As String ' Los distintos logros que se va mostrando en pantalla. Cada uno dura 1 segundo
    
    Tipo As eAnuncios
End Type

Public Anuncio As tAnuncios


Public Sub Anuncio_AddNew(ByVal Tittle As String, ByRef Text() As String, ByRef Tipo As eAnuncios)
    
    With Anuncio
        .Active = True
        .Tittle = Tittle
        .Text = Text
        .Last = 1
        .Tipo = Tipo
        .Alpha = 0
    End With
    
    
End Sub

' Actualización de los Anuncios
Public Sub Anuncio_Update_Render()

    Dim X As Long
    Dim Y As Long
    X = 144
    Y = 40
    
    Dim Div As Long
    

    
    
    Dim Mult As Long
    
    ' Anuncios
    If Anuncio.Active Then
        ' @Fondo
        Call Draw_Texture_Graphic_Gui(125, X, Y, To_Depth(7), 544 / 2, 155 / 2, 0, 0, 544 / 2, 155 / 2, ARGB(255, 255, 255, 255 - Anuncio.Alpha), 0, t_Alpha)
            
        ' @Redondel Fondo
        Call Draw_Texture_Graphic_Gui(126, X + 109, Y - 35, To_Depth(7, , , 2), 60, 60, 0, 0, 60, 60, ARGB(255, 255, 255, 255 - Anuncio.Alpha), 0, t_Alpha)
        
        ' @Dibuja el icono según el tipo de anuncio
        Anuncio_Update_Render_Tipo X, Y
        
        ' @Tittle
        Draw_Text f_Booter, 20, X + 140, Y + 18, To_Depth(7, , , 3), 0, ARGB(245, 212, 10, 255 - Anuncio.Alpha), FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_TOP, Anuncio.Tittle, False, True
            
        ' @Texto Vigente
        Draw_Text f_Booter, 18, X + 140, Y + 40, To_Depth(7, , , 3), 0, ARGB(255, 255, 255, 255 - Anuncio.Alpha), FONT_ALIGNMENT_CENTER Or FONT_ALIGNMENT_TOP, Anuncio.Text(Anuncio.Last), False, True
    End If

End Sub

Public Sub Anuncio_Update_Render_Tipo(ByVal X As Long, ByVal Y As Long)
    Call Draw_Texture_Graphic_Gui(127, X + 124, Y - 20, To_Depth(7, , , 3), 32, 32, 0, 0, 32, 32, ARGB(255, 255, 255, 255 - Anuncio.Alpha), 0, t_Alpha)
End Sub

Public Sub Anuncio_Update_Next_Text()

    Static Avance As Long
    
    With Anuncio
    
        If .Active Then
            Avance = Avance + 1
            
            If Avance Mod 120 = 0 Then
                .Last = .Last + 1
                Avance = 0
            End If
            
            If .Last >= UBound(.Text) Then
                .Last = UBound(.Text)
                .Alpha = .Alpha + 1
                
                If .Alpha >= 255 Then
                    Anuncio_Reset
                End If
            End If
        End If
    
    End With
End Sub

Public Sub Anuncio_Reset()
    With Anuncio
        .Active = False
        .Last = 0
        .Tittle = vbNullString
        .Tipo = 0
        .Alpha = 0
    End With
End Sub
