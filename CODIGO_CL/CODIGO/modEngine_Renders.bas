Attribute VB_Name = "Engine_Renders"
Option Explicit

'############################
'############################ RENDER CLANES ##########################
'############################
Private Const MAX_COMPROBACIONES As Byte = 4

'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private ContadorMacroClicks(1 To MAX_COMPROBACIONES) As Position




' Detecta El Colour segun alineation del clan
Public Function Guilds_Alineation_Colour(ByVal Alineation As eGuildAlineation) As Long
    Select Case Alineation
            
        Case eGuildAlineation.a_Armada
            Guilds_Alineation_Colour = ARGB(18, 180, 190, 255)
            
        Case eGuildAlineation.a_Legion
            Guilds_Alineation_Colour = ARGB(222, 120, 120, 255)
            
        Case eGuildAlineation.a_Neutral
            Guilds_Alineation_Colour = ARGB(190, 190, 190, 255)
        
        Case Else
            Guilds_Alineation_Colour = ARGB(190, 190, 190, 255)
        
    End Select
End Function
Public Function Guilds_Alineation_Text(ByVal Alineation As eGuildAlineation) As String
    Select Case Alineation
    
        Case eGuildAlineation.a_Armada
            Guilds_Alineation_Text = "Ciudadano"
            
        Case eGuildAlineation.a_Legion
            Guilds_Alineation_Text = "Criminal"
            
        Case eGuildAlineation.a_Neutral
            Guilds_Alineation_Text = "Neutral"
            
    End Select
End Function


' Creado para sacar Screenshots de Mapas Mega Grandes
Public Sub Render_MapGrande()
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cMapGrande))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, frmScreenShot.ScaleWidth, frmScreenShot.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    Dim X As Long, Y As Long
    
    For X = 1 To 100
        For Y = 1 To 100
            If InMapBounds(UserPos.X + X, UserPos.Y + Y) Then
                With MapData_Copy(UserPos.X + X, UserPos.Y + Y)
                    If .Graphic(1).GrhIndex <> 0 Then _
                       Draw_Grh .Graphic(1), (X - 1) * 32, (Y - 1) * 32, To_Depth(1), 1, 1, , _
                       -1, , eTechnique.t_Alpha, GrhData(.Graphic(1).GrhIndex).pixelWidth, GrhData(.Graphic(1).GrhIndex).pixelHeight
                            
                    If .Graphic(2).GrhIndex <> 0 Then _
                       Draw_Grh .Graphic(2), (X - 1) * 32, (Y - 1) * 32, To_Depth(2, X, Y), 1, 1, , -1, , eTechnique.t_Alpha, GrhData(.Graphic(2).GrhIndex).pixelWidth, GrhData(.Graphic(2).GrhIndex).pixelHeight
                                
                    If .Graphic(3).GrhIndex <> 0 Then _
                       Draw_Grh .Graphic(3), (X - 1) * 32, (Y - 1) * 32, To_Depth(3, X, Y), 1, 1, , -1, , eTechnique.t_Alpha, GrhData(.Graphic(3).GrhIndex).pixelWidth, GrhData(.Graphic(3).GrhIndex).pixelHeight
                               
                     If .Graphic(4).GrhIndex <> 0 Then _
                      Draw_Grh .Graphic(4), (X - 1) * 32, (Y - 1) * 32, To_Depth(4, X, Y), 1, 1, , -1, , eTechnique.t_Alpha, GrhData(.Graphic(4).GrhIndex).pixelWidth, GrhData(.Graphic(4).GrhIndex).pixelHeight
                End With
            End If
        Next Y
    Next X
    Call wGL_Graphic_Renderer.Flush
   
End Sub
Public Sub Render_MapGrandev2()
Call wGL_Graphic.Use_Device(g_Captions(eCaption.cMapGrande))
Call wGL_Graphic_Renderer.Update_Projection(&H0, frmScreenShot.ScaleWidth, frmScreenShot.ScaleHeight)
Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
Dim X As Long, Y As Long
    
For X = 1 To 100
    For Y = 1 To 100
            With MapData_Copy(X, Y)
                If .Graphic(1).GrhIndex <> 0 Then _
                   Draw_Grh .Graphic(1), (X - 1) * 32, (Y - 1) * 32, To_Depth(1), 1, 1, , _
                   -1, , eTechnique.t_Alpha
                            
            '    If .Graphic(2).GrhIndex <> 0 Then _
                   Draw_Grh .Graphic(2), (X - 1) * 32, (Y - 1) * 32, To_Depth(2, X, Y), 1, 1, , -1, , eTechnique.t_Alpha
                                
                'If .Graphic(3).GrhIndex <> 0 Then _
                   Draw_Grh .Graphic(3), (X - 1) * 32, (Y - 1) * 32, To_Depth(3, X, Y), 1, 1, , -1, , eTechnique.t_Alpha
                               
               '  If .Graphic(4).GrhIndex <> 0 Then _
                  Draw_Grh .Graphic(4), (X - 1) * 32, (Y - 1) * 32, To_Depth(4, X, Y), 1, 1, , -1, , eTechnique.t_Alpha
            End With
    Next Y
Next X

Call wGL_Graphic_Renderer.Flush
   
End Sub

Public Sub Render_CharAccount()

        '<EhHeader>
        On Error GoTo Render_CharAccount_Err

        '</EhHeader>
    
100     Call wGL_Graphic.Use_Device(g_Captions(eCaption.eCharAccount))
        Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmConnect_Account.ScaleWidth, FrmConnect_Account.ScaleHeight)
104     Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
        Dim X            As Long, Y As Long, Map As Integer

        Dim A            As Long

        Dim OffsetY      As Integer

        Dim x1           As Long

        Dim y1           As Long

        Dim ColourText   As Long
    
        Dim InitialX     As Integer

        Dim InitialY     As Integer
        
106

        Dim AlphaTexture As Long

        Dim Pixel        As Byte

        Dim Divition     As Byte

108     'Draw_Text f_Medieval, 75, 400, 25, To_Depth(5), 0, ARGB(250, 157, 30, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Desterium AO", True
110      'Draw_Text f_Medieval, 75, 402, 27, To_Depth(5), 0, ARGB(255, 255, 255, 50), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Desterium AO", False
    
112     AlphaTexture = ARGB(255, 255, 255, 255)
    
114     Select Case Account_PanelSelected
        
                ' Principal
            Case eAccount_PanelSelected.ePrincipal, eAccount_PanelSelected.ePanelAccountRecover, eAccount_PanelSelected.ePanelAccountRegister
116             Draw_Text f_Medieval, 30, 400, 65, To_Depth(7), 0, ARGB(250, 255, 255, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "V4.0.0", False
118             InitialX = 30
120             InitialY = 25
122             AlphaTexture = ARGB(255, 255, 255, 255)
            
124             Call Render_Principal
            
126             Pixel = 16
128             Divition = 2
        
                ' Panel de la Cuenta (Lista de Personajes)
130         Case eAccount_PanelSelected.ePanelAccount
                #If ModoBig = 0 Then
                    
                    Call Draw_Texture_Graphic_Gui(110, 0, 0, To_Depth(1), 800, 600, 0, 0, 800, 600, -1, 0, eTechnique.t_Alpha)
                #Else
                    Call Draw_Texture_Graphic_Gui(109, 0, 0, To_Depth(1), 1920, 1080, 0, 0, 1920, 1080, -1, 0, eTechnique.t_Alpha)
                #End If
        
132             'Draw_Text f_Medieval, 25, 950, 180, To_Depth(6), 0, ARGB(250, 255, 255, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Panel de la Cuenta", False
            
134             InitialX = 59
136             InitialY = 60
138             Call Render_Account
140             AlphaTexture = ARGB(255, 255, 255, 200)
            
142             Pixel = 32
144             Divition = 1

                ' Creación de un Nuevo Personaje
146         Case eAccount_PanelSelected.ePanelAccountCharNew

                #If ModoBig = 0 Then
                    Call Draw_Texture_Graphic_Gui(107, 0, 0, To_Depth(1), 800, 600, 0, 0, 800, 600, -1, 0, eTechnique.t_Alpha)
                #Else
                    Call Draw_Texture_Graphic_Gui(108, 0, 0, To_Depth(1), 1920, 1080, 0, 0, 1920, 1080, -1, 0, eTechnique.t_Alpha)
                #End If
            
150             InitialX = 36
152             InitialY = 45
154             AlphaTexture = ARGB(255, 255, 255, 255)
            
156             Call Render_NewChar
            
158             Pixel = 32
160             Divition = 1

        End Select
    
180     Call wGL_Graphic_Renderer.Flush
   
        '<EhFooter>
        Exit Sub

Render_CharAccount_Err:
        LogError err.Description & vbCrLf & "in ARGENTUM.Engine_Renders.Render_CharAccount " & "at line " & Erl

        Resume Next

        '</EhFooter>
End Sub

Private Sub Render_Principal()
    
    Dim ColourText As Long
    
    ' Recuadro de Email y Contraseña
    Call Draw_Texture_Graphic_Gui(47, 270, 446, To_Depth(4), 266, 106, 0, 0, 266, 106, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)

    Call Draw_Texture_Graphic_Gui(49, 305, 464, To_Depth(5), 200, 40, 0, 0, 200, 40, ARGB(255, 255, 255, 200), 0, eTechnique.t_Alpha)
    
    
    ' Email y Contraseña + Boton que dice Conectarse
    Draw_Text f_Tahoma, 14, 321, 460, To_Depth(6), 0, ARGB(250, 255, 255, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Email", False
    
    
    ' Button Recuperar
    ' Seleccionado
    If Account_PanelSelected = ePanelAccountRecover Then
        Call Draw_Texture_Graphic_Gui(11, 592, 310, To_Depth(4), 128, 128, 0, 0, 128, 128, ARGB(100, 100, 100, 200), 0, eTechnique.t_Alpha)
        Draw_Text f_Tahoma, 14, 400, 510, To_Depth(6), 0, ARGB(250, 255, 255, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Recibiras un correo electronico", False
    ElseIf Account_PanelSelected = ePanelAccountRegister Then
        Call Draw_Texture_Graphic_Gui(11, 633, 450, To_Depth(4), 128, 128, 0, 0, 128, 128, ARGB(100, 100, 100, 200), 0, eTechnique.t_Alpha)
    Else
    
        ' Boton Contraseña
        Draw_Text f_Tahoma, 14, 321, 504, To_Depth(6), 0, ARGB(250, 255, 255, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Contraseña", False
        Call Draw_Texture_Graphic_Gui(49, 305, 507, To_Depth(5), 200, 40, 0, 0, 200, 40, ARGB(255, 255, 255, 200), 0, eTechnique.t_Alpha)
    End If
    
    Draw_Text f_Tahoma, 14, 655, 410, To_Depth(6), 0, ARGB(200, 200, 200, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Lautaro", True
    Draw_Text f_Tahoma, 14, 655, 420, To_Depth(6), 0, ARGB(250, 157, 30, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "<Recuperar>", True
    
    ' Boton de Conectarse
    Call Draw_Texture_Graphic_Gui(11, 90, 370, To_Depth(4), 128, 128, 0, 0, 128, 128, ARGB(100, 100, 100, 200), 0, eTechnique.t_Alpha)
    ' ADAPTAR CLASICO
    'Call Draw_Texture(56146, 130, 420, To_Depth(5), GrhData(56146).pixelWidth, GrhData(56146).pixelHeight, -1, 0, t_Alpha)
    Draw_Text f_Tahoma, 14, 155, 480, To_Depth(6), 0, ARGB(200, 200, 200, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Bowser", True
    Draw_Text f_Tahoma, 14, 155, 490, To_Depth(6), 0, ARGB(250, 157, 30, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "<Conectar>", True
    
    ' Boton de Registro
    ' ADAPTAR CLASICO
    'Call Draw_Texture(56135, 680, 480, To_Depth(5), GrhData(56135).pixelWidth, GrhData(56135).pixelHeight, -1, 0, t_Alpha)
    Draw_Text f_Tahoma, 14, 700, 540, To_Depth(6), 0, ARGB(200, 200, 200, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Ursula", True
    Draw_Text f_Tahoma, 14, 700, 550, To_Depth(6), 0, ARGB(250, 157, 30, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "<Registrarse>", True
    
    ' Grietas
    ' ADAPTAR CLASICO
    'Call Draw_Texture(41707, 40, 140, To_Depth(3), GrhData(41707).pixelWidth, GrhData(41707).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 40, 350, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 40, 75, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 40, 140, To_Depth(3), GrhData(41707).pixelWidth, GrhData(41707).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 40, 0, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 150, 0, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 300, 0, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 350, 0, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 450, 0, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 550, 0, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 550, 100, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 550, 250, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 550, 450, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 550, 550, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 300, 380, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41708, 220, 380, To_Depth(3), GrhData(41708).pixelWidth, GrhData(41708).pixelHeight, -1, 0, t_Alpha)
   ' Call Draw_Texture(41707, 500, 400, To_Depth(4), GrhData(41707).pixelWidth, GrhData(41707).pixelHeight, -1, 0, t_Alpha)
        
End Sub

Private Sub Render_NewChar()
    
    Dim Body As Integer

    Dim Head As Integer
    
    Dim X    As Long

    Dim Y    As Long

    Dim Mult As Long
    Dim Width As Integer
    Dim Height As Integer
    
        Dim BasePng As Integer
        
        #If ModoBig > 0 Then
            BasePng = 114
            Mult = 2
            Width = 689
            Height = 375
        #Else
            BasePng = 67
            Mult = 1
            Width = 287
            Height = 208
        #End If
        
    '
    '
    ' BUTTON <GENERO>
    
    #If ModoBig = 0 Then
        X = 190
        Y = 175
    #Else
        X = 490
        Y = 310
    #End If

    If UserSexo = 2 Then
        Call Draw_Text(eFonts.f_Morpheus, 22, X + 24, Y, To_Depth(7), 0, ARGB(230, 250, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Mujer", True)
    Else
        Call Draw_Text(eFonts.f_Morpheus, 22, X + 24, Y, To_Depth(7), 0, ARGB(230, 250, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Hombre", True)

    End If
  
    '
    '
    ' BUTTON <CLASE>
    #If ModoBig = 0 Then
        X = 375
        Y = 210
    #Else
        X = 900
        Y = 375
    #End If

    If UserClase > 0 Then
        Call Draw_Text(eFonts.f_Morpheus, 22, X + 24, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, ListaClases(UserClase), True)
        Call Draw_Texture_Graphic_Gui(35 + UserClase, 80 * Mult, 240 * Mult, To_Depth(7), 123 * Mult, 184 * Mult, 0, 0, 123 * Mult, 184 * Mult, ARGB(255, 255, 255, 100), 0, eTechnique.t_Alpha)
        
        #If ModoBig = 0 Then
            X = 375
            Y = 490
        #Else
            X = 734
            Y = 880
        #End If

        

        Call Draw_Texture_Graphic_Gui(BasePng + UserClase, X - 116, (Y / 2) - 6, To_Depth(8), Width, Height, 0, 0, Width, Height, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)


    End If

    '
    '
    '
    ' BUTTON <RAZA>
    
    #If ModoBig = 0 Then
        X = 375
        Y = 175
    #Else
        X = 940
        Y = 310
    #End If
    
    ' Titulo 'Raza'
    '   Call Draw_Text(eFonts.f_Medieval, 22, X + 20, Y - 18, To_Depth(7), 0, ARGB(200, 200, 200, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Raza", True)
    ' Barra Negra
    '  Call Draw_Texture_Graphic_Gui(8, X - 55, Y + 2, To_Depth(5), 167, 21, 0, 0, 167, 21, ARGB(255, 255, 255, 100), 0, eTechnique.t_Alpha)
    ' Button «
    ' Call Draw_Texture_Graphic_Gui(34, X - 60, Y + 2, To_Depth(5), 21, 21, 0, 0, 21, 21, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    ' Call Draw_Text(eFonts.f_Tahoma, 22, X - 56, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "«", True)
    ' Button »
    ' Call Draw_Texture_Graphic_Gui(34, X + 90, Y + 2, To_Depth(5), 21, 21, 0, 0, 21, 21, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    '  Call Draw_Text(eFonts.f_Tahoma, 22, X + 94, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "»", True)
    
    If UserRaza > 0 Then
        Call Draw_Text(eFonts.f_Morpheus, 22, X + 24, Y, To_Depth(7), 0, ARGB(230, 250, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, ListaRazas(UserRaza), True)

    End If
    
    ' Cambio de Cabeza para las Cuenta Premium
    ' If Account.Premium > 0 Then
   
    #If ModoBig = 0 Then
        X = 610
        Y = 245
    #Else
        X = 610
        Y = 245
    #End If
    
    '  Call Draw_Text(eFonts.f_Verdana, 15, X - 100, Y + 3, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Cabeza (VIP)", True)
        
    ' Button «
    '  Call Draw_Texture_Graphic_Gui(34, X, Y + 2, To_Depth(5), 16, 16, 0, 0, 16, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    '  Call Draw_Text(eFonts.f_Tahoma, 18, X + 3, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "«", True)
    ' Button »
    '  Call Draw_Texture_Graphic_Gui(34, X + 43, Y + 2, To_Depth(5), 16, 16, 0, 0, 16, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    '   Call Draw_Text(eFonts.f_Tahoma, 18, X + 46, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "»", True)
  
    #If ModoBig = 0 Then
        X = 370
        Y = 440
    #Else
        X = 370
        Y = 440
    #End If

    ' Botones para cambiar el Heading (Lado del Personaje)
    ' Button «
    'Call Draw_Texture_Graphic_Gui(34, X, Y + 2, To_Depth(5), 16, 16, 0, 0, 16, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    'Call Draw_Text(eFonts.f_Tahoma, 18, X + 3, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "«", True)
    ' Button »
    '  Call Draw_Texture_Graphic_Gui(34, X + 43, Y + 2, To_Depth(5), 16, 16, 0, 0, 16, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    '  Call Draw_Text(eFonts.f_Tahoma, 18, X + 46, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "»", True)
  
    'Call Draw_Texture_Graphic_Gui(35 + UserClase, 600, 250, To_Depth(2), 167, 210, 0, 0, 167, 210, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
    
    '  End If
    
    If Account.RenderHeading Then
        ' Body/Head del personaje

        #If ModoBig = 0 Then
            X = 570
            Y = 200
        #Else
            X = 1370
            Y = 360
        #End If
        
        If UserHead > 0 Then
            ' Call Draw_Grh(BodyData(UserBody).Walk(Account.RenderHeading), X, Y, To_Depth(5), 1, 1, 0)
            Call Draw_Grh(HeadData(UserHead).Head(Account.RenderHeading), X + BodyData(UserBody).HeadOffset.X, Y + BodyData(UserBody).HeadOffset.Y, To_Depth(5), 1, 0)

        End If
        
    End If
    
    ' Nombre del Personaje
              
    #If ModoBig = 0 Then
        X = 403
        Y = 452
    #Else
        X = 403
        Y = 452
    #End If

    #If ModoBig = 0 Then
        X = 765
        Y = 265
    #Else
        X = 1820
        Y = 480
    #End If

    ' Atributos segun la RAZA
    Call Draw_Text(eFonts.f_Verdana, 15, X, Y, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, ModRaza(UserRaza).Fuerza, True)
    Call Draw_Text(eFonts.f_Verdana, 15, X, Y + 55, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, ModRaza(UserRaza).Agilidad, True)
    
    Call Draw_Text(eFonts.f_Verdana, 15, X, Y + 55 + 50, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, ModRaza(UserRaza).Constitucion, True)
    Call Draw_Text(eFonts.f_Verdana, 15, X, Y + 55 + 55 + 48, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, ModRaza(UserRaza).Inteligencia, True)
     
    ' Vida y Mana Nivel Máximo
    Call Draw_Text(eFonts.f_Verdana, 15, X, Y + 55 + 55 + 55 + 42, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, getVidaIdeal(47, UserClase, ModRaza(UserRaza).Constitucion), True)

    Dim Man As Integer

    Man = Balance_AumentoMANA(UserClase, UserRaza)
     
    Call Draw_Text(eFonts.f_Verdana, 15, X, Y + 55 + 55 + 55 + 55 + 40, To_Depth(7), 0, ARGB(230, 193, 50, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, IIf(Man, Format$(Man, "##,##"), "0"), True)

End Sub

#If ModoBig > 0 Then

    Private Sub Render_Account()

            '<EhHeader>
            On Error GoTo Render_Account_Err

            '</EhHeader>
    
            Dim A          As Long

            Dim ColourText As Long

            Dim X          As Long

            Dim Y          As Long

            Dim x1         As Long

            Dim y1         As Long
    
            Dim OffsetY    As Integer
    
100         X = 180
102         Y = 380
    
104         x1 = 25
106         y1 = 0

            ' Las Monedas de la Cuenta
108         Draw_Text f_Tahoma, 19, 140, 95, To_Depth(3), 0, ARGB(250, 250, 100, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, IIf(Account.Gld > 0, Format$(Account.Gld, "##,##") & "  ORO", "0"), False, True
110         Draw_Text f_Tahoma, 19, 140, 140, To_Depth(3), 0, ARGB(250, 250, 100, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, IIf(Account.Eldhir > 0, Format$(Account.Eldhir, "##,##") & " DSP", "0"), False, True
112         Draw_Text f_Tahoma, 19, 140, 185, To_Depth(3), 0, ARGB(250, 250, 100, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, IIf(UserPoints > 0, Format$(UserPoints, "##,##") & " P. Torneo", "0"), False, True
     
            Dim Mult As Single
     
            #If ModoBig = 0 Then
114             Mult = 2.3
            #Else
116             Mult = 1
            #End If
     
            ' Hasta 10 personajes dibujados
118         For A = 1 To 10

120             With Account.Chars(A)

122                 If .PosMap > 0 Then
                        ' Recuadro del Personaje
124                     Call Draw_Texture_Graphic_Gui(25, x1 + X - 80, y1 + Y - 70, To_Depth(4), 339 / Mult, 272 / Mult, 0, 0, 339 / Mult, 272 / Mult, -1, 0, eTechnique.t_Alpha)

126                     If .Blocked > 0 Then
128                         Call Draw_Texture_Graphic_Gui(27, X - 25, Y - 15, To_Depth(5), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)

                        End If
                
                        ' Mapa del Personaje
                        ' ADAPTAR CLASICO
130                     Call Draw_Texture_Graphic_Gui(62, X + 50 + 80, Y + 75, To_Depth(5), 32, 32, 0, 0, 32, 32, -1, 0, eTechnique.t_Alpha)
132                     Draw_Text f_Tahoma, 14, X + 66 + 80, Y + 10 + 75, To_Depth(7), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .PosMap, True, True
                
                        ' Recuadro del Personaje seleccionado
134                     If Account.SelectedChar = Account.Chars(A).ID Then
136                         Call Draw_Texture_Graphic_Gui(113, x1 + X - 80, y1 + Y - 70, To_Depth(5), 339 / Mult, 272 / Mult, 0, 0, 339 / Mult, 272 / Mult, -1, 0, eTechnique.t_Alpha)

                        End If
                
138                     ColourText = ARGB(255, 255, 255, 255)
                
140                     If .Guild <> vbNullString Then
142                         Draw_Text f_Morpheus, 18, X + 15, Y + 60, To_Depth(7), 0, .Colour, FONT_ALIGNMENT_CENTER, "<" & .Guild & ">", True, True

                        End If
                
                        ' Cuadrado Style Mu
144                     Draw_Text f_Booter, 30, X + 117, y1 + Y - 26, To_Depth(7), 0, ColourText, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, ListaClases(.Class) & " " & ListaRazas(.Raze), True, True
146                     Draw_Text f_Booter, 30, X + 117, y1 + Y - 6, To_Depth(7), 0, ColourText, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Nivel " & .Elv, True, True
148                     Draw_Text f_Booter, 25, X + 110, Y + 150, To_Depth(7), 0, .Colour, FONT_ALIGNMENT_CENTER, .Name, True, True
                
                        Dim AlphaBody As Long
                
150                     If .Body <> FRAGATA_FANTASMAL And .Body <> iCuerpoMuerto And .Body <> iCuerpoMuerto_Legion Then
152                         AlphaBody = ARGB(255, 255, 255, 255)
                    
154                         If .Head > 0 Then Call Draw_Grh(HeadData(.Head).Head(E_Heading.SOUTH), X + BodyData(.Body).HeadOffset.X + 80, Y + BodyData(.Body).HeadOffset.Y + 75, To_Depth(7), 1, 0, , , , , GrhData(HeadData(.Head).Head(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(HeadData(.Head).Head(E_Heading.SOUTH).GrhIndex).pixelHeight)
                
156                         If .Helm > 0 And .Helm <> 2 Then
158                             If .Helm = 54 Then
160                                 OffsetY = 13
                                Else
162                                 OffsetY = 0

                                End If
                        
164                             Call Draw_Grh(CascoAnimData(.Helm).Head(E_Heading.SOUTH), X + BodyData(.Body).HeadOffset.X + 80, Y + BodyData(.Body).HeadOffset.Y + OffsetY + 75, To_Depth(7, , , 3), 1, 0, , , , , GrhData(CascoAnimData(.Helm).Head(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(CascoAnimData(.Helm).Head(E_Heading.SOUTH).GrhIndex).pixelHeight)

                            End If
                
166                         If .Shield > 0 And .Shield <> 2 Then Call Draw_Grh(ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH), X + 80, Y + 75, To_Depth(7, , , 4), 1, 1, 0, , , , (GrhData(ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH).GrhIndex).pixelWidth / 2), (GrhData(ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH).GrhIndex).pixelHeight / 2))
                            
168                         If .Weapon > 0 And .Weapon <> 2 Then Call Draw_Grh(WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH), X + 80, Y + 75, To_Depth(8), 1, 1, 0, , , , GrhData(WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH).GrhIndex).pixelWidth / 2, GrhData(WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH).GrhIndex).pixelHeight / 2)
                        Else
170                         AlphaBody = ARGB(255, 255, 255, 150)

                        End If
                
172                     Call Draw_Grh(BodyData(.Body).Walk(E_Heading.SOUTH), X + 80, Y + 75, To_Depth(6), 1, 0, 0, AlphaBody, , eTechnique.t_Alpha, GrhData(BodyData(.Body).Walk(E_Heading.SOUTH).GrhIndex).pixelWidth / 2, GrhData(BodyData(.Body).Walk(E_Heading.SOUTH).GrhIndex).pixelHeight / 2)
            
                        'Else
                        '   Draw_Text f_Medieval, 18, X + 20, Y + 10, To_Depth(7), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "(CREAR PJ)", True, True
            
                    End If
            
174                 X = X + 340
            
176                 If A = 5 Then
                
178                     Y = Y + 280
180                     X = 180

                    End If

                End With

182         Next A

            '<EhFooter>
            Exit Sub

Render_Account_Err:
            LogError err.Description & vbCrLf & "in ARGENTUM.Engine_Renders.Render_Account " & "at line " & Erl

            Resume Next

            '</EhFooter>
    End Sub


#Else

Private Sub Render_Account()
        '<EhHeader>
        On Error GoTo Render_Account_Err
        '</EhHeader>
    
        Dim A          As Long
        Dim ColourText As Long
        Dim X          As Long
        Dim Y          As Long
        Dim x1         As Long
        Dim y1         As Long
    
        Dim OffsetY    As Integer
    
100     X = 90
102     Y = 240
    
104     x1 = 25
106     y1 = 0

        ' Las Monedas de la Cuenta
108     Draw_Text f_Tahoma, 15, 60, 50, To_Depth(3), 0, ARGB(250, 250, 100, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, IIf(Account.Gld > 0, Format$(Account.Gld, "##,##") & " ORO", "0"), False, True
110     Draw_Text f_Tahoma, 15, 60, 75, To_Depth(3), 0, ARGB(250, 250, 100, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, IIf(Account.Eldhir > 0, Format$(Account.Eldhir, "##,##") & " DSP", "0"), False, True
112     Draw_Text f_Tahoma, 15, 60, 103, To_Depth(3), 0, ARGB(250, 250, 100, 175), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, IIf(UserPoints > 0, Format$(UserPoints, "##,##") & " P. Torneo", "0"), False, True
    
     
         Dim Mult As Single
     
         #If ModoBig = 0 Then
114         Mult = 2.3
         #Else
116         Mult = 1
         #End If
     
         ' Hasta 10 personajes dibujados
118     For A = 1 To 10
120         With Account.Chars(A)
122             If .PosMap > 0 Then
                    ' Recuadro del Personaje
124                 Call Draw_Texture_Graphic_Gui(25, x1 + X - 80, y1 + Y - 70, To_Depth(4), 339 / Mult, 272 / Mult, 0, 0, 339 / Mult, 272 / Mult, -1, 0, eTechnique.t_Alpha)

126                 If .Blocked > 0 Then
128                     Call Draw_Texture_Graphic_Gui(27, X - 25, Y - 15, To_Depth(5), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
                    End If
                
                    ' Mapa del Personaje
                    ' ADAPTAR CLASICO
130                 Call Draw_Texture_Graphic_Gui(62, X + 30, Y, To_Depth(5), 32, 32, 0, 0, 32, 32, -1, 0, eTechnique.t_Alpha)
132                 Draw_Text f_Tahoma, 14, X + 46, Y + 10, To_Depth(7), 0, -1, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .PosMap, True, True
                
                     ' Recuadro del Personaje seleccionado
134                 If Account.SelectedChar = Account.Chars(A).ID Then
136                     Call Draw_Texture_Graphic_Gui(113, x1 + X - 80, y1 + Y - 70, To_Depth(5), 339 / Mult, 272 / Mult, 0, 0, 339 / Mult, 272 / Mult, -1, 0, eTechnique.t_Alpha)
                    End If
                
138                 ColourText = ARGB(255, 255, 255, 255)
                
                   ' If .Guild <> vbNullString Then
                   '     Draw_Text f_Morpheus, 18, X + 15, Y + 60, To_Depth(7), 0, .Colour, FONT_ALIGNMENT_CENTER, "<" & .Guild & ">", True, True
                  '  End If
                
                    ' Cuadrado Style Mu
140                 Draw_Text f_Tahoma, 14, X + 17, y1 + Y - 52, To_Depth(7), 0, ColourText, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, ListaClases(.Class) & " " & ListaRazas(.Raze), True, True
142                 Draw_Text f_Tahoma, 14, X + 17, y1 + Y - 42, To_Depth(7), 0, ColourText, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Nivel " & .Elv, True, True
                    'Draw_Text f_Tahoma, 14, X + 17, y1 + Y - 42, To_Depth(7), 0, ColourText, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Lvl" & .Elv, True, True
144                 Draw_Text f_Booter, 20, X + 15, Y + 42, To_Depth(7), 0, .Colour, FONT_ALIGNMENT_CENTER, .Name, True, True
                
                    Dim AlphaBody As Long
                
146                 If .Body <> FRAGATA_FANTASMAL And .Body <> iCuerpoMuerto And .Body <> iCuerpoMuerto_Legion Then
148                     AlphaBody = ARGB(255, 255, 255, 255)
                    
150                     If .Head > 0 Then _
                           Call Draw_Grh(HeadData(.Head).Head(E_Heading.SOUTH), X + BodyData(.Body).HeadOffset.X, Y + BodyData(.Body).HeadOffset.Y, To_Depth(6), 1, 0, , , , , _
                           GrhData(HeadData(.Head).Head(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(HeadData(.Head).Head(E_Heading.SOUTH).GrhIndex).pixelHeight)
                
152                     If .Helm <> 2 Then
154                         If .Helm = 54 Then
156                             OffsetY = 13
                            Else
158                             OffsetY = 0
                            End If
                        
160                         Call Draw_Grh(CascoAnimData(.Helm).Head(E_Heading.SOUTH), X + BodyData(.Body).HeadOffset.X, Y + BodyData(.Body).HeadOffset.Y + OffsetY, To_Depth(6, , , 3), 1, 0, , , , , _
                               GrhData(CascoAnimData(.Helm).Head(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(CascoAnimData(.Helm).Head(E_Heading.SOUTH).GrhIndex).pixelHeight)
                        End If
                
162                     If ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH).GrhIndex Then _
                           Call Draw_Grh(ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH), X, Y, To_Depth(7, , , 4), 1, 1, 0, , , , _
                           GrhData(ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(ShieldAnimData(.Shield).ShieldWalk(E_Heading.SOUTH).GrhIndex).pixelHeight)
                            
164                     If WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH).GrhIndex Then _
                           Call Draw_Grh(WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH), X, Y, To_Depth(7), 1, 1, 0, , , , _
                           GrhData(WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(WeaponAnimData(.Weapon).WeaponWalk(E_Heading.SOUTH).GrhIndex).pixelHeight)
                    Else
166                     AlphaBody = ARGB(255, 255, 255, 150)
                    End If
                
168                 Call Draw_Grh(BodyData(.Body).Walk(E_Heading.SOUTH), X, Y, To_Depth(6), 1, 0, 0, AlphaBody, , eTechnique.t_Alpha, _
                       GrhData(BodyData(.Body).Walk(E_Heading.SOUTH).GrhIndex).pixelWidth, GrhData(BodyData(.Body).Walk(E_Heading.SOUTH).GrhIndex).pixelHeight)
                
                
                            

            
                End If
170             X = X + 147
            
172             If A = 5 Then
                
174                 Y = Y + 140
176                 X = 86
                End If
            End With
178     Next A
        '<EhFooter>
        Exit Sub

Render_Account_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Engine_Renders.Render_Account " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

#End If


' Dibuja la Quest del NPCINDEX
' # Renderizamos el PANEL DE MISIONES
Public Sub Render_QuestPanel()

    Dim Pos     As Integer
    Dim line    As String
    Dim Color   As Long
    Dim PosX    As Integer
    Dim PosY    As Integer
    
    Dim A       As Long, b As Long
    Dim Colour  As Long
    Dim OffsetY As Integer
    
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cPivQuest))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, frmCriatura_Quest.ScaleWidth, frmCriatura_Quest.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    Dim X     As Integer
    Dim Y     As Integer
    
    Dim DescY As Integer
    
    ' Textura de Fondo
    '
    Call Draw_Texture_Graphic_Gui(56, 0, 0, To_Depth(1), 518, 460, 0, 0, 518, 460, ARGB(200, 200, 200, 200), 0, eTechnique.t_Default)
    
    ' Mapa de Fondo
    '
    
    Dim X_Base As Integer
    Dim Y_Base As Integer
    

    
    Dim InitialY_List As Integer
    Dim Alpha         As Byte
    InitialY_List = 128
    
    For A = 1 To QuestLast
        With QuestList(QuestNpc(A))
            
            If A = QuestIndex Then
                Alpha = 255
            Else
                Alpha = 175
            End If
            
            ' Grafico de la Barra de Info
            Call Draw_Texture_Graphic_Gui(1, 56, 6 + Pos + InitialY_List, To_Depth(2), 400, 24, 0, 0, 510, 24, ARGB(255, 255, 255, Alpha), 0, eTechnique.t_Default)
            
            ' Fondo & Recuadro del [ITEM]
            Call Draw_Texture_Graphic_Gui(3, 56, Pos + InitialY_List, To_Depth(4), 32, 32, 0, 0, 32, 32, ARGB(255, 255, 255, Alpha), 0, eTechnique.t_Default)
            
            PosX = 56
            PosY = PosY + Pos
            Call Draw_Texture(ObjData(QuestList(QuestNpc(A)).RewardObjs(1).ObjIndex).GrhIndex, PosX, ((A - 1) * 32) + InitialY_List, To_Depth(4), 32, 32, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
            
            If QuestList(QuestNpc(A)).RewardObjs(1).Amount > 1 Then
            Call Draw_Text(eFonts.f_Tahoma, 12, PosX, Pos + InitialY_List, To_Depth(8), 0, ARGB(255, 240, 240, Alpha), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "x" & QuestList(QuestNpc(A)).RewardObjs(1).Amount, True, True)
            End If
            Call Draw_Texture(6579, PosX, ((A - 1) * 32) + InitialY_List, To_Depth(5), 32, 32, ARGB(255, 255, 255, Alpha), 0, eTechnique.t_Default)
            Call Draw_Text(eFonts.f_Medieval, 20, PosX + 40, Pos + 10 + InitialY_List, To_Depth(5), 0, ARGB(255, 200, 0, Alpha), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, ObjData(QuestList(QuestNpc(A)).RewardObjs(1).ObjIndex).Name, True, True)

        
            Dim XBaseObj As Long
            
            XBaseObj = 215
            If .Obj > 0 Then
          
             'Call Draw_Texture_Graphic_Gui(7, 235, InitialY_List + 14, To_Depth(5), 2, 152, 0, 0, 2, 152, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
            ' Draw_Texture_Graphic_Gui(7, 235, InitialY_List + 95, To_Depth(5), 2, 152, 0, 0, 2, 152, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
                For b = 1 To 5
                    
                   ' Call Draw_Texture(30652, XBaseObj + PosX + (b * 32), ((A - 1) * 32) + InitialY_List, To_Depth(3), 32, 32, ARGB(255, 255, 255, Alpha), 0, eTechnique.t_Default)
                    'Call Draw_Texture(6579, XBaseObj + PosX + (b * 32), ((A - 1) * 32) + InitialY_List, To_Depth(4), 32, 32, ARGB(0, 150, 200, Alpha), 0, eTechnique.t_Default)
                    Call Draw_Texture_Graphic_Gui(3, XBaseObj + PosX + (b * 32), ((A - 1) * 32) + InitialY_List, To_Depth(3), 32, 32, 0, 0, 32, 32, ARGB(255, 255, 255, Alpha), 0, eTechnique.t_Default)
                    
                    If b <= .Obj Then
                    Call Draw_Texture(ObjData(.Objs(b).ObjIndex).GrhIndex, XBaseObj + PosX + (b * 32), ((A - 1) * 32) + InitialY_List, To_Depth(4), 32, 32, ARGB(255, 255, 255, Alpha), 0, eTechnique.t_Default)
                    
                        If .Objs(b).Amount > 1 Then
                        Call Draw_Text(eFonts.f_Tahoma, 12, XBaseObj + PosX + (b * 32), ((A - 1) * 32) + InitialY_List, To_Depth(8), 0, ARGB(255, 240, 240, Alpha), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "x" & .Objs(b).Amount, True, True)
                        End If
                    End If
                    'XBaseObj = XBaseObj - (b * 32)
                
                Next b
            End If
        End With
        
        Pos = (A * 32)
    Next A
    
    ' Selección
    If QuestIndex > 0 Then
        With QuestList(QuestNpc(QuestIndex))
           ' If QuestObjIndex > 0 Then
               ' If .Obj > 0 Then
                     'Call Draw_Text(eFonts.f_Medieval, 20, 220, 22, To_Depth(8), 0, ARGB(180, 180, 180, 140), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, ObjData(.Objs(QuestObjIndex).ObjIndex).Name, True)
                'End If
           ' End If
            
            If QuestNpcIndex > 0 Then
            
                If .Npc > 0 Then
                    Call Draw_Text(eFonts.f_Medieval, 19, 105, 380, To_Depth(8), 0, ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, NpcList(.Npcs(QuestNpcIndex).NpcIndex).Name & " (x" & .Npcs(QuestNpcIndex).Amount & ")", True, True)
                    Call Draw_Grh(BodyData(NpcList(.Npcs(QuestNpcIndex).NpcIndex).Body).Walk(E_Heading.EAST), 85, 420, To_Depth(5), 1, 0, 0)
                        
                    If NpcList(.Npcs(QuestNpcIndex).NpcIndex).Head > 0 Then
                        Call Draw_Grh(HeadData(NpcList(.Npcs(QuestNpcIndex).NpcIndex).Head).Head(E_Heading.EAST), 85 + BodyData(NpcList(.Npcs(QuestNpcIndex).NpcIndex).Body).HeadOffset.X, _
                           420 + BodyData(NpcList(.Npcs(QuestNpcIndex).NpcIndex).Body).HeadOffset.Y, To_Depth(5), 1, 0, 0)
                    End If
            
                    If .Npc <> 1 Then
                        Call Draw_Text(eFonts.f_Tahoma, 20, 222, 430, To_Depth(8), 0, ARGB(255, 240, 240, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "»", True, True)
                        Call Draw_Text(eFonts.f_Tahoma, 20, 200, 430, To_Depth(8), 0, ARGB(255, 240, 240, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "«", True, True)
                    End If
                End If
                
            Else
                If UBound(.Desc) <> -1 Then
                    For b = LBound(.Desc) To UBound(.Desc)
                        Call Draw_Text(eFonts.f_Verdana, 14, 250, 75 + DescY, To_Depth(8), 0, ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Desc(b), True, True)
                        DescY = DescY + 16
                    Next b
                End If
            End If
        End With
    End If
    Call wGL_Graphic_Renderer.Flush
End Sub

Public Sub Render_CriaturaInfo()

    Dim Y_Avance As Long
    Dim A As Long
    Dim b As Long
    
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cCriaturaInfo))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmCriatura_Info.ScaleWidth, FrmCriatura_Info.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    Call Draw_Texture_Graphic_Gui(6, 0, 0, To_Depth(1), 220, 350, 0, 0, 220, 350, ARGB(255, 255, 255, 200), 0, eTechnique.t_Default)
    
    With NpcList(SelectedNpcIndex)
        
        ' Cuerpo de la Criatura
        Call Draw_Grh(BodyData(.Body).Walk(E_Heading.EAST), 85, 200, To_Depth(5), 1, 0, 0)
        
        If .Head > 0 Then
            Call Draw_Grh(HeadData(.Head).Head(E_Heading.EAST), 85 + BodyData(.Body).HeadOffset.X, _
                200 + BodyData(.Body).HeadOffset.Y, To_Depth(5), 1, 0, 0)
        End If
        
        Y_Avance = 20
        
        ' Nombre de la Criatura
        Call Draw_Text(eFonts.f_Medieval, 20, 105, Y_Avance, To_Depth(9), 0, ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Name, True)
        Y_Avance = Y_Avance + 20
        
        
        ' Vida de la Criatura
        If .MaxHit > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Hp", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.MaxHp, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
        
        ' Hit de la Criatura
        If .MinHit > 0 Then
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Hit", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.MinHit, "#,###") & "/" & Format$(.MaxHit, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
        
        ' Defensa de la Criatura
        If .Def > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Def", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.Def, "#,###"), True)
            Y_Avance = Y_Avance + 20
        End If
        
        ' Defensa Magica de la Criatura
        If .DefM > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Def Mag", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.DefM, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
        
        ' Evasion
        If .PoderEvasion > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Evasion", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.PoderEvasion, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
        
        ' Poder de Ataque
        If .PoderAtaque > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Evasion", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.PoderAtaque, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
        
        ' Experiencia
        If .GiveExp > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Exp", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(50, 200, 200, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.GiveExp, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
           
            ' Oro
        If .GiveGld > 0 Then
            Call Draw_Text(eFonts.f_Morpheus, 20, 75, Y_Avance, To_Depth(9), 0, _
               ARGB(200, 200, 175, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Oro", True)
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 140, Y_Avance, To_Depth(9), 0, _
               ARGB(50, 200, 200, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, Format$(.GiveGld, "#,###"), True)
            
            Y_Avance = Y_Avance + 20
        End If
            
        Dim Y_ITEM As Integer
        Dim X_ITEM As Integer
            
        X_ITEM = 14
        Y_ITEM = 240
        
        ' Inventario del NPC
        If .NroItems > 0 Then
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 105, 220, To_Depth(9), 0, _
               ARGB(255, 200, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Inv", True)
               
            For A = 1 To 6
                 ' Fondo & Recuadro del ITEM
                'Call Draw_Texture_Graphic_Gui(2, X_ITEM, Y_ITEM, To_Depth(2), 32, 32, 0, 0, 32, 32, -1, 0, eTechnique.t_Default)
                Call Draw_Texture_Graphic_Gui(3, X_ITEM, Y_ITEM, To_Depth(3), 32, 32, 0, 0, 32, 32, -1, 0, eTechnique.t_Default)
                Call Draw_Texture(6579, X_ITEM, Y_ITEM, To_Depth(4), GrhData(6579).pixelWidth, GrhData(6579).pixelHeight, -1, 0, eTechnique.t_Default)
                
                If .Object(A).ObjIndex > 0 Then
                     Call Draw_Text(eFonts.f_Tahoma, 10, X_ITEM, Y_ITEM, To_Depth(9), 0, _
                        ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "x" & .Object(A).Amount, True)
                    Call Draw_Texture(ObjData(.Object(A).ObjIndex).GrhIndex, X_ITEM, Y_ITEM, To_Depth(4), 32, 32, -1, 0, eTechnique.t_Default)
                End If
                
                X_ITEM = X_ITEM + 32
                'If A Mod 7 = 0 Then Y_ITEM = Y_ITEM + 32
            Next A
        
        End If
        
        X_ITEM = 14
        Y_ITEM = 292
        
         ' Drop del NPC
        If .NroDrops > 0 Then
        
            Call Draw_Text(eFonts.f_Morpheus, 20, 105, 272, To_Depth(9), 0, _
               ARGB(40, 150, 40, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Drop", True)
               
            For A = 1 To 6
                 ' Fondo & Recuadro del ITEM
                Call Draw_Texture_Graphic_Gui(2, X_ITEM, Y_ITEM, To_Depth(2), 32, 32, 0, 0, 32, 32, -1, 0, eTechnique.t_Default)
                Call Draw_Texture_Graphic_Gui(3, X_ITEM, Y_ITEM, To_Depth(3), 32, 32, 0, 0, 32, 32, -1, 0, eTechnique.t_Default)
                Call Draw_Texture(6579, X_ITEM, Y_ITEM, To_Depth(4), GrhData(6579).pixelWidth, GrhData(6579).pixelHeight, ARGB(50, 200, 200, 255), 0, eTechnique.t_Alpha)
                
                If .Drop(A).ObjIndex > 0 Then
                    Call Draw_Text(eFonts.f_Tahoma, 10, X_ITEM, Y_ITEM, To_Depth(9), 0, _
                        ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, Porcentaje_String(.Drop(A).Probability), True)
                     Call Draw_Text(eFonts.f_Tahoma, 10, X_ITEM, Y_ITEM + 22, To_Depth(9), 0, _
                        ARGB(255, 255, 255, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "x" & .Drop(A).Amount, True)
                    
                    
                    
                    Call Draw_Texture(ObjData(.Drop(A).ObjIndex).GrhIndex, X_ITEM, Y_ITEM, To_Depth(4), 32, 32, -1, 0, eTechnique.t_Default)
                    
                    
                End If
                
                X_ITEM = X_ITEM + 32
                'If A Mod 7 = 0 Then Y_ITEM = Y_ITEM + 32
            Next A
        
        End If
        
    End With
    
    Call wGL_Graphic_Renderer.Flush
End Sub


Public Sub RenderScreen_MiniMapa()

    '<EhHeader>
    On Error GoTo RenderScreen_Err
    '</EhHeader>
        
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cMiniMapa))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmMain.MiniMapa.ScaleWidth, FrmMain.MiniMapa.ScaleHeight)
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
    Dim Mult As Byte
    
    #If ModoBig = 1 Then
    Mult = 2
    #Else
    Mult = 1
    #End If
    
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
        DrawableY = (Y - ScreenMinY) * Mult
    
        For X = 1 To 100
            DrawableX = (X - ScreenMinX) * Mult
            With MapData(X, Y)
                If (.Graphic(1).GrhIndex <> 0) Then
                    Call Draw_Grh(.Graphic(1), DrawableX, DrawableY, To_Depth(1, X, Y), 0, 1, , , , , Mult, Mult)
                End If
                    
                 If (.Graphic(2).GrhIndex <> 0) Then
                    Call Draw_Grh(.Graphic(2), DrawableX, DrawableY, To_Depth(2, X, Y), 0, 1, , , , , Mult, Mult)
                End If
                    
                If (.Graphic(3).GrhIndex <> 0) Then
                 Call Draw_Grh(.Graphic(3), DrawableX, DrawableY, To_Depth(3, X, Y), 0, 1, , , , , Mult, Mult)
                End If
                    
                               If (.Graphic(4).GrhIndex <> 0) Then
                 Call Draw_Grh(.Graphic(4), DrawableX, DrawableY, To_Depth(4, X, Y), 0, 1, , , , , Mult, Mult)
                End If
            End With
        Next X
    Next Y
    
    
    
    If Not RenderizandoMap Then
        ' Miembros del clan
        Dim A As Long
    
        For A = 1 To MAX_GUILD_MEMBER
            With MiniMap_Friends(A)
                If .X > 0 And .Y > 0 Then
                    Call Draw_Texture_Graphic_Gui(61, .X * Mult, .Y * Mult, To_Depth(5), 8, 8, 0, 0, 8, 8, ARGB(255, 6, 6, 255), 0, eTechnique.t_Alpha)
                End If
            End With
        Next A
        
        Call Draw_Texture_Graphic_Gui(61, UserPos.X * Mult, UserPos.Y * Mult, To_Depth(5), 8, 8, 0, 0, 8, 8, ARGB(200, 6, 6, 255), 0, eTechnique.t_Alpha)
    End If
    
    Call wGL_Graphic_Renderer.Flush
    '<EhFooter>
    Exit Sub

RenderScreen_Err:
    LogError err.Description & vbCrLf & _
       "in RenderScreen_MiniMapa " & _
       "at line " & Erl

    '</EhFooter>
End Sub
Public Sub RenderScreen_MiniMapa_PNG(ByVal UserMap As Integer)

    '<EhHeader>
    On Error GoTo RenderScreen_Err
    '</EhHeader>
        
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cMiniMapa))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmMain.MiniMapa.ScaleWidth, FrmMain.MiniMapa.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
   
    Dim Width As Integer
    Dim Height As Integer
    Dim Mult As Byte
    
    #If ModoBig = 1 Then
        Width = 200
        Height = 200
        Mult = 2
    #Else
                Width = 100
        Height = 100
        Mult = 1
    #End If
    
    Call Draw_Texture_Graphic_MiniMap(UserMap, 0, 0, To_Depth(1), Width, Height, 0, 0, Width, Height, ARGB(255, 255, 255, 255), 0, eTechnique.t_Alpha)
     
    
    ' Miembros del clan
    Dim A As Long
    
    For A = 1 To MAX_GUILD_MEMBER
        With MiniMap_Friends(A)
            If .X > 0 And .Y > 0 Then
                Call Draw_Texture_Graphic_Gui(61, .X * Mult, .Y * Mult, To_Depth(5), 8, 8, 0, 0, 8, 8, ARGB(255, 6, 6, 255), 0, eTechnique.t_Alpha)
            End If
        End With
    Next A
    
    Call Draw_Texture_Graphic_Gui(61, UserPos.X * Mult, UserPos.Y * Mult, To_Depth(5), 8, 8, 0, 0, 8, 8, ARGB(200, 6, 6, 255), 0, eTechnique.t_Alpha)
    Call wGL_Graphic_Renderer.Flush
    '<EhFooter>
    Exit Sub

RenderScreen_Err:
    LogError err.Description & vbCrLf & _
       "in RenderScreen_MiniMapa " & _
       "at line " & Erl

    '</EhFooter>
End Sub


Public Function ComprobarPosibleMacro(ByVal MouseX As Integer, ByVal MouseY As Integer) As Boolean
    Call CopyMemory(ContadorMacroClicks(2), ContadorMacroClicks(1), Len(ContadorMacroClicks(1)) * (MAX_COMPROBACIONES - 1))
    
    ContadorMacroClicks(1).X = MouseX
    ContadorMacroClicks(1).Y = MouseY
    
    Dim I As Byte
    
    For I = 1 To MAX_COMPROBACIONES
        If ContadorMacroClicks(I).X <> MouseX Or ContadorMacroClicks(I).Y <> MouseY Then
            ComprobarPosibleMacro = False
            Exit Function
        End If
    Next I
    
    
    ComprobarPosibleMacro = True
    Call WriteDenounce("[SEGURIDAD]: El abogado para averiguar datos no sale mucho")
End Function


Public Sub CountPacketIterations(ByRef packetControl As t_packetControl, ByVal expectedAverage As Double)
        '<EhHeader>
        On Error GoTo CountPacketIterations_Err
        '</EhHeader>

        Dim delta As Long
        Dim actualcount As Long
    
100     actualcount = FrameTime
    
102     delta = actualcount - packetControl.last_count
    
104     If delta < 40 Then Exit Sub
    
106     packetControl.last_count = actualcount
    
108     Call alterIndex(packetControl)
    
110     packetControl.iterations(10) = delta
        Dim percentageDiff As Double, average As Double
112     percentageDiff = getPercentageDiff(packetControl)
114     average = getAverage(packetControl)
       ' Debug.Print "Delta: " & delta & " Average: " & average
116     If percentageDiff < 5 Then
            'Debug.Print "DIFF: " & getPercentageDiff(packetControl)
            'Call AddtoRichTextBox(frmMain.RecTxt, "DIFF: " & getPercentageDiff(packetControl), 255, 200, 0, True)
118         Call WriteDenounce("[SEGURIDAD]: Está repitiendo un Macro" & getPercentageDiff(packetControl))
            'Debug.Print "DIFF: " & getPercentageDiff(packetControl)
        End If
    
120     If average > 50 And average < expectedAverage Then
122          Call WriteDenounce("[SEGURIDAD]: Está repitiendo un Macro" & getPercentageDiff(packetControl))
        End If
    
        '<EhFooter>
        Exit Sub

CountPacketIterations_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.Engine_Renders.LaEscaloneta " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Function getPercentageDiff(ByRef packetControl As t_packetControl) As Double

    Dim I As Long, min As Long, max As Long
    
    min = packetControl.iterations(1)
    max = packetControl.iterations(1)
    
    Dim Count As Long
    
    For I = 1 To 10
    
        If packetControl.iterations(I) < min Then
            min = packetControl.iterations(I)
        End If
        
        If packetControl.iterations(I) > max Then
            max = packetControl.iterations(I)
        End If
        
    Next I
    
    getPercentageDiff = 100 - ((min * 100) / max)
    
End Function

Private Function getAverage(ByRef packetControl As t_packetControl) As Double

    Dim I As Long, suma As Long
    
    For I = 1 To UBound(packetControl.iterations)
        suma = suma + packetControl.iterations(I)
    Next I
    
    getAverage = suma / UBound(packetControl.iterations)
    
End Function

Private Sub alterIndex(ByRef packetControl As t_packetControl)
    Dim I As Long
    
    For I = 1 To 10 ' packetControl.cant_iterations
        If I < 10 Then 'packetControl.cant_iterations Then
            packetControl.iterations(I) = packetControl.iterations(I + 1)
        End If
    Next I
End Sub
